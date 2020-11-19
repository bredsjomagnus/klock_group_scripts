from __future__ import print_function
import pandas as pd
from config import *
import datetime
import os
import sys
import getopt
import pickle
from tqdm import tqdm
import os.path
from env import *
from termcolor import colored, cprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request



def error_report(error_list):
    print()
    cprint("     *ROWS MISSING IN ELEVLISTA:ELEVLISTA*", 'yellow', attrs=['bold'])
    cprint("        Missing %d at:" % (len(error_list)), 'yellow')
    for row_number in tqmd(error_list, ascii=True, desc="Error report"):
        cprint("        - row %d" % (row_number), 'yellow')

def authenticate():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('sheet_token.pickle'):
        with open('sheet_token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'sheet_credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('sheet_token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    return service

# def authenticate():
#     creds = None
#     # The file token.pickle stores the user's access and refresh tokens, and is
#     # created automatically when the authorization flow completes for the first
#     # time.
#     if os.path.exists('token.pickle'):
#         with open('token.pickle', 'rb') as token:
#             creds = pickle.load(token)
#     # If there are no (valid) credentials available, let the user log in.
#     if not creds or not creds.valid:
#         if creds and creds.expired and creds.refresh_token:
#             creds.refresh(Request())
#         else:
#             flow = InstalledAppFlow.from_client_secrets_file(
#                 'credentials.json', SCOPES)
#             creds = flow.run_local_server(port=0)
#         # Save the credentials for the next run
#         with open('token.pickle', 'wb') as token:
#             pickle.dump(creds, token)

#     service = build('sheets', 'v4', credentials=creds)

#     return service

def get_elevlista_with_emails(service, ELEVLISTA_ID):
    sheet_range = 'elevlista!A1:E'
    nan_counter = 0
    elev_counter = 0
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ELEVLISTA_ID,
                                range=sheet_range).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        # print('Läser in elevnamn_till_elevmail från Driven: ', end="")
        klasser = []
        names = []
        groups = []
        emails = []
        for i, row in enumerate(tqdm(values, ascii=True, desc="Get Elevlista with emails")):
            if i > 0:
                try: # this will take care of eventually empty cells.
                    klass = row[0]
                    name = row[1]
                    group = row[2]
                    email = row[4]

                    klasser.append(klass)
                    names.append(name)
                    groups.append(group)
                    emails.append(email)
                    elev_counter += 1
                except Exception as e:
                    nan_counter += 1

        print("%d valid rows and %d nan values" % (elev_counter, nan_counter))
        print()
        elevlista_dict = {
            'Elev Klass': klasser,
            'Elev Namn': names,
            'Elev Grupper': groups,
            'Elev Mail': emails
        }
        df_elevlista = pd.DataFrame.from_dict(elevlista_dict)
        return df_elevlista

def get_elevlista_without_mail(service, ELEVLISTA_ID):
    """
    Get Infometor (elevlista) file in Drive and return it as a Dataframe
    """
    RANGE = 'elevlista!A1:D'

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ELEVLISTA_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in elevlista.')
    else:
        error_list = []

        klasser = []
        personnummers = []
        names = []
        groups = []

        for i, row in enumerate(tqdm(values, ascii=True, desc="Get Elevlista without email")):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    klass = row[0]
                    personnummer = str(row[3])
                    name = row[1]
                    group = row[2]

                    klasser.append(klass)
                    personnummers.append(personnummer)
                    names.append(name)
                    groups.append(group)
                    # emails.append(email)
                except Exception as e:
                    error_list.append(i+1)

        elevlista_dict = {
            'Klass': klasser,
            'Namn': names,
            "Grupper": groups,
            'Personnummer': personnummers,
        }
        df_elevlista = pd.DataFrame.from_dict(elevlista_dict)
    clear_range = 'elevlista!E2'
    request = service.spreadsheets().values().clear(spreadsheetId=ELEVLISTA_ID, range=clear_range).execute()

    return df_elevlista, error_list

def get_elevlista_with_personummer_as_index(service, ELEVLISTA_ID):
    """
    Hämta hem elevlistan i förberedelse att användas mot extenslistan för
    att bygga gruppimporttabellen.

    Skillnaden här är att funktionen inte utför clear och att personnummer sätts som index.
    """
    RANGE = 'elevlista!A1:D'

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ELEVLISTA_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in elevlista.')
    else:
        error_list = []

        klasser = []
        personnummers = []
        names = []
        groups = []

        for i, row in enumerate(tqdm(values, ascii=True, desc="Get elevlista prepped for matching against Extenslista for group import")):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    klass = row[0]
                    personnummer = str(row[3])
                    name = row[1]
                    group = row[2]

                    klasser.append(klass)
                    personnummers.append(personnummer)
                    names.append(name)
                    groups.append(group)
                    # emails.append(email)
                except Exception as e:
                    error_list.append(i+1)

        _dict = {
            'Klass': klasser,
            'Namn': names,
            "Grupper": groups,
            'Personnummer': personnummers,
        }
        df = pd.DataFrame.from_dict(_dict)
        df.set_index('Personnummer', inplace=True)
    return df, error_list

def get_groupimport(service, EXTENS_ID):
    """
    Hämtar personnumer och namn från Extenslistan. Här sätts personnumret som index.
    """
    RANGE = 'extens!A1:E'

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=EXTENS_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in elevlista.')
    else:
        error_list = []

        personnummers = []
        names = []

        for i, row in enumerate(tqdm(values, ascii=True, desc="Get personid and name from Extenslist")):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    personnummer = str(row[1])
                    lastname = str(row[2])
                    firstname = str(row[3])
                    personnummers.append(personnummer)
                    names.append(lastname+", "+firstname)
                    
                except Exception as e:
                    error_list.append(str(row[1]))

        _dict = {
            'Personid': personnummers,
            'Namn': names
        }
        df = pd.DataFrame.from_dict(_dict)
        df.set_index('Personid', inplace=True)

    return df, error_list

def set_leading_zero(value):
    return '0' + value

def get_group_import_content(df_group, df_elev):
    """
    Ser till att samla upp det som skall in i gruppimporttabellen och returnera det.

    Listan från Extens och Elevlistan jämförs här. Vid en direkt match av personnummerna
    läggs grupperna till från elevlistan.

    Matchar det inte kollas om de sex första siffrona matchar samtidigt som namnet stämmer. Även 
    då förs grupperna in men med en notering vem det gäller.

    Blir det ingen match alls skriv detta ut i raderna som MISSING MATCH och användaren meddelas
    i terminalen.
    """
    content = []
    _errors = []
    for index_grp, row_grp in tqdm(df_group.iterrows(), ascii=True, desc="Get Group Import Table Content"):
        """
        index_grp är index dvs personnumret i extenslistan
        """

        # Tar bort bindestrecket från personid i df_group
        index_grp = str(index_grp).replace('-', '')
        
        index_grp = index_grp.strip()

        # print("###")
        # print("")
        # print(index_grp)
        # print("")
        # print("####")
        if len(index_grp) < 10:
            set_leading_zero(index_grp)
            print(f'set_leading_zero({index_grp})')
        
        if index_grp in df_elev.index:
            """
            Om personid finns i elevlistan
            """
            row = [index_grp, df_elev.loc[index_grp].Grupper, '']

        else:
            """
            För de personnumer som saknas kan det oftast vara så att det är de fyra
            sista som inte stämmer. Letar därför igenom elevlistan och se om det finns
            match både för de 6 första siffrorna i personnumret och för namnet.
            Avlägger sedan en rapport för detta till användaren så att man kan dubbelkolla
            att det blev rätt.
            """
            
            # De första sex siffrorna i personnummret i extens
            # som skall jämföras med de första sex siffrorna i
            # elevlistan nu när det inte matchade med alla tio 
            # siffrorna.
            pn_first_part_grp = index_grp[0:6]
            
            # En bool för att kunna fånga upp om en match helt saknas efter
            # att ha kollat första sex siffrorna och namnet.
            missing_match = True
            for index_elv, row_elv in tqdm(df_elev.iterrows(), ascii=True, leave=True, desc="Retry: " + row_grp.Namn):
                # Tar bort bindestrecket från personid i df_elev.
                # Skall inte finnas men gör det fall i fall.
                index_elv = str(index_elv).replace('-', '')
                
                if len(index_elv) < 10:
                    set_leading_zero(index_elv)
                    print(f'set_leading_zero({index_elv})')

                # De första sex siffrorna i elevlistan.
                pn_first_part_elv = index_elv[0:6]
                
                # kolla om det finns en rad där första sex siffrorna i extens och 
                # elevlistan stämmer överräns samtidigt som namnet är detsamma
                if pn_first_part_grp == pn_first_part_elv and row_grp.Namn == row_elv.Namn and missing_match:
                    row = [index_grp, df_elev.loc[index_elv].Grupper, df_elev.loc[index_elv].Namn +": "+index_elv+" ->  "+ row_grp.Namn +": "+index_grp] # lägg till personid, grupperna, namnet samt pn i elevlista -> pn i extens
                    missing_match = False
                    
            if missing_match:
                row = [index_grp, 'MISSING MATCH', 'MISSING MATCH']
                _errors.append(index_grp)
        
        content.append(row)
        
    
    return content, _errors

def clear_sheet(service, _range, ID):
    """
    Rensar sagd _range från sheet med id ID
    """
    service.spreadsheets().values().clear(spreadsheetId=ID, range=_range).execute()
                
def edit_sheet(service, content, _range, ID):
    """
    Uppdaterar sagd _range i sheet med id ID med innehållet content
    """
    try:
        range = _range
        values = content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        service.spreadsheets().values().update(spreadsheetId=ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
    except Exception as e:
        print("While trying to append values error: ", e)

def get_edukonto_reference_list(service, ELEVLISTA_ID):
    """
    Get edukonto sheet file in Drive and return it as a Dataframe
    """
    RANGE = 'edukonto!A1:D'
    FILENAME = 'edukonto_check'

     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ELEVLISTA_ID,
                                range=RANGE).execute()
    values = result.get('values', [])

    if not values:
        print('No data found in elevlista.')
    else:
        error_list = {
            "row": []
        }

        klasser = []
        personnummers = []
        names = []
        emails = []
        for i, row in enumerate(tqdm(values, ascii=True, desc="Get Edukonto ref. list")):
            # print(".", end="")

            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    klass = row[0]
                    personnummer = str(row[2])
                    name = row[1]
                    email = row[3]

                    klasser.append(klass)
                    personnummers.append(personnummer)
                    names.append(name)
                    emails.append(email)
                except Exception as e:
                    row = [i+1]
                    error_list['row'].extend(row)

        edukonto_dict = {
            'Klass': klasser,
            'Namn': names,
            'Personnummer': personnummers,
            'Email': emails
        }
        df_edukonto = pd.DataFrame.from_dict(edukonto_dict)

    return df_edukonto, error_list

def find_email(pn, name, df_edukonto):

    email = ""
    message = "OK"
    try:
        edu_konto_series = df_edukonto[df_edukonto['Personnummer'].str.contains(pn)]
        if len(edu_konto_series.index) == 1:
            email = str(edu_konto_series['Email'].values[0])
        else:
            edu_konto_series = df_edukonto[df_edukonto['Namn'].str.contains(name)]
            if len(edu_konto_series.index) == 1:
                email = str(edu_konto_series['Email'].values[0])
            else:
                message = "KONTOT SAKNAS I EDUKONTO SHEET"
    except Exception as e:
        print("Error while looking for email in edukonto sheet ", e)

    return email, message

def check_mail(service, df_elevlista, df_edukonto, ELEVLISTA_ID):
    elevlista_klass_list = df_elevlista.loc[:, "Klass"].tolist()
    elevlista_namn_list = df_elevlista.loc[:, "Namn"].tolist()
    elevlista_groups_list = df_elevlista.loc[:, "Grupper"].tolist()
    elevlista_personnummer_list = df_elevlista.loc[:, "Personnummer"].tolist()

    edukonto_namn_list = df_edukonto.loc[:, "Namn"].tolist()
    edukonto_personnummer_list = df_edukonto.loc[:, "Personnummer"].tolist()
    edukonto_edukonto_list = df_edukonto.loc[:, "Email"].tolist()

    missing_mail_klass = [] # for error handling. Will build df and then save missing mails to csv
    missing_mail_name = []
    content = []
    for i, current_pn in enumerate(tqdm(elevlista_personnummer_list, desc="Checking email")):
        row = []
        current_name = elevlista_namn_list[i]
        current_klass = elevlista_klass_list[i]
        current_groups = elevlista_groups_list[i]
        email, message = find_email(current_pn, current_name, df_edukonto)
        if email == "":
            # print("%s %s: %s" % (current_klass, current_name, email))
            pass
        if message != "OK":
            if current_klass in relevent_classes:
                cprint("     - %s %s SAKNAS I EDUKONTO SHEET" % (current_klass, current_name), 'yellow')
                missing_mail_klass.append(current_klass)
                missing_mail_name.append(current_name)

        row = [current_klass, current_name, current_groups, current_pn, email]

        content.append(row)

    if len(missing_mail_klass) > 0:
        missing_mail_dict = {
            'Klass': missing_mail_klass,
            'Namn': missing_mail_name
        }
        missing_mail_df = pd.DataFrame.from_dict(missing_mail_dict)
        time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04
        filename = time_stamp + '_edukonto_som_saknas.csv'
        try:
            missing_mail_df.to_csv(filename, sep=";", index=False)
            print("     [Saved info about missing edu accounts to file ", end="")
            cprint("'%s']" % (filename), 'cyan')
        except PermissionError:
            print(
                f"     SAVING FILE '{filename}' PERMISSION ERROR: Is file already open?")
        except:
            print("     Something went wrong!")

    print()
    sheet_range = "elevlista!A2"
    try:
        range = sheet_range
        values = content
        resource = {
            "values": values
        }
        # use append to add rows and update to overwrite
        response = service.spreadsheets().values().update(spreadsheetId=ELEVLISTA_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
    except Exception as e:
        print("While trying to append values error: ", e)

def createfile(new_df, group_file_name, group_name, message):
    """
    Saves the new dataframe to csv-file in corresponding folder
    ex year_7_files/7ABCNO-1.csv
    """
    # print()
    # print(message)
    new_df.to_csv(group_file_name, sep=",", index=False)    # dataframe to csv

def create_excel_file(new_df, group_file_name, group_name, message):
    """
    Saves the new dataframe to xlsx-file in corresponding folder
    ex grupper_åk_7/7ABCNO-1.xlsx
    """
    new_df.to_excel(group_file_name, index=False)    # dataframe to excel

def log_difference(new_df, old_df, group_name):
    """
    Creates at logfile in changelogs/ for those csv-files that is updated during the process.
    """
    log = "CHANGES MADE\n\n"
    time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04
    logged = False
    # REMOVED
    removed_df = old_df.merge(new_df,indicator = True, how='left').loc[lambda x : x['_merge']!='both']
    if len(removed_df.index) > 0:
        log += "Removed from "+ group_name +":\n" + removed_df.to_string() + "\n\n"
        logged = True

    # ADDED
    added_df = new_df.merge(old_df,indicator = True, how='left').loc[lambda x : x['_merge']!='both']
    if len(added_df.index) > 0:
        log += "Added in "+ group_name +":\n" + added_df.to_string()
        logged = True

    if not logged:
        log += "NOTHING ADDED OR REMOVED BUT SOMETHING CHANGED:\n\nOld_df:\n" + old_df.to_string()
        log += "\n\nNew_df:\n" + new_df.to_string()

    # CREATE LOG
    filepath = os.path.join(os.path.dirname(__file__), "changelogs/"+time_stamp+" "+group_name+" LOG.txt")
    log_file = open(filepath, "w")
    log_file.write(log)
    log_file.close()

    return group_name+".csv changed! LOG FILE CREATED -> '" + filepath + "'"

def df_sub_group_from_list_in_column(df, column, needle):
    boolean_mask = []
    for index, value in df.iterrows():

        # Ser till att kolumnen innehåller en lista
        if not type(value[column]) is list:
            column_list = value[column].split(',')
        else:
            column_list = value[column]
        
        clean_list = []
        # Ser till att ta bort mellanslag i båda ändar av strängen
        for e in column_list:
            clean_list.append(e.strip())

        if needle in clean_list:
            boolean_mask.append(True)
        else:
            boolean_mask.append(False)
        
    return boolean_mask

def generate_groups(elevlista):
    messages = []
    empty_groups = []
    dirname = os.path.dirname(__file__)
    counter = 0     # Counter for number of group.csv files created
    # for year in tqdm(arskurser, ascii=True, desc="Generate groups"):          # year: 7,...
    for year in arskurser:          # year: 7,...
        year_folder = 'year_'+year+'_files/'         # the folder to save the .csv in
        year_excel_folder = 'grupper_åk_'+year+'/'   # the folder to save the .xlsx (for human readability and upload to drive) in
        no_prefix_folder = 'no_prefix/'         # 'Modersmål Dari', 'ModersmålSOM', 'Lilla Världen' etc
        kulturskolan_folder = 'kulturskolan/'   # 'Valhalla', 'Dans',...
        kulturskolan_excel_folder = 'grupper_kulturskolan/'     # the folder to save the .xlsx for kulturskolans groups

        for key in tqdm(grupper, ascii=True, desc="YEAR "+year):                 # key: no,...
            """
            grupper (dict) in config.py
            """
            for group in tqdm(grupper[key], ascii=True, desc="Key: "+key, leave=False):            # group: abcno-1, abcno-2, abcno-3,...
                """
                grupper[key] (list) in config.py
                1. set group_name, group_file_name, excel_file_name depending on key in 'grupper'
                """
                group_name = 'not_set_yet'
                group_file_name = 'not_set_yet'
                excel_file_name = 'not_set_yet'
                # get group_name depending on key in 'grupper' (dict) in config.py
                if key is not 'no_prefix' and key is not 'kulturskolan':
                    # print()
                    # print("folder: " + year_folder)
                    group_name = year + group                       # the correct group name 7abcno-1, 7abcno-2,...
                    group_file_name = os.path.join(dirname, year_folder + group_name+".csv")         #/year_7_files/7abcno-1.csv
                    excel_file_name = os.path.join(dirname, year_excel_folder + group_name+".xlsx")  #/grupper_åk_7/7abcno-1.xlsx
                elif key == 'kulturskolan':
                    # print()
                    # print("folder: " + kulturskolan_folder)
                    group_name = group
                    group_file_name = os.path.join(dirname, kulturskolan_folder + group_name+".csv")    #/kulturskolan/VALHALLA.csv
                    excel_file_name = os.path.join(dirname, kulturskolan_excel_folder + group_name+".xlsx")     #/kulturskolan/VALHALLA.xlsx
                else:
                    # print()
                    # print("folder: " + no_prefix_folder)
                    group_name = group
                    no_prefix_file_name = os.path.join(dirname, no_prefix_folder + group_name+".xlsx")  #/no_prefix/modersmål som F-3.xlsx

                group_email = group_name.lower() + email_tail    # the groups email address 7abcno-1@edu.he.....



                # Get group_df from 'Elev Grupper' or from 'Elev Klass'
                if key is not 'klass': # if group should be a group in 'Elev Grupper'
                    # group_df = elevlista[elevlista['Elev Grupper'].str.contains(group_name)]            # Ny dataframe med alla elever som är med i gruppen (groupname)
                    group_df = elevlista[df_sub_group_from_list_in_column(elevlista, 'Elev Grupper', group_name)]

                    
                else:                   # if the group should be the class
                    group_df = elevlista[elevlista['Elev Klass'].str.contains(group_name[:2])]          # Ny dataframe med alla elever som är med i klassen

                group_size = len(group_df.index)       # number of rows in merged dataframe

                # composing the columns in the csv file to be
                group_email_column = [group_email] * group_size             # Group Email
                member_email_column = group_df['Elev Mail'].tolist()        # Member Email
                member_type_column = ['USER'] * group_size                  # Member Type
                member_role_column = ['MEMBER'] * group_size                # Member Role

                # Dictionary for csv-files that are purposed for massupload in Google Admin
                group_dict = {
                    'Group Email [Required]': group_email_column,
                    'Member Email': member_email_column,
                    'Member Type': member_type_column,
                    'Member Role': member_role_column
                }

                # Dictionary for xlsx files that are puropsed for drive upload.
                excel_dict = {
                    'Grupp': [group_name] * group_size,
                    'Klass': group_df['Elev Klass'].tolist(),
                    'Namn':  group_df['Elev Namn'].tolist()
                }

                new_df = pd.DataFrame.from_dict(group_dict)               # dataframe from created dict: group_dict

                excel_df = pd.DataFrame.from_dict(excel_dict)             # dataframe from created dict: drive_dict



                if len(new_df.index) > 0:   # if new_df contains rows

                    if os.path.exists(group_file_name): # check if this csv already exists
                        old_df = pd.read_csv(group_file_name)   # get old_df


                        if not new_df.equals(old_df):   # compare if new_df is not equal to old_df
                            message = log_difference(new_df, old_df, group_name)
                            messages.append(message)
                            if key is not 'no_prefix':
                                # message = log_difference(new_df, old_df, group_name)
                                # messages.append(message)

                                createfile(new_df, group_file_name, group_name, message)

                                # creating xlsx-files in 'grupper_åk_x/group_name'
                                excel_message = group_name+".xlsx created for drive!"
                                create_excel_file(excel_df, excel_file_name, group_name, excel_message)
                                counter += 1
                            else:
                                create_excel_file(excel_df, no_prefix_file_name, group_name, excel_message)

                        else:
                            # print(".", end="")
                            pass

                    else:
                        if key is not 'no_prefix':
                            # creating csv-files in 'year_x_files/group_name'
                            message = group_name+".csv created!"
                            createfile(new_df, group_file_name, group_name, message)
                            messages.append(message)

                            # creating xlsx-files in 'grupper_åk_x/group_name'
                            excel_message = group_name+".xlsx created for drive!"
                            create_excel_file(excel_df, excel_file_name, group_name, excel_message)
                            counter += 1
                        else:
                            excel_message = group_name+".xlsx created for drive!"
                            create_excel_file(excel_df, no_prefix_file_name, group_name, excel_message)




                else:
                    empty_groups.append(group_name)
    print()
    if len(messages) > 0:
        print("FOLLOWING CHANGES FOUND:")
        for msg in messages:
            print(f'    - {msg}')
    else:
        print("NO CHANGES FOUND!")


    return empty_groups, counter

from __future__ import print_function
import pandas as pd
from config import *
import datetime
import os
import sys
import getopt
import pickle
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
    for row_number in error_list:
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
        for i, row in enumerate(values):
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

        for i, row in enumerate(values):
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
        for i, row in enumerate(values):
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

    content = []
    for i, current_pn in enumerate(elevlista_personnummer_list):
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
                cprint("     %s %s SAKNAS I EDUKONTO SHEET" % (current_klass, current_name), 'yellow')

        row = [current_klass, current_name, current_groups, current_pn, email]   
        
        content.append(row)
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
    print()
    print(message)
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
        log += "Removed:\n" + removed_df.to_string() + "\n\n"
        logged = True

    # ADDED
    added_df = new_df.merge(old_df,indicator = True, how='left').loc[lambda x : x['_merge']!='both']
    if len(added_df.index) > 0:
        log += "Added:\n" + added_df.to_string()
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

def generate_groups(elevlista):
    empty_groups = []
    dirname = os.path.dirname(__file__) 
    counter = 0     # Counter for number of group.csv files created
    for year in arskurser:          # year: 7,...
        folder = 'year_'+year+'_files/'         # the folder to save the .csv in
        excel_folder = 'grupper_åk_'+year+'/'   # the folder to save the .xlsx (drive files) in
        no_prefix_folder = 'no_prefix/'
        print()
        print("folder: " + folder)
        for key in grupper:                 # key: no,...
            for group in grupper[key]:            # group: abcno-1, abcno-2, abcno-3,...
                if key is not 'no_prefix':
                    group_name = year + group                       # the correct group name 7abcno-1, 7abcno-2,...
                else:
                    group_name = group

                group_email = group_name.lower() + email_tail    # the groups email address 7abcno-1@edu.he.....


                group_file_name = os.path.join(dirname, folder + group_name+".csv")         #/year_7_files/7abcno-1.csv
                excel_file_name = os.path.join(dirname, excel_folder + group_name+".xlsx")  #/grupper_åk_7/7abcno-1.xlsx
                no_prefix_file_name = os.path.join(dirname, no_prefix_folder + group_name+".xlsx")  #/no_prefix/modersmål som F-3.xlsx

                
                group_df = elevlista[elevlista['Elev Grupper'].str.contains(group_name)]                # Ny dataframe med alla elever som är med i gruppen (groupname)
                # merged_df = pd.merge(elevmail, group_name_df, on=['Elev Namn'], how='inner')          # lägger samman dataframsen med avseende på elevnamnet

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

                        
                        if not new_df.equals(old_df):   # compare if new_df is equal to new
                            message = log_difference(new_df, old_df, group_name)
                            if key is not 'no_prefix':
                                message = log_difference(new_df, old_df, group_name)
                                createfile(new_df, group_file_name, group_name, message)

                                # creating xlsx-files in 'grupper_åk_x/group_name'
                                excel_message = group_name+".xlsx created for drive!"
                                create_excel_file(excel_df, excel_file_name, group_name, excel_message)
                                counter += 1
                            else:
                                create_excel_file(excel_df, no_prefix_file_name, group_name, excel_message)
                            
                        else:
                            print(".", end="")

                    else:
                        if key is not 'no_prefix':
                            # creating csv-files in 'year_x_files/group_name'
                            message = group_name+".csv created!"
                            createfile(new_df, group_file_name, group_name, message)

                            # creating xlsx-files in 'grupper_åk_x/group_name'
                            excel_message = group_name+".xlsx created for drive!"   
                            create_excel_file(excel_df, excel_file_name, group_name, excel_message)
                            counter += 1
                        else:
                            excel_message = group_name+".xlsx created for drive!" 
                            create_excel_file(excel_df, no_prefix_file_name, group_name, excel_message)
                        

                        

                else:
                    empty_groups.append(group_name)
    
    return empty_groups, counter
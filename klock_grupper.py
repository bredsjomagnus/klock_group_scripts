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
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']
# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = SHEET_ID


HELPMSG = """
klock_grupper reads 'elevlista' file on drive and save a copy of it in folder.
The file is used to generate the students groups and save those groups both as .csv
for mass upload purpose in Google Admin and as .xlsx for more readability to be
uploaded to the Drive.

USAGE
$ python klock_grupper.py

--------------------------------------------------------------------------------------------------------

The file in the Drive ('elevlista') needs to have the columns; [Elev Klass, Elev Namn, Elev Grupper, Elev Mail]'

--------------------------------------------------------------------------------------------------------

The script creates files like this (.csv meant for mass uploading in Google Admin):
Group Email [Required],     Member Email,                           Member Type,    Member Role
9cbl-1@edu.hellefors.se,    firstname.lastname@edu.hellefors.se,    USER,           MEMBER

And like this (.xlsx and meant readability and destined for the Drive):
Grupp	Klass	Namn
8CMU-1	8C      Abdi, Iman

--------------------------------------------------------------------------------------------------------

The script checks older files against new ones. If there is any difference a changelog file will be created in
/changelogs/group.txt. The log will state if anything has been added and/or removed and what that change is.

--------------------------------------------------------------------------------------------------------

Options:
    -h or --help        Display this help message
"""

dirname = os.path.dirname(__file__)         # this directory
empty_groups = []
options = ['-h', '--help']

def authenticate():
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    return service

def update_elevnamn_till_elevmail(service, elevlist_file_name):
    SAMPLE_RANGE_NAME = 'elevlista!A1:E'
    FILENAME = elevlist_file_name
    nan_counter = 0
    elev_counter = 0


     # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Läser in elevnamn_till_elevmail från Driven: ', end="")
        klasser = []
        names = []
        groups = []
        emails = []
        # print(values)
        for i, row in enumerate(values):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    # name = "\""+row[0]+"\""
                    klass = row[0]
                    name = row[1]
                    group = row[2]
                    email = row[4]
                    # print('%s, %s' % (name, email))
                    klasser.append(klass)
                    names.append(name)
                    groups.append(group)
                    emails.append(email)
                    elev_counter += 1
                except Exception as e:
                    nan_counter += 1
                    # print("While reading rows from file error ->", e)
                    # print("Null at ", row[1])
        print("%d valid rows and %d nan values" % (elev_counter, nan_counter))
        print()
        print("Skapar DataFrame och sparar som %s" % (FILENAME))
        elevlista_dict = {
            'Elev Klass': klasser,
            'Elev Namn': names,
            'Elev Grupper': groups,
            'Elev Mail': emails
        }
        elevlista_df = pd.DataFrame.from_dict(elevlista_dict)
        elevlista_df.to_csv(FILENAME, sep=",", index=False)
        


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
    print()
    print(message)
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

opts, args = getopt.getopt(sys.argv[1:], "h", ["help"])     # get options and arguments from console

# iterate options
for opt in opts:
    for optvalue in opt:
        # check if option is in valid option list
        if optvalue in options:
            if optvalue == '-h' or optvalue == '--help':
                print(HELPMSG)
                exit()
print("Beginning process...")
print()
print("## Looking for the essential files ##")

# csv file needed as argument
# if len(args) == 0:
#     print("Correct use: $ python klock_grupper.py <csv-file>")
#     print("Need CSV-file")
#     print()
#     print("Type python --help klock_grupper.py, for help")
#     print()
#     print("FAILURE! Closing process...")
#     exit()
# else:
#     csvfile = args[0]

try:
    time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04
    elevlist_file_name = time_stamp + '_elevlista.csv'
    service = authenticate()
    update_elevnamn_till_elevmail(service, elevlist_file_name)
    elevlista = pd.read_csv(elevlist_file_name).dropna() # read reference list with names and corresponding emails
    print("Found - 'elevnamn_till_elevmail.csv'")
except:
    print("Failed! Could not find - 'elevnamn_till_elevmail.csv'")
    print("'elevnamn_till_elevmail.csv' is needed to get students email addresses.")
    print()
    print("FAILURE! Closing process...")
    exit()

# try:
#     df = elevmail
#     # df = pd.read_csv(csvfile, sep=';', index_col=None).dropna() # read list from Infomentor
#     # print("Found - '" + csvfile + "'")
# except:
#     # print("Failed! Could not find file - '" + csvfile + "'")
#     print()
#     print("FAILURE! Closing process...")
#     exit()



counter = 0     # Counter for number of group.csv files created
for year in arskurser:          # year: 7,...
    folder = 'year_'+year+'_files/'         # the folder to save the .csv in
    drive_folder = 'grupper_åk_'+year+'/'   # the folder to save the .xlsx (drive files) in
    print()
    print("folder: " + folder)
    for key in grupper:                 # key: no,...
        for group in grupper[key]:            # group: abcno-1, abcno-2, abcno-3,...
            group_name = year + group                       # the correct group name 7abcno-1, 7abcno-2,...
            group_email = group_name.lower() + email_tail    # the groups email address 7abcno-1@edu.he.....


            group_file_name = os.path.join(dirname, folder + group_name+".csv")         #/year_7_files/7abcno-1.csv
            drive_file_name = os.path.join(dirname, drive_folder + group_name+".xlsx")  #/grupper_åk_7/7abcno-1.xlsx
            
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
            drive_dict = {
                'Grupp': [group_name] * group_size,
                'Klass': group_df['Elev Klass'].tolist(),
                'Namn':  group_df['Elev Namn'].tolist()
            }

            new_df = pd.DataFrame.from_dict(group_dict)               # dataframe from created dict: group_dict

            drive_df = pd.DataFrame.from_dict(drive_dict)             # dataframe from created dict: drive_dict

            # print(drive_df)

            if len(new_df.index) > 0:   # if new_df contains rows
                if os.path.exists(group_file_name): # check if this csv already exists
                    old_df = pd.read_csv(group_file_name)   # get old_df

                    
                    if not new_df.equals(old_df):   # compare if new_df is equal to new
                        message = log_difference(new_df, old_df, group_name)
                        createfile(new_df, group_file_name, group_name, message)
                        counter += 1
                    else:
                        print(".", end="")

                else:
                    # creating csv-files in 'year_x_files/group_name'
                    message = group_name+".csv created!"
                    createfile(new_df, group_file_name, group_name, message)

                    # creating xlsx-files in 'grupper_åk_x/group_name'
                    drive_message = group_name+".xlsx created for drive!"
                    create_excel_file(drive_df, drive_file_name, group_name, drive_message)
                    counter += 1

                    

            else:
                empty_groups.append(group_name)
                

            
print()
print()
print("Empty group(s):", empty_groups, ". Skipped!")
print()
print("DONE! %d files created!" % (counter))
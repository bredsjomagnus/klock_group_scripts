from __future__ import print_function
import pickle
import os.path
import sys
import re
import pprint
import pandas as pd
from env import *
from sheet_config import *
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

pp = pprint.PrettyPrinter(indent=2)
# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
# SAMPLE_SPREADSHEET_ID = SHEET_TO_UPDATE_ID
# SAMPLE_RANGE_NAME = 'Blad1!A1:B'

# Placeholders for soon to be created spreadsheet
spreadsheet = {}
SPREADSHEET_ID = ""

# FILENAME = 'elevnamn_till_elevmail_TEST.csv'
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

def create_spreadsheet(service):
    """
    Create new spreadsheet
    """
    print("Beginning process...")
    SPREADSHEET_ID = ""
    spreadsheet_body = {
        "properties": {
            "title": SPREADSHEET_TITLE
        }
    }
    try:
        request = service.spreadsheets().create(body=spreadsheet_body)
        spreadsheet = request.execute()
        SPREADSHEET_ID = spreadsheet['spreadsheetId']
        print("spreadsheet id: ", SPREADSHEET_ID)
    except Exception as e:
        print("While trying to create new spreadsheet error: ", e)
        sys.exit()
    
    return SPREADSHEET_ID

def update_spreadsheet(service, SPREADSHEET_ID, body, message="No message"):
    """
    Update batch of requests in body
    """

    response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    print(message)

    return response


def create_sheets(service, SPREADSHEET_ID, sheet_objects):
    requests = []
    print("Creating sheets...")
    # CREATE SHEETS
    for sheet in sheet_objects:
        requests.append(sheet_objects[sheet])

    # Trying to update spreadsheet with assigned requests
    try:
        body = {
            "requests": requests
        }
        response = update_spreadsheet(service, SPREADSHEET_ID, body, "Sheets created")
    except Exception as e:
        print("While trying to batchUpdate error: ", e)
        sys.exit()

    return response
    

def get_sheet_ids(service, SPREADSHEET_ID):
    """
    Getting the proporites of the created sheets.
    """
    print("Getting sheet proporties...")
    sheet_ids = []
    # Trying to get sheetIds
    try:
        fields="sheets.properties"
        request = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields=fields)
        response = request.execute()

        # TODO: Change code below to process the `response` dict:
        # pp.pprint(response)
        
        # How to iterate the response and get title of a sheet
        for prop in response['sheets']:
            print(prop['properties']['title'])
            sheet_ids.append(prop['properties']['sheetId'])
    except Exception as e:
        print("While trying spreadsheets().get() error: ", e)
        sys.exit()

    return sheet_ids

def customize_columns(service, SPREADSHEET_ID, columns):
    """
    Set column width in sheets. columns is generated in sheet_config.py
    based on template and sheetId
    """
    print("Setting column and row widths in the sheets")
    requests = []
    for key in columns:
        requests.append(columns[key])
    
    # Trying to update spreadsheet with assigned requests
    try:
        body = {
            "requests": requests
        }
        response = update_spreadsheet(service, SPREADSHEET_ID, body, "Columns set")
        return response
    except Exception as e:
        print("While trying to batchUpdate error: ", e)
        sys.exit()

def add_columns(service, SPREADSHEET_ID, add_column_objects):
    """
    Add columns to those sheets where need be. add_column_objects is generated
    in sheet_config.py based on tamplate and sheetId
    """
    print("Adding columns to sheets where necessary")
    requests = []
    for key in add_column_objects:
        requests.append(add_column_objects[key])
    # Trying to update spreadsheet with assigned requests
    try:
        body = {
            "requests": requests
        }
        response = update_spreadsheet(service, SPREADSHEET_ID, body, "Added columns")
        return response
    except Exception as e:
        print("While trying to batchUpdate error: ", e)
        sys.exit()
    
def add_content(service, SPREADSHEET_ID):
    """
    Add content to sheets cells.
    """
    print("Adding content to sheets...")
    df = pd.read_csv("elevlista.csv", index_col=False)
    for sheet in sheet_names:
        re_object = re.search("(\s\d\w+)", sheet)
        klass = re_object.group().strip()
        # print("sheet: %s, klass: %s" % (sheet, klass))
        klass_df = df[df['Elev Klass'].str.contains(klass)]

        # Collect students names
        elever_namn = klass_df.loc[:,'Elev Namn'].tolist()

        # Collect SVA students and put there names in a list
        elever_sva = klass_df[klass_df['Elev Grupper'].str.contains('SVA')]
        elever_sva = elever_sva.loc[:, 'Elev Namn'].tolist()

        # Collect personnummer and determine gender
        elever_personnummer = klass_df.loc[:, 'Elev Personnummer'].tolist()
        elever_gender = []
        for pn in elever_personnummer:
            pn = str(pn)
            pn = pn[-2:-1]
            if int(pn)%2 == 0:
                elever_gender.append('Flicka')
            else:
                elever_gender.append('Pojke')

        

        # print(elever_namn)
        # print(elever_personnummer)

        content = template_dict['template_1']['header']
        klass_range = sheet + "!A1:F"
        try:
            range = klass_range
            values = content
            resource = {
                "values": values
            }
            # use append to add rows and update to overwrite
            response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
            # print("appended value reponse: ", response)
        except Exception as e:
            print("While trying to append values error: ", e)
            sys.exit()

        # print("Klass %s: %s" % (klass, elever_namn))

        content = []
        # print("header", template_dict["template_1"]["header"])
        for i, namn in enumerate(elever_namn):
            elev = []
            elev.append(namn)
            if namn in elever_sva:
                elev.append('Ja')
            else:
                elev.append('Nej')
            elev.append(elever_gender[i])
            content.append(elev)
        # print("content", content)
        # template_dict["template_1"]["header"] = template_dict["template_1"]["header"]
        klass_range = sheet + "!A4:F"
        try:
            range = klass_range
            values = content
            resource = {
                "values": values
            }
            # use append to add rows and update to overwrite
            response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
            # print("appended value reponse: ", response)
        except Exception as e:
            print("While trying to append values error: ", e)
     


def main():
    """
    Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    

    service = authenticate()

    SPREADSHEET_ID = create_spreadsheet(service)

    create_sheets(service, SPREADSHEET_ID, sheet_objects)
    
    sheetIds = get_sheet_ids(service, SPREADSHEET_ID)

    add_column_objects = generate_add_column_object(sheet_template, sheetIds)

    if len(add_column_objects) > 0:
        add_columns(service, SPREADSHEET_ID, add_column_objects)
    
    columns = generate_columns_update_object(template_complied_results, sheetIds)
    
    customize_columns(service, SPREADSHEET_ID, columns)

    add_content(service, SPREADSHEET_ID)

    

if __name__ == '__main__':
    main()
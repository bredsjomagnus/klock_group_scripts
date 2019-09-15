from __future__ import print_function
import pickle
import os.path
import sys
import pprint
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
    SPREADSHEET_ID = ""
    spreadsheet_body = {
        "properties": {
            "title": "API genererat dokument"
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

    response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()

    # print(message)

    return response


def create_sheets(service, SPREADSHEET_ID, sheet_objects):
    requests = []

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
     


def main():
    """Shows basic usage of the Sheets API.
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
    



    # for setting in spreadsheet_settings:
    #     requests.append(spreadsheet_settings[setting])
    
    # for col_width in columns:
    #     requests.append(columns[col_width])

    # Change the spreadsheet's title.
    # requests.append({
    #     'updateSpreadsheetProperties': {
    #         'properties': {
    #             'title': 'KLOCKARHAGSSKOLAN Diamanttest'
    #         },
    #         'fields': 'title'
    #     }
    # })

    # # Update sheet name of first sheet.
    # requests.append({
    #     "updateSheetProperties": {
    #         "properties": {
    #             "sheetId": 0,
    #             "title": "Klass 7A",
    #         },
    #         "fields": "title"
    #     }
    # })
    # # Change tab color on first sheet
    # requests.append({
    #     "updateSheetProperties": {
    #         "properties": {
    #             "sheetId": 0,
    #             "tabColor": {
    #                 "red": 0.4,
    #                 "green": 0.3,
    #                 "blue": 1.0
    #             }
    #         },
    #         "fields": "tabColor"
    #     }
    # })
    # # Adjust column width
    # requests.append({
    #     "updateDimensionProperties": {
    #         "range": {
    #             "dimension": "COLUMNS",
    #             "startIndex": 0,
    #             "endIndex": 1
    #         },
    #         "properties": {
    #             "pixelSize": 50
    #         },
    #         "fields": "pixelSize"
    #     }
    # })
    # # Add new sheet tab
    # requests.append({
    #     "addSheet": {
    #         "properties": {
    #             "title": "Klass 7B",
    #             "tabColor": {
    #                 "red": 1.0,
    #                 "green": 0.3,
    #                 "blue": 0.4
    #             }
    #         }
    #     }
    # })
    

    # # Trying to get the properties of the spreadsheet
    # try:
    #     fields="sheets.properties"
    #     request = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID, fields=fields)
    #     response = request.execute()

    #     # TODO: Change code below to process the `response` dict:
    #     # pp.pprint(response)
        
    #     # How to iterate the response and get title of a sheet
    #     # for prop in response['sheets']:
    #     #     print(prop['properties']['title'])
    # except Exception as e:
    #     print("While trying spreadsheets().get() error: ", e)

    # Trying to update spreadsheet with assigned requests
    # try:
    #     # Gather the requests to body and batchUpdate
    #     body = {
    #         'requests': requests
    #     }
    #     response = service.spreadsheets().batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()
        
    #     print("SUCCESS")
    # except Exception as e:
    #     print("While trying to batchUpdate error: ", e)
    # # find_replace_response = response.get('replies')[1].get('findReplace')
    # # print('{0} replacements made.'.format(
    # #     find_replace_response.get('occurrencesChanged')))

    # try:
    #     range = "Klass 7A!A1:B"
    #     values = [
    #         ["Klass", "Namn"],
    #         ["7A", "Andersson, Magnus"]
    #     ]
    #     resource = {
    #         "values": values
    #     }
    #     # use append to add rows and update to overwrite
    #     response = service.spreadsheets().values().update(spreadsheetId=SPREADSHEET_ID, range=range, body=resource, valueInputOption="USER_ENTERED").execute()
    #     # print("appended value reponse: ", response)
    # except Exception as e:
    #     print("While trying to append values error: ", e)

if __name__ == '__main__':
    main()
from __future__ import print_function
import pandas as pd
import datetime
import os
import sys
import getopt
import pickle
import os.path
from tqdm import tqdm
from termcolor import colored, cprint
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

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

def get_sheet_values_service(service, ID, _range):
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=ID, range=_range).execute()
    return result.get('values', [])

def update_sheet_service(service, ID, _range, content):
    try:
        values = content
        resource = {
            "values": values
        }
        
        # use append to add rows and update to overwrite
        service.spreadsheets().values().update(spreadsheetId=ID, range=_range, body=resource, valueInputOption="USER_ENTERED").execute()
    except Exception as e:
        print("While trying to append values error: ", e)

def prepp_df_dict(_dict, org_header):
    checks_out = True
    res = {}
    header = []
    for key, value in tqdm(_dict.items(), ascii=True, desc="prepp_df_dict"):
        res[value] = []
        header.append(key)
        if not key in org_header:
            checks_out = False
    return res, header, checks_out

def get_sheet_as_df(service, ID, _range, col_map):
    """
    Hämtar ett sheet från Drive med id ID inom range _range och med
    header mapping enlig col_map.
    
    Returnerar DataFrame
    """
    
    values = get_sheet_values_service(service, ID, _range)
    
    # Den faktiska headern i sheeten
    org_header = values[0]
    
    if not values:
        print('No data found in ' + ID + ', range ' + _range + ' with col_map ' + col_map)
    else:
        errors = []
        
        # Preparerar data:
        # _dict (dict) får keys utefter col_map med tom lista som value för varje key.
        # header (list) sätts utefter col_map (vilka rubriker som skall extraheras)
        # checks_out (boolean) True om col_map keys finns med i org_header annars False
        _dict, header, checks_out = prepp_df_dict(col_map, org_header)
        if checks_out:
            for i, row in enumerate(tqdm(values, ascii=True, desc="get_sheet_as_df")):
                if i == 0:
                    # Skippar rubrikraden
                    pass
                else:
                    try:
                        for key in tqdm(header, ascii=True, leave=False):
                            _dict[col_map[key]].append(row[org_header.index(key)])
    
                    except Exception as e:
                        errors.append(row)
                            
        else:
            print(f'{col_map} and {org_header} doesn\'t match')
            exit()
    
    return pd.DataFrame.from_dict(_dict), errors

def leading_zeroes(df, columns, limit=10):
    """
    Lägger till inledande nolla till värdena i cellerna i
    columnerna columns (list) om värdet i cellen har är 
    kortare än limit (int)
    """
    header = df.columns.tolist()
    for index, row in tqdm(df.iterrows(), ascii=True, desc="leading_zeroes"):
        for col in tqdm(columns, ascii=True, leave=False, desc="col"):
            if len(str(row[col])) < limit:
                row[col] = '0' + str(row[col])
                    
    return df

def update_column_via_df(service, ID, _range, column, df):
    """
    TL;DR Ny data kommer från column (string) och sätts in på _range (string).
    
    Uppdaterar en kolumn i sheet med id ID på range _range.
    Den nya kolumnen hämtas från kolumn i df med rubrik column (string).
    """
    content = []
    service.spreadsheets().values().clear(spreadsheetId=ID, range=_range).execute()
    
    for index, row in tqdm(df.iterrows(), ascii=True, desc="update_column_via_df"):
        new_row = []
        new_row.append(row[column])
        content.append(new_row)
    
    update_sheet_service(service, ID, _range, content)


ID = input('ID: ')

service = authenticate()

col_map = {
    'Elev Personnummer': 'Personnummer'
}

df, errors = get_sheet_as_df(service, ID, 'Export!A1:D', col_map)

if len(errors) > 0:
    cprint("Errors found: ", 'yellow')
    print(errors)
    anw = ""
    while anw != 'j' or anw != 'n':
        anw = input("Conintue j/n? ")
        if anw == 'n':
            exit()
    

df = leading_zeroes(df, ['Personnummer'])

update_column_via_df(service, ID, 'Export!D2:D', 'Personnummer', df)

print()
print(f'service.spreadsheets().values().update -> ID: {ID} ', end="")
cprint("*** SUCCESS ***", 'green')
from __future__ import print_function
import pickle
import os.path
from env import *
import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = SHEET_ID
SAMPLE_RANGE_NAME = 'elevlista!B1:C'

FILENAME = 'elevnamn_till_elevmail.csv'

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
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

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])

    if not values:
        print('No data found.')
    else:
        print('Läser in elevnamn_till_elevmail från Driven:')
        names = []
        emails = []
        for i, row in enumerate(values):
            if i > 0:
                # Print columns A and E, which correspond to indices 0 and 4.
                try: # this will take care of eventually empty cells.
                    # name = "\""+row[0]+"\""
                    name = row[0]
                    email = row[1]
                    # print('%s, %s' % (name, email))
                    names.append(name)
                    emails.append(email)
                except Exception as e:
                    print("While reading rows from file error ->", e)
        print("Skapar DataFrame och sparar som %s" % (FILENAME))
        elevlista_dict = {
            'Elev Namn': names,
            'Elev Mail': emails
        }
        elevlista_df = pd.DataFrame.from_dict(elevlista_dict)
        elevlista_df.to_csv(FILENAME, sep=",", index=False)

if __name__ == '__main__':
    main()
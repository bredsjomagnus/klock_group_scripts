from __future__ import print_function
import pandas as pd
from config import *
from functions import *
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

dirname = os.path.dirname(__file__)         # this directory

options = ['-h', '--help']

opts, args = getopt.getopt(sys.argv[1:], "h", ["help"])     # get options and arguments from console

# iterate options
for opt in opts:
    for optvalue in opt:
        # check if option is in valid option list
        if optvalue in options:
            if optvalue == '-h' or optvalue == '--help':
                print(HELPMSG)
                exit()
# print("Beginning process...")
print()
print("### 1/4 CHECKING EDU-MAIL ###")
print()

service = authenticate()


df_elevlista, elevlist_errors = get_elevlista_without_mail(service, ELEVLISTA_ID)
# print("Reading 'elavlista:elevlista.")
# print("Clearing 'Elev Mail' column.")
if len(elevlist_errors) > 0:
    error_report(elevlist_errors)
else:
    print("Get Infometor (elevlista) file in Drive and return it as a Dataframe: ", end="")
    cprint("*** SUCCESS ***", 'green')


df_edukonto, edulist_errors = get_edukonto_reference_list(service, ELEVLISTA_ID)
# print("Reading 'elevlista:edukonto.")
if len(edulist_errors['row']) == 0:
    print("Get edukonto sheet file in Drive and return it as a Dataframe: ", end="")
    cprint("*** SUCCESS ***", 'green')
    print()
else:
    cprint("--- len(edulist_errors['row']) != 0 ---", 'red')
    cprint("    TOTALLY %d ROWS IN ERROR LIST" % (len(edulist_errors['row'])), "yellow")
    cprint("    edulist_errors: %s " % (edulist_errors), 'yellow')
    print()

# print()
# print("Checking mail.")
# print()
check_mail(service, df_elevlista, df_edukonto, ELEVLISTA_ID)
# print("Update elevlista:elevlista, email set.")
print("### 1/4 DONE ###")
print()
print("### 2/4 CHECKING LANGUAGE ###")
print()

# Get elevlista with emails as df
sheet_name = 'elevlista'
col_map = {
    'Elev Klass': 'Klass',
    'Elev Namn': 'Namn',
    'Elev Grupper': 'Grupper',
    'Elev Personnummer': 'Personnummer',
    'Elev Mail': 'Mail',
}

df_elevlista_with_mails, errors = get_sheet_as_df(service, ELEVLISTA_ID, sheet_name+'!A1:E', col_map)

# Get sva_sv as df
sheet_name = 'sva_sv'
col_map = {
    'Elev': 'Namn',
    'Personnummer': 'Personnummer',
    'Språk': 'Språk'
}
df_sva_sv, errors = get_sheet_as_df(service, ELEVLISTA_ID, sheet_name+'!A1:E', col_map)

check_language(service, df_elevlista_with_mails, df_sva_sv, ELEVLISTA_ID)


# print()

# print()
cont = input("3/4 CONTINUE WITH GROUPS j/n? ")
if cont == "j":
    df_elevlista = get_elevlista_with_emails(service, ELEVLISTA_ID)
    empty_groups, files_created = generate_groups(df_elevlista)
    # print()
    # print()
    # print("Empty group(s):", empty_groups, ". Skipped!")
    print()
    print("### 3/4 DONE! %d files created! ###" % (files_created))
else:
    print()
    print()
    print("Mail checked! Aborting.")

cont = input("4/4 CONTINUE WITH GROUPS IMPORTS j/n? ")
if cont == "j":
    df_group, _errors_gruppimport = get_groupimport(service, EXTENS_ID)
    if len(_errors_gruppimport) > 0:
        print(f'_errors_gruppimport: {_errors_gruppimport}')
        print('Shutting down')
        exit()
    df_elev, _errors_elevlista = get_elevlista_with_personummer_as_index(service, ELEVLISTA_ID)
    if len(_errors_elevlista) > 0:
        print(f'_errors_elevlista: {_errors_elevlista}')
        print('Shutting down')
        exit()

    content, _errors = get_group_import_content(df_group, df_elev)

    clear_sheet(service, 'gruppimport!A2:C', ELEVLISTA_ID)
    edit_sheet(service, content, "gruppimport!A2", ELEVLISTA_ID)
    
    print()
    print("service.spreadsheets().values().update elevlista->gruppimport: ", end="")
    cprint("*** SUCCESS ***", 'green')

    print()
    if len(_errors) > 0:
        print(f'Totally missing matches: {_errors}')
    else:
        print("NO MISSING MATCHES!")

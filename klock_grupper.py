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
print("Beginning process...")
print()

service = authenticate()

df_elevlista = get_elevlista_without_mail(service, ELEVLISTA_ID)
df_edukonto = get_edukonto_reference_list(service, ELEVLISTA_ID)
check_mail(service, df_elevlista, df_edukonto, ELEVLISTA_ID)
cont = input("Vill du fortsätta j/n? ")
if cont == "j":
    df_elevlista = get_elevlista_with_emails(service, ELEVLISTA_ID)
    empty_groups, files_created = generate_groups(df_elevlista)
    print()
    print()
    print("Empty group(s):", empty_groups, ". Skipped!")
    print()
    print("DONE! %d files created!" % (files_created))
else:
    print()
    print()
    print("Mail checked! Aborting.")


import pandas as pd
from config import *
import datetime
import os
import sys
import getopt

HELPMSG = """
klock_grupper reads csv downloaded from Infomentor with students name and
groups. Then create one csv-file for every group that can be used to
create these groups in Google Admin Groups.

To make this work there needs to be a reference csv with the students
names and email addresses [elevnamn_till_elevmail.csv]

USAGE
$ python klock_grupper.py <csv-file>
The csv-file needs to have semicolon as separator. But creates csv-files with comma as separator.

csv-file head example from Infomentor:
Elev Grupper;Elev Namn;Elev Klass;Årskurs
9ABCNO-2, 9CBL-1, 9CEN, 9CHKK-1, 9CIDH, 9CMU-1, 9CSL-1, 9CSO;last name, first name;9C;9
...

Then creates files like this:
Group Email [Required],Member Email,Member Type,Member Role
9cbl-1@edu.hellefors.se,firstname.lastname@edu.hellefors.se,USER,MEMBER
...

Options:
    -h or --help        Display this help message
"""

dirname = os.path.dirname(__file__)         # this directory
empty_groups = []
options = ['-h', '--help']

print("Beginning process...")
print()


def createfile(new_df, group_file_name, group_name, message):
    """
    Saves the new dataframe to csv-file in corresponding folder
    ex year_7_files/7ABCNO-1.csv
    """
    print()
    print(message)
    new_df.to_csv(group_file_name, sep=",", index=False)    # csv from dataframe

def log_difference(new_df, old_df, group_name):
    """
    Creates at logfile in changelogs/ for those csv-files that is updated during the process.
    """
    log = "CHANGES MADE\n\n"
    time_stamp = str(datetime.datetime.now())[:10] # 2019-09-04

    # REMOVED
    removed_df = old_df.merge(new_df,indicator = True, how='left').loc[lambda x : x['_merge']!='both']
    if len(removed_df.index) > 0:
        log += "Removed:\n" + removed_df.to_string() + "\n\n"

    # ADDED
    added_df = new_df.merge(old_df,indicator = True, how='left').loc[lambda x : x['_merge']!='both']
    if len(added_df.index) > 0:
        log += "Added:\n" + added_df.to_string()

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

print("## Looking for the essential files ##")

# csv file needed as argument
if len(args) == 0:
    print("Correct use: $ python klock_grupper.py <csv-file>")
    print("Need CSV-file")
    print()
    print("Type python --help klock_grupper.py, for help")
    print()
    print("FAILURE! Closing process...")
    exit()
else:
    csvfile = args[0]


try:
    df = pd.read_csv(csvfile, sep=';', index_col=None).dropna()       # read list from Infomentor
    print("Found - '" + csvfile + "'")
except:
    print("Failed! Could not find file - '" + csvfile + "'")
    print()
    print("FAILURE! Closing process...")
    exit()

try:
    elevmail = pd.read_csv('elevnamn_till_elevmail.csv')                # read reference list with names and corresponding emails
    print("Found - 'elevnamn_till_elevmail.csv'")
except:
    print("Failed! Could not find - 'elevnamn_till_elevmail.csv'")
    print("'elevnamn_till_elevmail.csv' is needed to get students email addresses.")
    print()
    print("FAILURE! Closing process...")
    exit()
counter = 0     # Counter for number of group.csv files created
for year in arskurser:          # year: 7,...
    folder = 'year_'+year+'_files/' # the folder to save the .csv in
    print()
    print("folder: " + folder)
    for key in grupper:                 # key: no,...
        for group in grupper[key]:            # group: abcno-1, abcno-2, abcno-3,...
            group_name = year + group                       # the correct group name
            group_email = group_name.lower() + email_tail    # the groups email address

            group_file_name = os.path.join(dirname, folder + group_name+".csv")
            
            group_name_df = df[df['Elev Grupper'].str.contains(group_name)]                 # Ny dataframe med alla elever som är med i gruppen (groupname)
            merged_df = pd.merge(elevmail, group_name_df, on=['Elev Namn'], how='inner')    # lägger samman dataframsen med avseende på elevnamnet

            group_size = len(merged_df.index)       # number of rows in merged dataframe

            # composing the columns in the csv file to be
            group_email_column = [group_email] * group_size             # Group Email
            member_email_column = merged_df['Elev Mail'].tolist()       # Member Email
            member_type_column = ['USER'] * group_size                  # Member Type
            member_role_column = ['MEMBER'] * group_size                # Member Role

            group_dict = {
                'Group Email [Required]': group_email_column,
                'Member Email': member_email_column,
                'Member Type': member_type_column,
                'Member Role': member_role_column
            }

            new_df = pd.DataFrame.from_dict(group_dict)               # dataframe from created dict: group_dict

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
                    message = group_name+".csv created!"
                    createfile(new_df, group_file_name, group_name, message)
                    counter += 1
            else:
                empty_groups.append(group_name)
                

            
print()
print()
print("Empty group(s):", empty_groups, ". Skipped!")
print()
print("DONE! %d files created!" % (counter))
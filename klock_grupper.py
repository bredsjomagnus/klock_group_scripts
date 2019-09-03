import pandas as pd
from config import *
import os

dirname = os.path.dirname(__file__)         # this directory

df = pd.read_csv('grupper_alla.csv', sep=';', index_col=None)       # read list from Infomentor
elevmail = pd.read_csv('elevnamn_till_elevmail.csv')                # read reference list with names and corresponding emails
counter = 0     # Counter for number of group.csv files created
for year in arskurser:          # year: 7,...
    folder = 'year_'+year+'_files/' # the folder to save the .csv in
    for key in grupper:                 # key: no,...
        for group in grupper[key]:            # group: abcno-1, abcno-2, abcno-3,...
            group_name = year + group                       # the correct group name
            group_email = group_name.lower() + email_tail    # the groups email address
            
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

            final_df = pd.DataFrame.from_dict(group_dict)               # dataframe from created dict: group_dict

            final_df.to_csv(os.path.join(dirname, folder + group_name+".csv"), sep=",", index=False)    # csv from dataframe
            counter += 1

print("DONE! %d files created!" % (counter))
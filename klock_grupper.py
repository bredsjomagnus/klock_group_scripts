import pandas as pd
from config import *
import numpy as np

df = pd.read_csv('grupper_alla.csv', sep=';', index_col=None)       # läser in listan exporterad från Infomentor
elevmail = pd.read_csv('elevnamn_till_elevmail.csv')                # läser in referenslistan som kopplar elevnamn till elevmail

# print(elevmail.loc[:, 'Elev Namn'])


# Gets the dataframe with all students in group '7ABCNO-1'
# NO1 = df[df['Elev Grupper'].str.contains("7ABCNO-1")]

# print(NO1.loc[: , ['Elev Namn', 'Elev Klass']])

for year in arskurser:                          # 7,...
    for key in grupper:                             # no,...
        for group in grupper[key]:                      #abcno-1, abcno-2, abcno-3,...
            groupname = year + group    # gruppnamnet
            groupmail = groupname.lower() + mailtail    # gruppens mailadress

            groupdf = df[df['Elev Grupper'].str.contains(groupname)]                # Ny dataframe med alla elever som är med i gruppen (groupname)
            maildf = pd.merge(elevmail, groupdf, on=['Elev Namn'], how='inner')     # lägger samman dataframsen med avseende på elevnamnet
            maildf['Elev Mail']                                                     # en series med eleverna i gruppens mailadresser
            # print()
            # print(groupname.upper())
            # print(type(mergedStuff['Elev Mail']))
            # print(mergedStuff['Elev Mail'])
            # print(mergedStuff)
            # print(groupdf.loc[: , ['Elev Namn', 'Elev Klass']])
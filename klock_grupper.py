import pandas as pd
from config import *

df = pd.read_csv('grupper_alla.csv', sep=';', index_col=None)
elevmail = pd.read_csv('elevnamn_till_elevmail.csv')

# print(elevmail.loc[:, 'Elev Namn'])


# Gets the dataframe with all students in group '7ABCNO-1'
NO1 = df[df['Elev Grupper'].str.contains("7ABCNO-1")]

# print(NO1.loc[: , ['Elev Namn', 'Elev Klass']])

for year in arskurser:
    for key in grupper:
        for group in grupper[key]:
            groupname = year + group    # gruppnamnet
            groupdf = df[df['Elev Grupper'].str.contains(groupname)]    # Ny dataframe med alla elever som är med i gruppen

            mergedStuff = pd.merge(elevmail, groupdf, on=['Elev Namn'], how='inner')
            print(groupname.lower() + mailtail)
            # print()
            # print(groupname.upper())
            # print(mergedStuff['Elev Mail'])
            # print(mergedStuff)
            # print(groupdf.loc[: , ['Elev Namn', 'Elev Klass']])
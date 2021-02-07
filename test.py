import os
from datetime import date

from tabulate import tabulate as tab
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image, ImageOps

today = date.today()
workDirectory = r'C:/Users/Georges/Downloads/'
outputExcelFile = workDirectory+str(today)+' Stats AMS.xlsx'
# For now from an Excel import, later we will use the API
inputExcelFile = workDirectory+'User_export_'+str(today)+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='Export', engine='openpyxl',
                   usecols=['ID', 'Email', 'Not blocked', 'Created at', 'Account activation date', 'Live Location:Country',
                            'Industries:Industries',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_', 'Last Membership:Type name'
                            ])


# COUNT ACTIVATION
activeUsers = df[df['Not blocked'] == True].count()['Account activation date']
nonActiveUsers = df['Account activation date'].isna().sum()
blockedUsers = df[df['Not blocked'] == False].count()['Account activation date']
allUsers = df['ID'].count()

activationLabel = []
myCounts = []
activationLabel.extend(('Confirmed', 'Unconfirmed', 'Blocked', 'Total'))
myCounts.extend((activeUsers, nonActiveUsers, blockedUsers, allUsers))

ActivationDict = list(zip(activationLabel, myCounts))
df_ActivationCount = pd.DataFrame(ActivationDict, columns =['Users', 'Total'])

df_ActivationCount['%'] = (df_ActivationCount['Total'] / allUsers) * 100
df_ActivationCount['%'] = df_ActivationCount['%'].round(decimals=2)

print(tab(df_ActivationCount, headers='keys', tablefmt='psql', showindex=False))
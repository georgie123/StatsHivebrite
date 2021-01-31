
from datetime import date
import pandas as pd
from tabulate import tabulate as tab

today = date.today()

# For now from an Excel import, later we will use the API
df = pd.read_excel('C:/Users/Georges/Downloads/User_export_2021-01-28.xlsx',
                   sheet_name='Export', engine='openpyxl',
                   usecols=['Live Location:Country', 'Email'])

# COUNT COUNTRY
df_Country_count = pd.DataFrame(df.groupby(['Live Location:Country']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT EMAIL DOMAINS
df['Domain'] = df['Email'].str.split('@').str[1]
df_Email_DNS_count = pd.DataFrame(df.groupby(['Domain']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# EXCEL FILE
writer = pd.ExcelWriter('C:/Users/Georges/Downloads/'+str(today)+' Stats Hivebrite.xlsx', engine='xlsxwriter')
df_Country_count.to_excel(writer, index=False, sheet_name='Country', header=['Country', 'Total'])
df_Email_DNS_count.to_excel(writer, index=False, sheet_name='Email DNS', header=['Domain', 'Total'])
writer.save()

print("OK, export done!")


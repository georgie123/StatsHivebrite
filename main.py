from datetime import date

import pandas as pd
from tabulate import tabulate as tab

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

today = date.today()

outputExcelFile = r'C:/Users/Georges/Downloads/'+str(today)+' Stats Hivebrite.xlsx'

# For now from an Excel import, later we will use the API
inputExcelFile = r'C:/Users/Georges/Downloads/User_export_'+str(today)+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='Export', engine='openpyxl',
                   usecols=['ID', 'Email', 'Account activation date', 'Live Location:Country', 'Industries:Industries',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_', 'Last Membership:Type name'
                            ])

# COUNT ACTIVATION
activeUsers = df['Account activation date'].count()
nonActiveUsers = df['Account activation date'].isna().sum()
allUsers = df['ID'].count()

myLabels = []
myCounts = []
myLabels.extend(('Confirmed', 'Unconfirmed', 'Total'))
myCounts.extend((activeUsers, nonActiveUsers, allUsers))

ActivationDict = list(zip(myLabels, myCounts))
df_ActivationCount = pd.DataFrame(ActivationDict, columns =['Users', 'Total'])

df_ActivationCount['%'] = (df_ActivationCount['Total'] / allUsers) * 100
df_ActivationCount['%'] = df_ActivationCount['%'].round(decimals=1)

# COUNT COUNTRY
df_Country_count = pd.DataFrame(df.groupby(['Live Location:Country'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Country_count = df_Country_count.fillna('Unknow')

df_Country_count['Percent'] = (df_Country_count['Total'] / df_Country_count['Total'].sum()) * 100
df_Country_count['Percent'] = df_Country_count['Percent'].round(decimals=1)

# COUNT CATEGORIES (CUSTOM FIELD _8f70fe1e_Occupation)
df['Categories'] = df['_8f70fe1e_Occupation'].str.split(': ').str[0]
df_Categories_count = pd.DataFrame(df.groupby(['Categories'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Categories_count = df_Categories_count.fillna('Unknow')

df_Categories_count['Percent'] = (df_Categories_count['Total'] / df_Categories_count['Total'].sum()) * 100
df_Categories_count['Percent'] = df_Categories_count['Percent'].round(decimals=1)

# COUNT SPECIALTIES (CUSTOM FIELD _8f70fe1e_Occupation)
df['Specialties'] = df['_8f70fe1e_Occupation'].str.split(': ').str[1]
df_Specialties_count = pd.DataFrame(df.groupby(['Specialties'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Specialties_count = df_Specialties_count.fillna('Unknow')

df_Specialties_count['Percent'] = (df_Specialties_count['Total'] / df_Specialties_count['Total'].sum()) * 100
df_Specialties_count['Percent'] = df_Specialties_count['Percent'].round(decimals=1)

# COUNT SPECIALTIES PER COUNTRY
df_SpecialtiesPerCountry_count = pd.DataFrame(df.groupby(['Live Location:Country', 'Specialties'], dropna=False)\
    .size(), columns=['Total']).sort_values(['Live Location:Country', 'Total'], ascending=[True, False]).reset_index()
df_SpecialtiesPerCountry_count = df_SpecialtiesPerCountry_count.fillna('Unknow')

df_SpecialtiesPerCountry_count['Percent'] = (df_SpecialtiesPerCountry_count['Total'] / df_SpecialtiesPerCountry_count['Total'].sum()) * 100
df_SpecialtiesPerCountry_count['Percent'] = df_SpecialtiesPerCountry_count['Percent'].round(decimals=2)

# COUNT EXPERTISE & INTERESTS (CUSTOM FIELD Industries:Industries)
df_tempIndustries = pd.DataFrame(pd.melt(df['Industries:Industries'].str.split(',', expand=True))['value'])
df_Industries_count = pd.DataFrame(df_tempIndustries.groupby(['value'], dropna=False).size(), columns=['Total'])\
    .reset_index()
df_Industries_count = df_Industries_count.fillna('AZERTY')

df_Industries_count['Percent'] = (df_Industries_count['Total'] / df.shape[0]) * 100
df_Industries_count['Percent'] = df_Industries_count['Percent'].round(decimals=2)

# EMPTY VALUES
industriesEmpty = df['Industries:Industries'].isna().sum()
industriesEmptyPercent = round((industriesEmpty / df.shape[0]) * 100, 2)

# REPLACE EMPTY VALUES AND SORT
df_Industries_count.loc[(df_Industries_count['value'] == 'AZERTY')] = [['Unknow', industriesEmpty, industriesEmptyPercent]]
df_Industries_count = df_Industries_count.sort_values(['Total'], ascending=False)

# COUNT EMAIL DOMAINS
df['Domain'] = df['Email'].str.split('@').str[1]
df_Email_DNS_count = pd.DataFrame(df.groupby(['Domain'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Email_DNS_count = df_Email_DNS_count.fillna('Unknow')

df_Email_DNS_count['Percent'] = (df_Email_DNS_count['Total'] / df_Email_DNS_count['Total'].sum()) * 100
df_Email_DNS_count['Percent'] = df_Email_DNS_count['Percent'].round(decimals=1)

# COUNT HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How_did_you_hear_about_us_)
df_HowDidYouHearAboutUs_count = pd.DataFrame(df.groupby(['_ed5be3a0_How_did_you_hear_about_us_'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_HowDidYouHearAboutUs_count = df_HowDidYouHearAboutUs_count.fillna('Unknow')

df_HowDidYouHearAboutUs_count['Percent'] = (df_HowDidYouHearAboutUs_count['Total'] / df_HowDidYouHearAboutUs_count['Total'].sum()) * 100
df_HowDidYouHearAboutUs_count['Percent'] = df_HowDidYouHearAboutUs_count['Percent'].round(decimals=1)

# COUNT MEMBERSHIP (FIELD Last Membership:Type name)
df_Membership_count = pd.DataFrame(df.groupby(['Last Membership:Type name'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Membership_count = df_Membership_count.fillna('Basic Membership')

df_Membership_count['Percent'] = (df_Membership_count['Total'] / df_Membership_count['Total'].sum()) * 100
df_Membership_count['Percent'] = df_Membership_count['Percent'].round(decimals=1)

# EXCEL FILE
writer = pd.ExcelWriter(outputExcelFile, engine='xlsxwriter')

df_ActivationCount.to_excel(writer, index=False, sheet_name='Status', header=True)
df_Country_count.to_excel(writer, index=False, sheet_name='Countries', header=['Country', 'Total', '%'])
df_Categories_count.to_excel(writer, index=False, sheet_name='Categories', header=['Category', 'Total', '%'])
df_Specialties_count.to_excel(writer, index=False, sheet_name='Specialties', header=['Specialty', 'Total', '%'])
df_SpecialtiesPerCountry_count.to_excel(writer, index=False, sheet_name='Specialties per country', header=['Country', 'Specialty', 'Total', '%'])
df_Industries_count.to_excel(writer, index=False, sheet_name='Expertise & Interests', header=['Expertise or Interest', 'Total', '%'])
df_Email_DNS_count.to_excel(writer, index=False, sheet_name='Email domains', header=['Email domain', 'Total', '%'])
df_HowDidYouHearAboutUs_count.to_excel(writer, index=False, sheet_name='How Did You Hear', header=['How did you hear about us', 'Total', '%'])
df_Membership_count.to_excel(writer, index=False, sheet_name='Memberships', header=['Membership', 'Total', '%'])

writer.save()

# EXCEL FILTERS
workbook = openpyxl.load_workbook(outputExcelFile)
sheetsLits = workbook.sheetnames

for sheet in sheetsLits:
    if sheet == 'Status':
        continue
    worksheet = workbook[sheet]
    FullRange = 'A1:' + get_column_letter(worksheet.max_column) + str(worksheet.max_row)
    worksheet.auto_filter.ref = FullRange
    workbook.save(outputExcelFile)

# EXCEL COLORS
for sheet in sheetsLits:
    worksheet = workbook[sheet]
    for cell in workbook[sheet][1]:
        worksheet[cell.coordinate].fill = PatternFill(fgColor = 'FFC6C1C1', fill_type = 'solid')
        workbook.save(outputExcelFile)

# EXCEL COLUMN SIZE
for sheet in sheetsLits:
    for cell in workbook[sheet][1]:
        if get_column_letter(cell.column) == 'A':
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 30
        else:
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 10
        workbook.save(outputExcelFile)

# TERMINAL OUTPUTS
print(tab(df_Country_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Categories_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Specialties_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_SpecialtiesPerCountry_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Industries_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Email_DNS_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_HowDidYouHearAboutUs_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Membership_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_ActivationCount, headers='keys', tablefmt='psql', showindex=False))
print(today)
print("OK, export done!")

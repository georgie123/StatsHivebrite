import os
from datetime import date
from tabulate import tabulate as tab

import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image, ImageOps

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

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


# COUNT REGISTRATIONS BY DATE (FIELD Created at)
df['Created'] = pd.to_datetime(df['Created at'])
df['Created'] = df['Created'].dt.to_period("M")
df_TempCreated = pd.DataFrame(df['Created'])

df_Created_count = pd.DataFrame(df_TempCreated.groupby(['Created'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Created'], ascending=True).reset_index()

df_Created_count['Created'] = df_Created_count['Created'].dt.strftime('%b %Y')

ind_drop = df_Created_count[df_Created_count['Created'].apply(lambda x: x.startswith('Feb 2020'))].index
df_Created_count = df_Created_count.drop(ind_drop)


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

# REPLACE SOME VALUES
df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'] = df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'].replace(['Email from partners (AMWC, VCS, FACE etc)'],'Email from partners')
df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'] = df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'].replace(['Other: please specify'],'Other')


# COUNT MEMBERSHIP (FIELD Last Membership:Type name)
df_Membership_count = pd.DataFrame(df.groupby(['Last Membership:Type name'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Membership_count = df_Membership_count.fillna('Basic Membership')

df_Membership_count['Percent'] = (df_Membership_count['Total'] / df_Membership_count['Total'].sum()) * 100
df_Membership_count['Percent'] = df_Membership_count['Percent'].round(decimals=1)


# EXCEL FILE
writer = pd.ExcelWriter(outputExcelFile, engine='xlsxwriter')

df_ActivationCount.to_excel(writer, index=False, sheet_name='Registrations', header=True)
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
    if sheet == 'Registrations':
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


# CHART USERS STATUS
activationLabel.pop()

activationValue = []
activationValue.extend([activeUsers, nonActiveUsers, blockedUsers])

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig1 = plt.figure()
plt.pie(activationValue, labels=activationLabel, colors=colors, autopct='%1.2f%%', shadow=False, startangle=90)
plt.axis('equal')
plt.title('User status', pad=20, fontsize=15)

fig1.savefig(workDirectory+'myplot1.png', dpi=70)
plt.clf()

im = Image.open(workDirectory+'myplot1.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot1.png')

# INSERT CHART IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot1.png')
img.anchor = 'A8'

workbook['Registrations'].add_image(img)
workbook.save(outputExcelFile)


# CHART CATEGORIES
chartLabel = df_Categories_count['Categories'].tolist()
chartLegendLabel = df_Categories_count['Categories'].tolist()
chartValue = df_Categories_count['Total'].tolist()
chartLegendPercent = df_Categories_count['Percent'].tolist()

chartLabel[-1] = ''

legendLabels = []
for i, j in zip(chartLegendLabel, map(str, chartLegendPercent)):
    legendLabels.append(i + ' (' + j + ' %)')

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig2 = plt.figure()
plt.pie(chartValue, labels=chartLabel, colors=colors, autopct=None, shadow=False, startangle=90)
plt.axis('equal')
plt.title('Categories', pad=20, fontsize=15)

plt.legend(legendLabels, loc='best', fontsize=8)

fig2.savefig(workDirectory+'myplot2.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot2.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot2.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot2.png')
img.anchor = 'E4'

workbook['Categories'].add_image(img)
workbook.save(outputExcelFile)


# CHART HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How_did_you_hear_about_us_)
ind_drop = df_HowDidYouHearAboutUs_count[df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'].apply(lambda x: x.startswith('Unknow'))].index
df_HowDidYouHearAboutUs_count = df_HowDidYouHearAboutUs_count.drop(ind_drop)

chartLabel = df_HowDidYouHearAboutUs_count['_ed5be3a0_How_did_you_hear_about_us_'].tolist()
chartValue = df_HowDidYouHearAboutUs_count['Total'].tolist()

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig3 = plt.figure()
plt.pie(chartValue, labels=None, colors=colors, autopct='%1.1f%%', shadow=False, startangle=90)

plt.axis('equal')
plt.title('How did you hear about us (know)', pad=20, fontsize=15)

plt.legend(chartLabel, loc='best', fontsize=8)

fig3.savefig(workDirectory+'myplot3.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot3.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot3.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot3.png')
img.anchor = 'E4'

workbook['How Did You Hear'].add_image(img)
workbook.save(outputExcelFile)


# CHART MEMBERSHIP (FIELD Last Membership:Type name)
chartLabel = df_Membership_count['Last Membership:Type name'].tolist()
chartValue = df_Membership_count['Total'].tolist()
chartLegendPercent = df_Membership_count['Percent'].tolist()

legendLabels = []
for i, j in zip(chartLabel, map(str, chartLegendPercent)):
    legendLabels.append(i + ' (' + j + ' %)')

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig4 = plt.figure()
plt.pie(chartValue, labels=None, colors=colors, autopct=None, shadow=False, startangle=90)

plt.axis('equal')
plt.title('Memberships', pad=20, fontsize=15)

plt.legend(legendLabels, loc='best', fontsize=8)

fig4.savefig(workDirectory+'myplot4.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot4.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot4.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot4.png')
img.anchor = 'E4'

workbook['Memberships'].add_image(img)
workbook.save(outputExcelFile)


# CHART COUNT REGISTRATIONS BY DATE (FIELD Created at)
chartLabel = df_Created_count['Created'].tolist()
chartValue = df_Created_count['Total'].tolist()

fig5 = plt.figure(figsize=(13,6))
bar_plot = plt.bar(chartLabel, chartValue)

# plt.ylabel('yyy')
# plt.xlabel('xxx')
plt.xticks(rotation=30, ha='right')

# HIDE BORDERS
plt.gca().spines['left'].set_color('none')
plt.gca().spines['right'].set_color('none')
plt.gca().spines['top'].set_color('none')

# HIDE TICKS
plt.tick_params(axis='y', labelsize=0, length=0)
plt.yticks([])

# ADD VALUE ON THE END OF HORIZONTAL BARS
# for index, value in enumerate(chartValue):
#     plt.text(value, index, str(value))

# ADD VALUE ON THE TOP OF VERTICAL BARS
def autolabel(rects):
    for idx, rect in enumerate(bar_plot):
        height = rect.get_height()
        plt.text(rect.get_x() + rect.get_width()/2, height,
                chartValue[idx],
                ha='center', va='bottom', rotation=0)

autolabel(bar_plot)

plt.title('Registrations by month', pad=20, fontsize=15)

fig5.savefig(workDirectory+'myplot5.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot5.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot5.png')

# INSERT IN EXCEL
img = openpyxl.drawing.image.Image(workDirectory+'myplot5.png')
img.anchor = 'F2'

workbook['Registrations'].add_image(img)
workbook.save(outputExcelFile)


# REMOVE PICTURES
os.remove(workDirectory+'myplot1.png')
os.remove(workDirectory+'myplot2.png')
os.remove(workDirectory+'myplot3.png')
os.remove(workDirectory+'myplot4.png')
os.remove(workDirectory+'myplot5.png')


# TERMINAL OUTPUTS
# print(tab(df_Country_count.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_Categories_count, headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_Specialties_count.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_SpecialtiesPerCountry_count.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_Industries_count.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_Email_DNS_count.head(10), headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_HowDidYouHearAboutUs_count, headers='keys', tablefmt='psql', showindex=False))
# print(tab(df_Membership_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_ActivationCount, headers='keys', tablefmt='psql', showindex=False))
print(today)
print("OK, export done!")
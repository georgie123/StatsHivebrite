
from datetime import date
import pandas as pd
from tabulate import tabulate as tab

today = date.today()

# For now from an Excel import, later we will use the API
df = pd.read_excel('C:/Users/Georges/Downloads/User_export_'+str(today)+'.xlsx',
                   sheet_name='Export', engine='openpyxl',
                   usecols=['Email', 'Live Location:Country', 'Industries:Industries',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_',
                            'Last Membership:Type name'
                            ])

# COUNT COUNTRY
df_Country_count = pd.DataFrame(df.groupby(['Live Location:Country'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Country_count['Percent'] = (df_Country_count['Total'] / df_Country_count['Total'].sum()) * 100

# COUNT CATEGORIES (CUSTOM FIELD _8f70fe1e_Occupation)
df['Categories'] = df['_8f70fe1e_Occupation'].str.split(': ').str[0]
df_Categories_count = pd.DataFrame(df.groupby(['Categories'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT SPECIALTIES (CUSTOM FIELD _8f70fe1e_Occupation)
df['Specialties'] = df['_8f70fe1e_Occupation'].str.split(': ').str[1]
df_Specialties_count = pd.DataFrame(df.groupby(['Specialties'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT SPECIALTIES PER COUNTRY
df_SpecialtiesPerCountry_count = pd.DataFrame(df.groupby(['Live Location:Country', 'Specialties'])\
    .size(), columns=['Total']).sort_values(['Live Location:Country', 'Total'], ascending=[True, False]).reset_index()

# COUNT EXPERTISE & INTERESTS (CUSTOM FIELD Industries:Industries)
df_tempIndustries = pd.DataFrame(pd.melt(df['Industries:Industries'].str.split(',', expand=True))['value'])
df_Industries_count = pd.DataFrame(df_tempIndustries.groupby(['value']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT EMAIL DOMAINS
df['Domain'] = df['Email'].str.split('@').str[1]
df_Email_DNS_count = pd.DataFrame(df.groupby(['Domain']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT HOW DID YOU HEAR ABOUT US (CUSTOM FIELD How_did_you_hear_about_us_)
df_HowDidYouHearAboutUs_count = pd.DataFrame(df.groupby(['_ed5be3a0_How_did_you_hear_about_us_']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# COUNT MEMBERSHIP (FIELD Last Membership:Type name)
df_Membership_count = pd.DataFrame(df.groupby(['Last Membership:Type name']).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()

# EXCEL FILE
writer = pd.ExcelWriter('C:/Users/Georges/Downloads/'+str(today)+' Stats Hivebrite.xlsx', engine='xlsxwriter')

df_Country_count.to_excel(writer, index=False, sheet_name='Countries', header=['Country', 'Total', '%'])
df_Categories_count.to_excel(writer, index=False, sheet_name='Categories', header=['Category', 'Total'])
df_Specialties_count.to_excel(writer, index=False, sheet_name='Specialties', header=['Specialty', 'Total'])
df_SpecialtiesPerCountry_count.to_excel(writer, index=False, sheet_name='Specialties per country', header=['Country', 'Specialty', 'Total'])
df_Industries_count.to_excel(writer, index=False, sheet_name='Expertise & Interests', header=['Expertise or Interest', 'Total'])
df_Email_DNS_count.to_excel(writer, index=False, sheet_name='Email domains', header=['Email domain', 'Total'])
df_HowDidYouHearAboutUs_count.to_excel(writer, index=False, sheet_name='How Did You Hear', header=['How did you hear about us', 'Total'])
df_Membership_count.to_excel(writer, index=False, sheet_name='Memberships', header=['Membership', 'Total'])

writer.save()

# TERMINAL OUTPUTS
print(tab(df_Country_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Categories_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Specialties_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_SpecialtiesPerCountry_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Industries_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Email_DNS_count.head(10), headers='keys', tablefmt='psql', showindex=False))
print(tab(df_HowDidYouHearAboutUs_count, headers='keys', tablefmt='psql', showindex=False))
print(tab(df_Membership_count, headers='keys', tablefmt='psql', showindex=False))
print(today)
print("OK, export done!")
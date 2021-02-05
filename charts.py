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
                   usecols=['ID', 'Email', 'Account activation date', 'Live Location:Country', 'Industries:Industries',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_', 'Last Membership:Type name'
                            ])


# COUNT MEMBERSHIP (FIELD Last Membership:Type name)
df_Membership_count = pd.DataFrame(df.groupby(['Last Membership:Type name'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Total'], ascending=False).reset_index()
df_Membership_count = df_Membership_count.fillna('Basic Membership')

df_Membership_count['Percent'] = (df_Membership_count['Total'] / df_Membership_count['Total'].sum()) * 100
df_Membership_count['Percent'] = df_Membership_count['Percent'].round(decimals=1)


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
plt.show()
plt.clf()

im = Image.open(workDirectory+'myplot4.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot4.png')



os.remove(workDirectory+'myplot4.png')


print(tab(df_HowDidYouHearAboutUs_count, headers='keys', tablefmt='psql', showindex=False))

# INSERT IN EXCEL
# img = openpyxl.drawing.image.Image(workDirectory+'myplot2.png')
# img.anchor = 'E6'
#
# workbook['Categories'].add_image(img)
# workbook.save(outputExcelFile)

import os
from datetime import date
from tabulate import tabulate as tab

import pandas as pd

import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib.patches import Polygon
from matplotlib.collections import PatchCollection
from mpl_toolkits.basemap import Basemap

import numpy as np

from PIL import Image, ImageOps

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

today = date.today()

shp_simple_countries = r'C:/Users/Georges/PycharmProjects/data/simple_countries/simple_countries'
workDirectory = r'C:/Users/Georges/Downloads/'
outputExcelFile = workDirectory+str(today)+' Stats AMS Users.xlsx'

# For now from an Excel import, later we will use the API
inputExcelFile = workDirectory+'User_export_'+str(today)+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='Export', engine='openpyxl',
                   usecols=['ID', 'Email', 'Not blocked', 'Created at', 'Account activation date', 'Live Location:Country',
                            'Industries:Industries', 'Groups Member:Group Member',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_', 'Last Membership:Type name'
                            ])


# COUNT GROUPS (FIELD Groups Member:Group Member)
df_tempGroups = pd.DataFrame(pd.melt(df['Groups Member:Group Member'].str.split(',', expand=True))['value'])
df_Groups_count = pd.DataFrame(df_tempGroups.groupby(['value'], dropna=False).size(), columns=['Total'])\
    .reset_index()
df_Groups_count = df_Groups_count.fillna('AZERTY')

df_Groups_count['Percent'] = (df_Groups_count['Total'] / df.shape[0]) * 100
df_Groups_count['Percent'] = df_Groups_count['Percent'].round(decimals=2)

# EMPTY VALUES
groupsEmpty = df['Groups Member:Group Member'].isna().sum()
groupsEmptyPercent = round((groupsEmpty / df.shape[0]) * 100, 2)

# REPLACE EMPTY VALUES AND SORT
df_Groups_count.loc[(df_Groups_count['value'] == 'AZERTY')] = [['None', groupsEmpty, groupsEmptyPercent]]
df_Groups_count = df_Groups_count.sort_values(['Total'], ascending=False)

df_Groups_count['value'] = df_Groups_count['value'].replace(['17794'], 'AMS North American Chapter')
df_Groups_count['value'] = df_Groups_count['value'].replace(['19659'], 'No name')
df_Groups_count['value'] = df_Groups_count['value'].replace(['19858'], 'Industries')
df_Groups_count['value'] = df_Groups_count['value'].replace(['19859'], 'Medicinae Doctor')
df_Groups_count['value'] = df_Groups_count['value'].replace(['22580'], 'Euro Aesthetics')
df_Groups_count['value'] = df_Groups_count['value'].replace(['23831'], 'AMS Eastern Europe')


# CHART GROUPS (FIELD Groups Member:Group Member)
chartLabel = df_Groups_count['value'].tolist()
chartValue = df_Groups_count['Total'].tolist()
chartLegendPercent = df_Groups_count['Percent'].tolist()

legendLabels = []
for i, j in zip(chartLabel, map(str, chartLegendPercent)):
    legendLabels.append(i + ' (' + j + ' %)')

colors = plt.rcParams['axes.prop_cycle'].by_key()['color']

fig6 = plt.figure()
plt.pie(chartValue, labels=None, colors=colors, autopct=None, shadow=False, startangle=90)

plt.axis('equal')
plt.title('Groups', pad=20, fontsize=15)

plt.legend(legendLabels, loc='best', fontsize=8)

fig6.savefig(workDirectory+'myplot6.png', dpi=100)
plt.clf()

im = Image.open(workDirectory+'myplot6.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot6.png')


print(tab(df_Groups_count, headers='keys', tablefmt='psql', showindex=False))
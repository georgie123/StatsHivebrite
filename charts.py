import os
from datetime import date

from tabulate import tabulate as tab

import pandas as pd
import matplotlib.pyplot as plt
from PIL import Image, ImageOps

today = date.today()

workDirectory = r'C:/Users/Georges/Downloads/'

outputExcelFile = workDirectory+str(today)+' Stats AMS.xlsx'

# For now from an Excel import, later we will use the API
inputExcelFile = workDirectory+'User_export_'+str(today)+'.xlsx'
df = pd.read_excel(inputExcelFile, sheet_name='Export', engine='openpyxl',
                   usecols=['ID', 'Email', 'Created at', 'Account activation date', 'Live Location:Country', 'Industries:Industries',
                            '_8f70fe1e_Occupation', '_ed5be3a0_How_did_you_hear_about_us_', 'Last Membership:Type name'
                            ])


# COUNT REGISTRATIONS BY DATE (FIELD Created at)
df['Created'] = pd.to_datetime(df['Created at'])
df['Created'] = df['Created'].dt.to_period("M")
df_TempCreated = pd.DataFrame(df['Created'])

df_Created_count = pd.DataFrame(df_TempCreated.groupby(['Created'], dropna=False).size(), columns=['Total'])\
    .sort_values(['Created'], ascending=True).reset_index()

df_Created_count['Created'] = df_Created_count['Created'].dt.strftime('%b %Y')

ind_drop = df_Created_count[df_Created_count['Created'].apply(lambda x: x.startswith('Feb 2020'))].index
df_Created_count = df_Created_count.drop(ind_drop)


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
plt.show()
plt.clf()

im = Image.open(workDirectory+'myplot5.png')
bordered = ImageOps.expand(im, border=1, fill=(0, 0, 0))
bordered.save(workDirectory+'myplot5.png')



os.remove(workDirectory+'myplot5.png')


print(tab(df_Created_count.head(30), headers='keys', tablefmt='psql', showindex=False))

# INSERT IN EXCEL
# img = openpyxl.drawing.image.Image(workDirectory+'myplot2.png')
# img.anchor = 'E6'
#
# workbook['Categories'].add_image(img)
# workbook.save(outputExcelFile)
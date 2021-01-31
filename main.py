
import pandas as pd
from tabulate import tabulate as tab

# For now from an Excel import, later we will use the API
df = pd.read_excel('C:/Users/Georges/Downloads/User_export_2021-01-28.xlsx',
                   sheet_name='Export', engine='openpyxl', header=None)

print(tab(df.head(10), headers='firstrow', tablefmt='psql'))



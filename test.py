from datetime import date

import pandas as pd
from tabulate import tabulate as tab

from main import (df)

today = date.today()



print(tab(df_Industries_count.head(30), headers='keys', tablefmt='psql', showindex=False))

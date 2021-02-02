from datetime import date
import openpyxl
from openpyxl.utils import get_column_letter

today = date.today()

outputExcelFile = r'C:/Users/Georges/Downloads/'+str(today)+' Stats Hivebrite.xlsx'
workbook = openpyxl.load_workbook(outputExcelFile)
sheetsLits = workbook.sheetnames


# EXCEL COLUMN SIZE
for sheet in sheetsLits:
    for cell in workbook[sheet][1]:
        if get_column_letter(cell.column) == 'A':
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 30
        else:
            workbook[sheet].column_dimensions[get_column_letter(cell.column)].width = 10

        workbook.save(outputExcelFile)



# ws.column_dimensions['A'].width = 75
# wb.save("C:/Users/Georges/Downloads/2021-02-02 Stats Hivebrite.xlsx")
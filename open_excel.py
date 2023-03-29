import openpyxl
import win32com.client as win32
from pathlib import Path

mzdy = Path('C:/Users/Uzivatel/Desktop/mzdy')
xlsxfile = 'temp-Fiala-up.xlsx'
xfile = Path(mzdy / xlsxfile)

excel = win32.gencache.EnsureDispatch('Excel.Application')
excel.Visible = False

wb = excel.Workbooks.Open(xfile)
wb.Close(True)

excel.Application.Quit()

wb = openpyxl.load_workbook(mzdy / 'temp-fiala-up.xlsx', data_only=True)
ws = wb.worksheets[0]
print(ws.cell(21, 7).value)
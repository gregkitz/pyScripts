import json
import os
import xlsxwriter
from pprint import pprint

direc = r""
otherdirec = r""

workbook = xlsxwriter.Workbook('Queries.xlsx')
worksheet = workbook.add_worksheet()

worksheet.set_column(1, 1, 200)

directory = os.fsencode(otherdirec)
mm = 0
special = []

for file in os.listdir(directory):
    filename = os.fsdecode(os.path.join(directory, file))
    if filename.endswith(".json") or filename.endswith(".txt"):
        try:
            print(filename)
            with open(str(filename), encoding="utf-8-sig") as jsonfile:
                data = json.load(jsonfile)
                name = 
                connectionString = 
                commandText = 
                mm += 1
                cool = (name, connectionString, commandText)
                special.append(cool)
                pprint(cool)
            continue
        except Exception as e:
            print("Exception: " + str(e))

row = 1
col = 0

for name, connectionString, commandText in special:
    worksheet.write_string(row, col, name)
    worksheet.write_string(row, col+1, connectionString)
    worksheet.write_string(row, col+2, commandText)
    row += 1

workbook.close()
import win32.client as win32
excel = win32.gencache.EnsureDispatch('Excel.Application')
wb = excel.Workbooks.Open(r'file.xlsx')
ws = wb.Worksheets("Sheet1")
ws.Columns.AutoFit()
wb.Save()
excel.Application.Quit()

print(mm)

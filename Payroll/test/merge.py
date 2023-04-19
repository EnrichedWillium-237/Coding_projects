# Code for organizing payroll overtime calculation.
# Reads in output_raw.xlsx

# Source files
import openpyxl
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.styles import Alignment, Border, Side, Font
from datetime import datetime, date, timedelta

flagDebug = False

# Input file
workbook = load_workbook('output_raw.xlsx')
sheet = workbook.active
for i in range(1, 10000):
    val0 = sheet.cell(row = i, column = 1).value
    val1 = sheet.cell(row = i + 1, column = 1).value
    val2 = sheet.cell(row = i + 2, column = 1).value
    if val0 is None and val1 is None and val2 is None:
        continue
    Nrow = i

# Setup output file and spreadsheet
newbook1 = openpyxl.Workbook()
newsheet1 = newbook1.active

for i in range(8, Nrow + 1):
    val0 = sheet.cell(row = i - 1, column = 1).value
    val1 = sheet.cell(row = i, column = 1).value
    val2 = sheet.cell(row = i + 1, column = 1).value
    if val1 is not None:
        if val0 is None:
            valName = val1
            pcnt = 1
        if "---Total" in val1:
            c0 = newsheet1.cell(row = i, column = 1)
            c0.value = valName
            pcnt += 1
            c0 = newsheet1.cell(row = i + pcnt, column = 1)
            c0.value = valName

# Delete empty rows
indx = []
for i in range(len(tuple(newsheet1.rows))):
    flag = False
    for cell in tuple(newsheet1.rows)[i]:
        if cell.value != None:
            flag = True
            break
    if flag == False:
        indx.append(i)
indx.sort()
for i in range(len(indx)):
    newsheet1.delete_rows(idx = indx[i]+1-i)


output_name = "output_merged.xlsx"
newbook1.save(output_name)
print("\n")
print("file output written to", output_name)
print("\nSpreadsheet informarion merged.\n")

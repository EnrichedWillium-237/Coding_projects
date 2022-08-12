### Code for reading in a payroll spreadsheet and organizing net payments ###
### This is a work in progress and does things in a brute force fashion. Will be improved later. ###

# Source files
import openpyxl
import array as arr
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side
import numpy as np

# Payroll file location
workbook = load_workbook('payroll_input.xlsx')

# Payroll output
newbook1 = openpyxl.Workbook()
newsheet1 = newbook1.active

sheet = workbook.active

Nrow = sheet.max_row
Nrow = Nrow - 1 # offset from spreadsheet
Ncell = 1 # index for excel output

arrName = np.empty(500, dtype = 'object')
arrCat = np.empty(500, dtype = 'object')
arrHrs = np.empty(500, dtype = 'f')
arrRate = np.empty(500, dtype = 'f')
arrTot = np.empty(500, dtype = 'f')
arrReim = np.empty(500, dtype = 'f')
arrGross = np.empty(500, dtype = 'f')
arrNote = np.empty(500, dtype = 'object')

#print(arrName)

for i in range(2, Nrow):
#for i in range(2, 25):

    valName = sheet.cell(row = i, column = 1)
    valCat = sheet.cell(row = i, column = 2)
    valHrs = sheet.cell(row = i, column, = 3)
    valName = valName.value
    arrName[i] = valName

    if valName is None or 0:
        break

for i in range(2, Nrow):
    print(arrName[i])
    if arrName[i] is None or 0:
        break

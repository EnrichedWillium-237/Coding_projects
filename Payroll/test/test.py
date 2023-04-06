# Code for calculating OT values
# Download .csv file "Assigned Shift Details - General" with the following parameters:
# "Position Name", "Date", "Start Time", "End Time", "Duration", "Employee Name", "Employee Pay Rate".
# Save .csv file as "data.xlms" and put in same directory as code.

# Source files
import openpyxl
import array as arr
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Border, Side
from datetime import datetime
from datetime import date

# Week end dates (must be changed for each pay cycle)
import datetime
week1_end = datetime.date(2023, 3, 26)
week2_end = datetime.date(2023, 4, 2)

# File location
workbook = load_workbook('data.xlsx')

# Payroll output
#newbook1 = openpyxl.Workbook()
#newsheet1 = newbook1.active

sheet = workbook.active
label_1 = sheet.cell(row=1, column=1)
label_2 = sheet.cell(row=1, column=2)
label_3 = sheet.cell(row=1, column=3)
label_4 = sheet.cell(row=1, column=4)
label_5 = sheet.cell(row=1, column=5)
label_6 = sheet.cell(row=1, column=6)
label_7 = sheet.cell(row=1, column=7)
print(label_1.value,"\t",label_2.value,"\t",label_3.value,"\t",label_4.value,"\t",label_5.value,"\t",label_6.value,"\t",label_7.value,"\t")

Nrow = sheet.max_row # total number of rows
Ncell = 1

# Calculation for crosschecks
GrandHrs = 0 # total number of hours across all names and positions
for i in range(2, Nrow-1):
    valHrs = sheet.cell(row = i, column = 5)
    GrandHrs += valHrs.value
print("Total hours for all names and positions:  ", GrandHrs)

# Calculate number of shifts
numShifts = 0
for i in range(2, Nrow+1):
    valName = sheet.cell(row = i, column = 6)
    Name = valName.value
    valName = sheet.cell(row = i+1, column = 6)
    NameNxt = valName.value
    #print(Name)

# Convert date format
date_obj = sheet.cell(row = 2, column = 2)
shift_date = date_obj.value
#print(shift_date.day," ",shift_date.month)

week1Hrs = 0
week2Hrs = 0
OT40week1 = 0
OT40week2 = 0
OT40week1_Pos1 = 0
OT40week1_Pos2 = 0
OT40week1_Pos3 = 0
OT40week2_Pos1 = 0
OT40week2_Pos2 = 0
OT40week2_Pos3 = 0
OT40week1_Pos1_name = 0

rowmin = 270
rowmax = 280
rowmid = 0
valName = sheet.cell(row = rowmin, column = 6).value
for i in range(rowmin, rowmax+1):
    valDate = sheet.cell(row = i, column = 2)
    valHrs = sheet.cell(row = i, column = 5)
    valHrs = valHrs.value
    valday = valDate.value.day
    valmonth = valDate.value.month
    if valday <= week1_end.day and valmonth <= week1_end.month:
        week1Hrs += valHrs
    else:
        week2Hrs += valHrs
    if valday > week1_end.day and valmonth <= week1_end.month:
        if rowmid == 0:
            rowmid = i

if week1Hrs > 40:
    OT40week1 = week1Hrs - 40
    #for i in range(rowmin)
if week2Hrs > 40:
    OT40week2 = week2Hrs - 40
#for i in range(rowmax+1, rowmin, -1):

print(valName, "\ttotal OT+40 week 1: ", OT40week1, "\ttotal OT+40 week 2: ", OT40week2)

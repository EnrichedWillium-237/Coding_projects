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

def abs(x):
    return x if x>=0 else -x

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
RegHrsweek1_Pos1 = 0
RegHrsweek2_Pos1 = 0
OT40week1_Pos1 = 0
OT40week2_Pos1 = 0
OT12_pos1 = 0

rowmin = 167
rowmax = 176
# rowmin = 270
# rowmax = 280
# rowmin = 678
# rowmax = 687
rowmid = 0 # find end of week one
valName = sheet.cell(row = rowmin, column = 6).value
for i in range(rowmin, rowmax+1):
    valPos = sheet.cell(row = i, column = 1)
    valDate = sheet.cell(row = i, column = 2)
    valHrs = sheet.cell(row = i, column = 5)
    valName = sheet.cell(row = i, column = 6)
    valPos = valPos.value
    valHrs = valHrs.value
    valday = valDate.value.day
    valmonth = valDate.value.month
    valName = valName.value
    if valday <= week1_end.day and valmonth <= week1_end.month:
        week1Hrs += valHrs
    else:
        week2Hrs += valHrs
    if valday > week1_end.day and valmonth <= week1_end.month:
        if rowmid == 0:
            rowmid = i
# OT +40 for week one
if week1Hrs > 40:
    OT40week1 = week1Hrs - 40
    OT = OT40week1
    OTn = OT
    OTcnt = 0
    flag1 = False
    for i in range(rowmid-1, rowmin-1, -1): # minus one offset in for loop because we're counting backwards
        valPos = sheet.cell(row = i, column = 1)
        valHrs = sheet.cell(row = i, column = 5)
        valName = sheet.cell(row = i, column = 6)
        valPos = valPos.value
        valHrs = valHrs.value
        valName = valName.value
        x = valHrs
        if flag1 is True:
            OTn = 0
            y = 0
        if flag1 is False:
            OTn = OTn - x
            if OTn > 0:
                y = x
                OTcnt += y
            if OTn < 0:
                y = OT - OTcnt
                flag1 = True
        if valHrs > 12: # OT +12 calculation
            z = valHrs - 12
            if y > z: z = 0
            else: z = z - y
        else: z = 0
        RegHrsweek1_Pos1 = valHrs - y - z
        OT40week1_Pos1 = y
        OT12_pos1 = z
        print("  valHrs",valHrs,"  Standard: ",RegHrsweek1_Pos1,"  OT+12:",OT12_pos1,"  OT+40: ",OT40week1_Pos1)
print(valName,"  Week 1 --- total: ", week1Hrs," shift+40 total: ",OT40week1)
# OT +40 for week two
if week2Hrs > 40:
    OT40week2 = week2Hrs - 40
    OT = OT40week2
    OTn = OT
    OTcnt = 0
    flag1 = False
    for i in range(rowmax, rowmid-1, -1):
        valPos = sheet.cell(row = i, column = 1)
        valHrs = sheet.cell(row = i, column = 5)
        valName = sheet.cell(row = i, column = 6)
        valPos = valPos.value
        valHrs = valHrs.value
        valName = valName.value
        x = valHrs
        if valHrs > 12: OT12_pos1 = valHrs - 12 # OT +12 calculation
        else: OT12_pos1 = 0
        if flag1 is True:
            OTn = 0
            y = 0
        if flag1 is False:
            OTn = OTn - x
            if OTn > 0:
                y = x
                OTcnt += y
            if OTn < 0:
                y = OT - OTcnt
                flag1 = True
        if valHrs > 12: # OT +12 calculation
            z = valHrs - 12
            if y > z: z = 0
            else: z = z - y
        else: z = 0
        RegHrsweek2_Pos1 = valHrs - y - z
        OT40week2_Pos1 = y
        OT12_pos1 = z
        print("  valHrs",valHrs,"  Standard: ",RegHrsweek2_Pos1,"  OT+12:",OT12_pos1,"  OT+40: ",OT40week2_Pos1)
print(valName,"  Week 2 --- total: ", week2Hrs," shift+40 total: ",OT40week2)

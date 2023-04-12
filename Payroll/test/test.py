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
flagDebug = True

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
Nrow = sheet.max_row # total number of rows
list1 = ["", "", 0, 0, 0, 0]
list2 = ["", "", 0, 0, 0, 0]
list3 = ["", "", 0, 0, 0, 0]
list4 = ["", "", 0, 0, 0, 0]
list5 = ["", "", 0, 0, 0, 0]
list6 = ["", "", 0, 0, 0, 0]
list7 = ["", "", 0, 0, 0, 0]
list8 = ["", "", 0, 0, 0, 0]
list9 = ["", "", 0, 0, 0, 0]

# Calculation for crosschecks
GrandHrs = 0 # total number of hours across all names and positions
for i in range(2, Nrow-1):
    valHrs = sheet.cell(row = i, column = 5)
    GrandHrs += valHrs.value
print("\n")
print("Total hours for all names and positions:  ", f'{GrandHrs:.9}')
print("\n")

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
    RegHrsweek1 = 0
    OT40week1 = 0
    OT12 = 0
    OT40week1 = week1Hrs - 40
    OT = OT40week1
    OTn = OT
    OTcnt = 0
    flag1 = False
    valName = sheet.cell(row = rowmid-1, column = 6)
    valName = valName.value
    listWeek1 = [list1, list2, list3, list4, list5, list6, list7, list8, list9]
    for i in range(rowmid-1, rowmin-1, -1): # minus one offset in for loop because we're counting backwards
        valPos = sheet.cell(row = i, column = 1)
        valHrs = sheet.cell(row = i, column = 5)
        valPos = valPos.value
        valHrs = valHrs.value
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
        RegHrsweek1 = valHrs - y - z
        OT12 = z
        if i == rowmid - 1:   list1 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 2: list2 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 3: list3 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 4: list4 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 5: list5 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 6: list6 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 7: list7 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 8: list8 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 9: list9 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        if i <= rowmid - 10: print("Employee has worked too many shifts in one week!!! Give them some time off!!!")
        if flagDebug: print(valName, "  ", valPos, "  Shift", valHrs, "  Standard: ", RegHrsweek1, "  OT+12:", OT12, "  OT+40: ", y)
    # Sort by position
    pos1 = list1[1]
    pos2 = list2[1]
    pos3 = list3[1]
    pos4 = list4[1]
    pos5 = list5[1]
    pos6 = list6[1]
    pos7 = list7[1]
    pos8 = list8[1]
    pos9 = list9[1]
    hrsPos1 = list1[2]
    RegHrs1 = list1[3]
    totOT12_1 = list1[4]
    totOT40_1 = list1[5]
    if flagDebug: print("\n")
    if flagDebug: print(valName,"  Position: \ttotal hrs:  reg hrs: \tOT+12: \tOT+40")
    # Position 1
    if (    pos2 is not None and pos2 in pos1):
        hrsPos1 += list2[2]
        RegHrs1 += list2[3]
        totOT12_1 += list2[4]
        totOT40_1 += list2[5]
    if pos3 is not None and pos3 in pos1:
        hrsPos1 += list3[2]
        RegHrs1 += list3[3]
        totOT12_1 += list3[4]
        totOT40_1 += list3[5]
    if pos4 is not None and pos4 in pos1:
        hrsPos1 += list4[2]
        RegHrs1 += list4[3]
        totOT12_1 += list4[4]
        totOT40_1 += list4[5]
    if pos5 is not None and pos5 in pos1:
        hrsPos1 += list5[2]
        RegHrs1 += list5[3]
        totOT12_1 += list5[4]
        totOT40_1 += list5[5]
    if pos6 is not None and pos6 in pos1:
        hrsPos1 += list6[2]
        RegHrs1 += list6[3]
        totOT12_1 += list6[4]
        totOT40_1 += list6[5]
    if pos7 is not None and pos7 in pos1:
        hrsPos1 += list7[2]
        RegHrs1 += list7[3]
        totOT12_1 += list7[4]
        totOT40_1 += list7[5]
    if pos8 is not None and pos8 in pos1:
        hrsPos1 += list8[2]
        RegHrs1 += list8[3]
        totOT12_1 += list8[4]
        totOT40_1 += list8[5]
    if pos9 is not None and pos9 in pos1:
        hrsPos1 += list9[2]
        RegHrs1 += list9[3]
        totOT12_1 += list9[4]
        totOT40_1 += list9[5]
    if flagDebug: print("  ",pos1,"  ",hrsPos1,"   ",RegHrs1,"   ",totOT12_1,"   ",totOT40_1)
    list1 = [valName, pos1, hrsPos1, RegHrs1, totOT12_1, totOT40_1]
    # Position 2
    hrsPos2 = 0
    RegHrs2 = 0
    totOT12_2 = 0
    totOT40_2 = 0
    if (    pos2 is not None and pos2 not in pos1):
        hrsPos2 = list2[2]
        RegHrs2 = list2[3]
        totOT12_2 = list2[4]
        totOT40_2 = list2[5]
        if pos3 is not None and pos3 in pos2:
            hrsPos2 += list3[2]
            RegHrs2 += list3[3]
            totOT12_2 += list3[4]
            totOT40_2 += list3[5]
        if pos4 is not None and pos4 in pos2:
            hrsPos2 += list4[2]
            RegHrs2 += list4[3]
            totOT12_2 += list4[4]
            totOT40_2 += list4[5]
        if pos5 is not None and pos5 in pos2:
            hrsPos2 += list5[2]
            RegHrs2 += list5[3]
            totOT12_2 += list5[4]
            totOT40_2 += list5[5]
        if pos6 is not None and pos6 in pos2:
            hrsPos2 += list6[2]
            RegHrs2 += list6[3]
            totOT12_2 += list6[4]
            totOT40_2 += list6[5]
        if pos7 is not None and pos7 in pos2:
            hrsPos2 += list7[2]
            RegHrs2 += list7[3]
            totOT12_2 += list7[4]
            totOT40_2 += list7[5]
        if pos8 is not None and pos8 in pos2:
            hrsPos2 += list8[2]
            RegHrs2 += list8[3]
            totOT12_2 += list8[4]
            totOT40_2 += list8[5]
        if pos9 is not None and pos9 in pos2:
            hrsPos2 += list9[2]
            RegHrs2 += list9[3]
            totOT12_2 += list9[4]
            totOT40_2 += list9[5]
        if flagDebug: print("  ",pos2,"  ",hrsPos2,"   ",RegHrs2,"   ",totOT12_2,"   ",totOT40_2)
        list2 = [valName, pos2, hrsPos2, RegHrs2, totOT12_2, totOT40_2]
    # Position 3
    hrsPos3 = 0
    RegHrs3 = 0
    totOT12_3 = 0
    totOT40_3 = 0
    if (    pos3 is not None and pos3 not in pos1 and pos3 not in pos2):
        hrsPos3 += list3[2]
        RegHrs3 += list3[3]
        totOT12_3 += list3[4]
        totOT40_3 += list3[5]
        if pos4 is not None and pos4 in pos3:
            hrsPos3 += list4[2]
            RegHrs3 += list4[3]
            totOT12_3 += list4[4]
            totOT40_3 += list4[5]
        if pos5 is not None and pos5 in pos3:
            hrsPos3 += list5[2]
            RegHrs3 += list5[3]
            totOT12_3 += list5[4]
            totOT40_3 += list5[5]
        if pos6 is not None and pos6 in pos3:
            hrsPos3 += list6[2]
            RegHrs3 += list6[3]
            totOT12_3 += list6[4]
            totOT40_3 += list6[5]
        if pos7 is not None and pos7 in pos3:
            hrsPos3 += list7[2]
            RegHrs3 += list7[3]
            totOT12_3 += list7[4]
            totOT40_3 += list7[5]
        if pos8 is not None and pos8 in pos3:
            hrsPos3 += list8[2]
            RegHrs3 += list8[3]
            totOT12_3 += list8[4]
            totOT40_3 += list8[5]
        if pos9 is not None and pos9 in pos3:
            hrsPos3 += list9[2]
            RegHrs3 += list9[3]
            totOT12_3 += list9[4]
            totOT40_3 += list9[5]
        if flagDebug: print("  ",pos3,"  ",hrsPos3,"   ",RegHrs3,"   ",totOT12_3,"   ",totOT40_3)
        list3 = [valName, pos3, hrsPos3, RegHrs3, totOT12_3, totOT40_3]
    # Position 4
    hrsPos4 = 0
    RegHrs4 = 0
    totOT12_4 = 0
    totOT40_4 = 0
    if (    pos4 is not None and pos4 not in pos1 and pos4 not in pos2 and pos4 not in pos3):
        hrsPos4 += list4[2]
        RegHrs4 += list4[3]
        totOT12_4 += list4[4]
        totOT40_4 += list4[5]
        if pos5 is not None and pos5 in pos4:
            hrsPos4 += list5[2]
            RegHrs4 += list5[3]
            totOT12_4 += list5[4]
            totOT40_4 += list5[5]
        if pos6 is not None and pos6 in pos4:
            hrsPos4 += list6[2]
            RegHrs4 += list6[3]
            totOT12_4 += list6[4]
            totOT40_4 += list6[5]
        if pos7 is not None and pos7 in pos4:
            hrsPos4 += list7[2]
            RegHrs4 += list7[3]
            totOT12_4 += list7[4]
            totOT40_4 += list7[5]
        if pos8 is not None and pos8 in pos4:
            hrsPos4 += list8[2]
            RegHrs4 += list8[3]
            totOT12_4 += list8[4]
            totOT40_4 += list8[5]
        if pos9 is not None and pos9 in pos4:
            hrsPos4 += list9[2]
            RegHrs4 += list9[3]
            totOT12_4 += list9[4]
            totOT40_4 += list9[5]
        if flagDebug: print("  ",pos4,"  ",hrsPos4,"   ",RegHrs4,"   ",totOT12_4,"   ",totOT40_4)
        list4 = [valName, pos4, hrsPos4, RegHrs4, totOT12_4, totOT40_4]
    # Position 5
    hrsPos5 = 0
    RegHrs5 = 0
    totOT12_5 = 0
    totOT40_5 = 0
    if (    pos5 is not None and pos5 not in pos1 and pos5 not in pos2 and pos5 not in pos3 and
            pos5 not in pos4):
        hrsPos5 += list5[2]
        RegHrs5 += list5[3]
        totOT12_5 += list5[4]
        totOT40_5 += list5[5]
        if pos6 is not None and pos6 in pos5:
            hrsPos5 += list6[2]
            RegHrs5 += list6[3]
            totOT12_5 += list6[4]
            totOT40_5 += list6[5]
        if pos7 is not None and pos7 in pos5:
            hrsPos5 += list7[2]
            RegHrs5 += list7[3]
            totOT12_5 += list7[4]
            totOT40_5 += list7[5]
        if pos8 is not None and pos8 in pos5:
            hrsPos5 += list8[2]
            RegHrs5 += list8[3]
            totOT12_5 += list8[4]
            totOT40_5 += list8[5]
        if pos9 is not None and pos9 in pos5:
            hrsPos5 += list9[2]
            RegHrs5 += list9[3]
            totOT12_5 += list9[4]
            totOT40_5 += list9[5]
        if flagDebug: print("  ",pos5,"  ",hrsPos5,"   ",RegHrs5,"   ",totOT12_5,"   ",totOT40_5)
        list5 = [valName, pos5, hrsPos5, RegHrs5, totOT12_5, totOT40_5]
    # Position 6
    hrsPos6 = 0
    RegHrs6 = 0
    totOT12_6 = 0
    totOT40_6 = 0
    if (    pos6 is not None and pos6 not in pos1 and pos6 not in pos2 and pos6 not in pos3 and
            pos6 not in pos4 and pos6 not in pos5):
        hrsPos6 += list6[2]
        RegHrs6 += list6[3]
        totOT12_6 += list6[4]
        totOT40_6 += list6[5]
        if pos7 is not None and pos7 in pos6:
            hrsPos6 += list7[2]
            RegHrs6 += list7[3]
            totOT12_6 += list7[4]
            totOT40_6 += list7[5]
        if pos8 is not None and pos8 in pos6:
            hrsPos6 += list8[2]
            RegHrs6 += list8[3]
            totOT12_6 += list8[4]
            totOT40_6 += list8[5]
        if pos9 is not None and pos9 in pos6:
            hrsPos6 += list9[2]
            RegHrs6 += list9[3]
            totOT12_6 += list9[4]
            totOT40_6 += list9[5]
        if flagDebug: print("  ",pos6,"  ",hrsPos6,"   ",RegHrs6,"   ",totOT12_6,"   ",totOT40_6)
        list6 = [valName, pos6, hrsPos6, RegHrs6, totOT12_6, totOT40_6]
    # Position 7
    hrsPos7 = 0
    RegHrs7 = 0
    totOT12_7 = 0
    totOT40_7 = 0
    if (    pos7 is not None and pos7 not in pos1 and pos7 not in pos2 and pos7 not in pos3 and
            pos7 not in pos4 and pos7 not in pos5 and pos7 not in pos6):
        hrsPos7 += list7[2]
        RegHrs7 += list7[3]
        totOT12_7 += list7[4]
        totOT40_7 += list7[5]
        if pos8 is not None and pos8 in pos7:
            hrsPos7 += list8[2]
            RegHrs7 += list8[3]
            totOT12_7 += list8[4]
            totOT40_7 += list8[5]
        if pos9 is not None and pos9 in pos7:
            hrsPos7 += list9[2]
            RegHrs7 += list9[3]
            totOT12_7 += list9[4]
            totOT40_7 += list9[5]
        if flagDebug: print("  ",pos7,"  ",hrsPos7,"   ",RegHrs7,"   ",totOT12_7,"   ",totOT40_7)
        list7 = [valName, pos7, hrsPos7, RegHrs7, totOT12_7, totOT40_7]
    # Position 8
    hrsPos8 = 0
    RegHrs8 = 0
    totOT12_8 = 0
    totOT40_8 = 0
    if (    pos8 is not None and pos8 not in pos1 and pos8 not in pos2 and pos8 not in pos3 and
            pos8 not in pos4 and pos8 not in pos5 and pos8 not in pos6 and pos8 not in pos7):
        hrsPos8 += list8[2]
        RegHrs8 += list8[3]
        totOT12_8 += list8[4]
        totOT40_8 += list8[5]
        if pos9 is not None and pos9 in pos8:
            hrsPos8 += list9[2]
            RegHrs8 += list9[3]
            totOT12_8 += list9[4]
            totOT40_8 += list9[5]
        if flagDebug: print("  ",pos8,"  ",hrsPos8,"   ",RegHrs8,"   ",totOT12_8,"   ",totOT40_8)
        list8 = [valName, pos8, hrsPos8, RegHrs8, totOT12_8, totOT40_8]
    # Position 9
    hrsPos9 = 0
    RegHrs9 = 0
    totOT12_9 = 0
    totOT40_9 = 0
    if (    pos9 is not None and pos9 not in pos1 and pos9 not in pos2 and pos9 not in pos3 and
            pos9 not in pos4 and pos9 not in pos5 and pos9 not in pos6 and pos9 not in pos7 and
            pos9 not in pos8):
        hrsPos9 += list9[2]
        RegHrs9 += list9[3]
        totOT12_9 += list9[4]
        totOT40_9 += list9[5]
        if flagDebug: print("  ",pos9,"  ",hrsPos9,"   ",RegHrs9,"   ",totOT12_9,"   ",totOT40_9)
        list9 = [valName, pos9, hrsPos9, RegHrs9, totOT12_9, totOT40_9]
print(valName,"  Week 1 --- total: ", week1Hrs," shift+40 total: ",OT40week1)
if flagDebug: print("\n")

# OT +40 for week two
if week2Hrs > 40:
    RegHrsweek2 = 0
    OT40week2 = 0
    OT12 = 0
    OT40week2 = week2Hrs - 40
    OT = OT40week2
    OTn = OT
    OTcnt = 0
    flag1 = False
    listWeek2 = [list1, list2, list3, list4, list5, list6, list7, list8, list9]
    for i in range(rowmax, rowmid-1, -1):
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
        RegHrsweek2 = valHrs - y - z
        OT12 = z
        if i == rowmid - 1:   list1 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 2: list2 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 3: list3 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 4: list4 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 5: list5 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 6: list6 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 7: list7 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 8: list8 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        elif i == rowmid - 9: list9 = [valName, valPos, valHrs, RegHrsweek1, OT12, y]
        if i <= rowmid - 10: print("Employee has worked too many shifts in one week!!! Give them some time off!!!")
        print(valName, "  ", valPos, "  Shift", valHrs, "  Standard: ", RegHrsweek2, "  OT+12:", OT12, "  OT+40: ", y)
print(valName,"  Week 2 --- total: ", week2Hrs," shift+40 total: ",OT40week2)

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
    if val1 is not None:
        if val0 is None:
            valName = val1
        if "---Total" in val1:
            cnt = 0
            for j in range(0, 9):
                pos = sheet.cell(row = i + 1 + j, column = 1).value
                reg = sheet.cell(row = i + 1 + j, column = 3).value
                ot  = sheet.cell(row = i + 1 + j, column = 4).value
                flag12 = False
                fix12 = sheet.cell(row = i + 1 + j, column = 7).value
                if fix12 is not None and fix12.__contains__("Check"): flag12 = True
                if pos is None: break
                else:
                    c0 = newsheet1.cell(row = i + 1 + cnt, column = 2)
                    c0.value = valName
                    c0 = newsheet1.cell(row = i + 1 + cnt, column = 3)
                    c0.value = pos
                    if pos.__contains__("ARMED") and ot == 0: cat = "Armed"
                    elif pos.__contains__("Admin Work"): cat = "Admin"
                    elif pos.__contains__("Training"): cat = "Training"
                    elif pos.__contains__("Sick"): cat = "Sick"
                    elif pos.__contains__("Covid") or pos.__contains__("COVID"): cat = "Covid"
                    else: cat = "Unarmed"
                    if flag12 is True:
                        c2 = newsheet1.cell(row = i + 1 + cnt, column = 11)
                        c2.value = "HAND CALCULATE OT+12"
                        c2.font = Font(bold = 'single')
                    if reg != 0 and ot == 0:
                        c0 = newsheet1.cell(row = i + 1 + cnt, column = 6)
                        c0.value = reg
                        c1 = newsheet1.cell(row = i + 1 + cnt, column = 4)
                        c1.value = cat
                        cnt += 1
                    if reg == 0 and ot != 0:
                        c0 = newsheet1.cell(row = i + 1 + cnt, column = 6)
                        c0.value = ot
                        if cat.__contains__("Unarmed"): cat = "OT Unarmed"
                        if cat.__contains__("Armed"): cat = "OT Armed"
                        if cat.__contains__("Admin"): cat = "OT Admin"
                        c1 = newsheet1.cell(row = i + 1 + cnt, column = 4)
                        c1.value = cat
                        cnt += 1
                    if reg != 0 and ot != 0:
                        c0 = newsheet1.cell(row = i + 1 + cnt, column = 6)
                        c0.value = reg
                        c1 = newsheet1.cell(row = i + 1 + cnt, column = 4)
                        c1.value = cat
                        c0 = newsheet1.cell(row = i + 2 + cnt, column = 2)
                        c0.value = valName
                        c0 = newsheet1.cell(row = i + 2 + cnt, column = 3)
                        c0.value = pos
                        c0 = newsheet1.cell(row = i + 2 + cnt, column = 6)
                        c0.value = ot
                        if cat.__contains__("Unarmed"): cat = "OT Unarmed"
                        if cat.__contains__("Armed"): cat = "OT Armed"
                        if cat.__contains__("Admin"): cat = "OT Admin"
                        c1 = newsheet1.cell(row = i + 2 + cnt, column = 4)
                        c1.value = cat
                        cnt += 2

            # if pos is not None:
            #     c0 = newsheet1.cell(row = i + 1, column = 2)
            #     c0.value = valName
            #     c0 = newsheet1.cell(row = i + 1, column = 3)
            #     c0.value = pos
            #     if reg != 0 and ot == 0:
            #         c0 = newsheet1.cell(row = i + 1, column = 5)
            #         c0.value = reg
            #         cnt += 1
            #     if reg == 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1, column = 5)
            #         c0.value = ot
            #         cnt += 1
            #     if reg != 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1, column = 5)
            #         c0.value = reg
            #         c0 = newsheet1.cell(row = i + 2, column = 2)
            #         c0.value = valName
            #         c0 = newsheet1.cell(row = i + 2, column = 3)
            #         c0.value = pos
            #         c0 = newsheet1.cell(row = i + 2, column = 5)
            #         c0.value = ot
            #         cnt += 2
            # pos = sheet.cell(row = i + 1 + 1, column = 1).value
            # reg = sheet.cell(row = i + 1 + 1, column = 2).value
            # ot  = sheet.cell(row = i + 1 + 1, column = 3).value
            # if pos is None: continue
            # if pos is not None:
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 2)
            #     c0.value = valName
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 3)
            #     c0.value = pos
            #     if reg != 0 and ot == 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         cnt += 1
            #     if reg == 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 1
            #     if reg != 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 2)
            #         c0.value = valName
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 3)
            #         c0.value = pos
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 2
            # pos = sheet.cell(row = i + 1 + 2, column = 1).value
            # reg = sheet.cell(row = i + 1 + 2, column = 2).value
            # ot  = sheet.cell(row = i + 1 + 2, column = 3).value
            # if pos is None: continue
            # if pos is not None:
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 2)
            #     c0.value = valName
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 3)
            #     c0.value = pos
            #     if reg != 0 and ot == 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         cnt += 1
            #     if reg == 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 1
            #     if reg != 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 2)
            #         c0.value = valName
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 3)
            #         c0.value = pos
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 2
            # pos = sheet.cell(row = i + 1 + 3, column = 1).value
            # reg = sheet.cell(row = i + 1 + 3, column = 2).value
            # ot  = sheet.cell(row = i + 1 + 3, column = 3).value
            # if pos is None: continue
            # if pos is not None:
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 2)
            #     c0.value = valName
            #     c0 = newsheet1.cell(row = i + 1 + cnt, column = 3)
            #     c0.value = pos
            #     if reg != 0 and ot == 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         cnt += 1
            #     if reg == 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 1
            #     if reg != 0 and ot != 0:
            #         c0 = newsheet1.cell(row = i + 1 + cnt, column = 5)
            #         c0.value = reg
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 2)
            #         c0.value = valName
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 3)
            #         c0.value = pos
            #         c0 = newsheet1.cell(row = i + 2 + cnt, column = 5)
            #         c0.value = ot
            #         cnt += 2





# Reorder by employee last name


c0 = newsheet1.cell(row = 1, column = 1)
c0.value = "Employee Number"
c0 = newsheet1.cell(row = 1, column = 2)
c0.value = "Employee Name"
c0 = newsheet1.cell(row = 1, column = 3)
c0.value = "Position"
c0 = newsheet1.cell(row = 1, column = 4)
c0.value = "Category"
c0 = newsheet1.cell(row = 1, column = 5)
c0.value = "Number of Shifts"
c0 = newsheet1.cell(row = 1, column = 6)
c0.value = "Number of Hours"
c0 = newsheet1.cell(row = 1, column = 7)
c0.value = "Pay Rate"
c0 = newsheet1.cell(row = 1, column = 8)
c0.value = "Total Pay"
c0 = newsheet1.cell(row = 1, column = 9)
c0.value = "Reimbursement/Deductions"
c0 = newsheet1.cell(row = 1, column = 10)
c0.value = "Grand Total"
c0 = newsheet1.cell(row = 1, column = 11)
c0.value = "Notes"

newsheet1.column_dimensions["A"].width = 20
newsheet1.column_dimensions["B"].width = 25
newsheet1.column_dimensions["C"].width = 48
newsheet1.column_dimensions["D"].width = 12
newsheet1.column_dimensions["E"].width = 12
newsheet1.column_dimensions["F"].width = 14
newsheet1.column_dimensions["G"].width = 12
newsheet1.column_dimensions["H"].width = 12
newsheet1.column_dimensions["I"].width = 12
newsheet1.column_dimensions["J"].width = 12
newsheet1.column_dimensions["K"].width = 12

# Delete empty rows
# indx = []
# for i in range(len(tuple(newsheet1.rows))):
#     flag = False
#     for cell in tuple(newsheet1.rows)[i]:
#         if cell.value != None:
#             flag = True
#             break
#     if flag == False:
#         indx.append(i)
# indx.sort()
# for i in range(len(indx)):
#     newsheet1.delete_rows(idx = indx[i] + 1 - i)


output_name = "output_merged.xlsx"
newbook1.save(output_name)
print("\n")
print("file output written to", output_name)
print("\nSpreadsheet informarion merged.\n")

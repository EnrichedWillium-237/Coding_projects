# Reading an excel file using Python
import openpyxl
from openpyxl import load_workbook

# Give location of file
wb = load_workbook('payroll.xlsx')

sheet = wb.active
c0_0 = sheet.cell(row=1,column=1)
c0_1 = sheet.cell(row=2,column=1)
c0_2 = sheet.cell(row=3,column=1)
c0_3 = sheet.cell(row=4,column=1)
print(c0_0.value, c0_1.value, c0_2.value, c0_3.value)

for i in range(1, 5):
    c1_0 = sheet.cell(row=i,column=1)
    c1_1 = sheet.cell(row=i,column=2)
    c1_2 = sheet.cell(row=i,column=3)
    c1_3 = sheet.cell(row=i,column=4)
    c1_4 = sheet.cell(row=i,column=5)
    c1_5 = sheet.cell(row=i,column=6)
    print(c1_0.value,"\t",c1_1.value,"\t",c1_2.value,"\t",c1_3.value,"\t",c1_4.value,"\t",c1_5.value)

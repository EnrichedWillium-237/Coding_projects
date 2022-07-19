### Code for reading in a payroll spreadsheet and organizing net payments ###

# Source files
import openpyxl
from openpyxl import load_workbook

# Payroll file location
workbook = load_workbook('payroll.xlsx')

sheet = workbook.active
label_0 = sheet.cell(row=1,column=1)
label_1 = sheet.cell(row=1,column=2)
label_2 = sheet.cell(row=1,column=3)
label_3 = sheet.cell(row=1,column=4)
label_4 = sheet.cell(row=1,column=5)
label_5 = sheet.cell(row=1,column=6)
label_6 = sheet.cell(row=1,column=7)
print(label_0.value, label_1.value, label_2.value, label_3.value, label_4.value, label_5.value, label_6.value)

Nrow = sheet.max_row
Nrow = Nrow - 2 # offset from spreadsheet
#print("Nrows =",Nrow)

for i in range (2, 12):
#for i in range(2, Nrow+1):
    valName = sheet.cell(row=i,column=1)
    valCat = sheet.cell(row=i,column=2)
    valHours = sheet.cell(row=i,column=3)
    valRate = sheet.cell(row=i,column=4)
    valTot = sheet.cell(row=i,column=5)
    valReim = sheet.cell(row=i,column=6)
    valNet = sheet.cell(row=i,column=7)

    Name = valName.value
    Cat = valCat.value
    Hours = valHours.value
    Rate = valRate.value
    Tot = valTot.value
    Reim = valReim.value
    if Reim is None:
        Reim = 0
    Net = Tot + Reim

    Name_nxt = sheet.cell(row=i+1,column=1)
    Name_nxt = Name_nxt.value
    Tot_nxt = Tot
    while Name == Name_nxt:
        Tot += Tot_nxt
    else: break
    print(i,Name,Name_nxt,Tot,Tot_nxt)
#    if Name_nxt == Name:
#        print(i,Name,Name_nxt)
    #print(Name,"\t",Cat,"\t",Hours,"\t",Rate,"\t",Tot,"\t",Reim,"\t",Net)

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
print(label_0.value, "\t", label_1.value, "\t", label_2.value, "\t", label_3.value, "\t", label_4.value, "\t", label_5.value, "\t", label_6.value)

Nrow = sheet.max_row
Nrow = Nrow - 1 # offset from spreadsheet
#print("Nrows =",Nrow)

hrsUnarmed = 0
hrsArmed = 0
hrsAdmin = 0
hrsOT = 0
hrsTrain = 0
hrsSick = 0
hrsCOVID = 0

for i in range (2, Nrow):
#for i in range(173, 188):
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
    Net = Tot + Reim # calculate reimbursements
    Tot_nxt = Tot

    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
        hrsUnarmed += Hours
    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
        hrsArmed += Hours
    elif Cat.__contains__("Admin"):
        hrsAdmin += Hours
    elif Cat.__contains__("OT"):
        hrsOT += Hours
    elif Cat.__contains__("Train"):
        hrsTrain += Hours
    elif Cat.__contains__("Sick"):
        hrsSick += Hours
    elif Cat.__contains__("Covid"):
        hrsCOVID += Hours
    else:
        print("Unknown category for hours!")

    Name_nxt = sheet.cell(row=i+1,column=1)
    Cat_nxt = sheet.cell(row=i+1,column=2)
    Hours_nxt = sheet.cell(row=i+1,column=3)
    Rate_nxt = sheet.cell(row=i+1,column=4)
    Tot_nxt = sheet.cell(row=i+1,column=5)
    Reim_nxt = sheet.cell(row=i+1,column=6)
    Net_nxt = sheet.cell(row=i+1,column=7)
    Name_nxt = Name_nxt.value
    Cat_nxt = Cat_nxt.value
    Hours_nxt = Hours_nxt.value
    Rate_nxt = Rate_nxt.value
    Tot_nxt = Tot_nxt.value
    Reim_nxt = Reim_nxt.value
    Net_nxt = Net_nxt.value

    # Reset values for new employee name
    if Name_nxt is None:
        print("Employee:",Name," Hours unarmed: ",hrsUnarmed,"armed:",hrsArmed,"admin:",hrsAdmin,"OT:",hrsOT,"training:",hrsTrain,"sick pay:",hrsSick,"COVID:",hrsCOVID)
        break
    if Name_nxt not in Name:
        print("Employee:",Name," Hours unarmed: ",hrsUnarmed,"armed:",hrsArmed,"admin:",hrsAdmin,"OT:",hrsOT,"training:",hrsTrain,"sick pay:",hrsSick,"COVID:",hrsCOVID)
        Name = Name_nxt # next employee
        hrsUnarmed = 0
        hrsArmed = 0
        hrsAdmin = 0
        hrsOT = 0
        hrsTrain = 0
        hrsSick = 0
        hrsCOVID = 0

print()
print("Done")
    #print(i,Name,Tot,Tot_nxt)
#    if Name_nxt == Name:
#        print(i,Name,Name_nxt)
    #print(Name,"\t",Cat,"\t",Hours,"\t",Rate,"\t",Tot,"\t",Reim,"\t",Net)

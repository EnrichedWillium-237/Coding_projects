### Code for reading in a payroll spreadsheet and organizing net payments ###

# Source files
import openpyxl
import array as arr
from openpyxl import load_workbook

# Payroll file location
workbook = load_workbook('payroll.xlsx')

# Payroll output
newbook = openpyxl.Workbook()
newsheet = newbook.active

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

# hours per category
hrsUnarmed = 0
hrsArmed = 0
hrsAdmin = 0
hrsOT = 0
hrsTrain = 0
hrsSick = 0
hrsCOVID = 0
hrsTotal = 0

GrandUnarmed = 0
GrandArmed = 0
GrandAdmin = 0
GrandOT = 0
GrandTrain = 0
GrandSick = 0
GrandCovid = 0

# hours per payrate
hrs_rate = arr.array("d",[0, 0, 0, 0, 0, 0, 0, 0, 0, 0])
rate_arr = arr.array("d",[0, 0, 0, 0, 0, 0, 0, 0, 0, 0])

Net = 0 # total pay per employee
GrandHrs = 0 # total number of hours
GrandTot = 0 # total payroll
GrandReim = 0 # total reimbursement
GrandNet = 0 # total payout: total + reimbursement

for i in range (2, Nrow):
#for i in range(2, 12):
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
#    Net = Tot + Reim # calculate reimbursements
    valNet = valNet.value
    Net += valNet
    Tot_nxt = Tot

    GrandHrs += Hours
    GrandTot += Tot
    GrandReim += Reim
    GrandNet += valNet

    # Sort by category
    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
        hrsUnarmed += Hours
        GrandUnarmed += Hours
    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
        hrsArmed += Hours
        GrandArmed += Hours
    elif Cat.__contains__("Admin"):
        hrsAdmin += Hours
        GrandAdmin += Hours
    elif Cat.__contains__("OT"):
        hrsOT += Hours
        GrandOT += Hours
    elif Cat.__contains__("Train"):
        hrsTrain += Hours
        GrandTrain += Hours
    elif Cat.__contains__("Sick"):
        hrsSick += Hours
        GrandSick += Hours
    elif Cat.__contains__("Covid"):
        hrsCOVID += Hours
        GrandCovid += Hours
    else:
        print("Unknown category for hours!")
    hrsTotal += Hours

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

    # Sort by payrate
    """
    k = 0
    for j in range(0, 10):
        nameNxt = sheet.cell(row=(i+j),column=1)
        nameNxt = nameNxt.value
        if nameNxt not in Name:
            break
        k += 1
    for j in range(0, k):
        nameNxt = sheet.cell(row=(i+j),column=1)
        nameNxt = nameNxt.value
        ratehrs = sheet.cell(row=(i+j),column=3)
        ratehrs = ratehrs.value
        rate1 = sheet.cell(row=(i+j),column=4)
        rate1 = rate1.value
 #       print(Name,nameNxt,i,j,k,rate1,ratehrs)
    """

    # Reset values for new employee name
    if Name_nxt is None:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        c1 = newsheet.cell(row = i, column = 1)
        c1.value = Name
        c2 = newsheet.cell(row = i, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet.cell(row = i, column = 3)
        c3.value = hrsArmed
        c4 = newsheet.cell(row = i, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet.cell(row = i, column = 5)
        c5.value = hrsOT
        c6 = newsheet.cell(row = i, column = 6)
        c6.value = hrsTrain
        c7 = newsheet.cell(row = i, column = 7)
        c7.value = hrsSick
        c8 = newsheet.cell(row = i, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet.cell(row = i, column = 9)
        c9.value = hrsTotal
        c10 = newsheet.cell(row = i, column = 10)
        c10.value = Net
        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        break
    if Name_nxt not in Name:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        c1 = newsheet.cell(row = i, column = 1)
        c1.value = Name
        c2 = newsheet.cell(row = i, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet.cell(row = i, column = 3)
        c3.value = hrsArmed
        c4 = newsheet.cell(row = i, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet.cell(row = i, column = 5)
        c5.value = hrsOT
        c6 = newsheet.cell(row = i, column = 6)
        c6.value = hrsTrain
        c7 = newsheet.cell(row = i, column = 7)
        c7.value = hrsSick
        c8 = newsheet.cell(row = i, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet.cell(row = i, column = 9)
        c9.value = hrsTotal
        c10 = newsheet.cell(row = i, column = 10)
        c10.value = Net
        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        Name = Name_nxt # next employee
        hrs_rate = arr.array("d",[0, 0, 0, 0, 0, 0, 0, 0, 0, 0]) # clear payrate array
        rate_arr = arr.array("d",[0, 0, 0, 0, 0, 0, 0, 0, 0, 0]) # clear payrate array
        Net = 0
        hrsUnarmed = 0
        hrsArmed = 0
        hrsAdmin = 0
        hrsOT = 0
        hrsTrain = 0
        hrsSick = 0
        hrsCOVID = 0
        hrsTotal = 0

GrandHrs = str(round(GrandHrs, 2))
GrandTot = str(round(GrandTot, 2))
GrandReim = str(round(GrandReim, 2))
GrandNet = str(round(GrandNet, 2))

cGrandHrs = newsheet.cell(row = Nrow+2, column = 1)
cGrandHrs.value = GrandHrs
cGrandTot = newsheet.cell(row = Nrow+2, column = 2)
cGrandTot.value = GrandTot
cGrandReim = newsheet.cell(row = Nrow+2, column = 3)
cGrandReim.value = GrandReim
cGrandNet = newsheet.cell(row = Nrow+2, column = 4)
cGrandNet.value = GrandNet

cGrandUnarmed = newsheet.cell(row = Nrow+3, column = 1)
cGrandUnarmed.value = GrandUnarmed
cGrandArmed = newsheet.cell(row = Nrow+3, column = 2)
cGrandArmed.value = GrandArmed
cGrandAdmin = newsheet.cell(row = Nrow+3, column = 3)
cGrandAdmin.value = GrandAdmin
cGrandOT = newsheet.cell(row = Nrow+3, column = 4)
cGrandOT.value = GrandOT
cGrandTrain = newsheet.cell(row = Nrow+3, column = 5)
cGrandTrain.value = GrandTrain
cGrandSick = newsheet.cell(row = Nrow+3, column = 6)
cGrandSick.value = GrandSick
cGrandCovid = newsheet.cell(row = Nrow+3, column = 7)
cGrandCovid.value = GrandCovid
newbook.save("/mnt/c/Users/Jacob/macros/test/output.xlsx")

print()
print("Total hours:",GrandHrs,"\t Total payroll:",GrandTot,"\t Reimbursements:",GrandReim,"\t Total net payroll:",GrandNet)
print("Total unarmed:",GrandUnarmed,"\t armed:",GrandArmed,"\t admin:",GrandAdmin,"\t OT:",GrandOT,"\t training:",GrandTrain,"\t sicktime:",GrandSick,"\t COVID:",GrandCovid)
print()
print("Done")
    #print(i,Name,Tot,Tot_nxt)
#    if Name_nxt == Name:
#        print(i,Name,Name_nxt)
    #print(Name,"\t",Cat,"\t",Hours,"\t",Rate,"\t",Tot,"\t",Reim,"\t",Net)

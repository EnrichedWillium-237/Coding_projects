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
Ncell = 1 # index for excel output

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

Net = 0 # total pay per employee
GrandHrs = 0 # total number of hours
GrandTot = 0 # total payroll
GrandReim = 0 # total reimbursement
GrandNet = 0 # total payout: total + reimbursement

c0 = newsheet.cell(row = 1, column = 1)
c0.value = "Employee name"
c0 = newsheet.cell(row = 1, column = 2)
c0.value = "Unarmed"
c0 = newsheet.cell(row = 1, column = 3)
c0.value = "Armed"
c0 = newsheet.cell(row = 1, column = 4)
c0.value = "Admin"
c0 = newsheet.cell(row = 1, column = 5)
c0.value = "OT"
c0 = newsheet.cell(row = 1, column = 6)
c0.value = "Training"
c0 = newsheet.cell(row = 1, column = 7)
c0.value = "Sick"
c0 = newsheet.cell(row = 1, column = 8)
c0.value = "COVID"
c0 = newsheet.cell(row = 1, column = 9)
c0.value = "Total hours"
c0 = newsheet.cell(row = 1, column = 10)
c0.value = "Net pay"

nrowsEmp = 0 # number of rows per employee
rate1 = 0
rate2 = 0
rate3 = 0
rate4 = 0
rate5 = 0
rate6 = 0
rate7 = 0
rate8 = 0
rate9 = 0
rate10 = 0
rate11 = 0
rate12 = 0
rate1hrs = 0
rate2hrs = 0
rate3hrs = 0
rate4hrs = 0
rate5hrs = 0
rate6hrs = 0
rate7hrs = 0
rate8hrs = 0
rate9hrs = 0
rate10hrs = 0
rate11hrs = 0
rate12hrs = 0

#for i in range (2, Nrow):
for i in range(2, 25):
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
    Tot_nxt = sheet.cell(row=i+1,column=5)
    Name_nxt = Name_nxt.value
    Tot_nxt = Tot_nxt.value

    nrowsEmp += 1

    # Sort by pay rate
    # The following is messy as heck. Need to fix at a later time.
    if Name_nxt is not Name:
        k = nrowsEmp
        #print(k,i,i-k+1)
        l = 0
        for j in range(i-k+1, i+1):
            valRate = sheet.cell(row=j,column=4)
            Rate = valRate.value
            valHours_tmp = sheet.cell(row=j,column=3)
            Hours_tmp = valHours_tmp.value
            #print(i-k+1,j,i+1,Rate,Hours_tmp)
            l += 1
            if l == 1:
                rate1 = Rate
                rate1hrs += Hours_tmp
            if l == 2:
                rate2 = Rate
                rate2hrs += Hours_tmp
            if l == 3:
                rate3 = Rate
                rate3hrs += Hours_tmp
            if l == 4:
                rate4 = Rate
                rate4hrs += Hours_tmp
            if l == 5:
                rate5 = Rate
                rate5hrs += Hours_tmp
            if l == 6:
                rate6 = Rate
                rate6hrs += Hours_tmp
            if l == 7:
                rate7 = Rate
                rate7hrs += Hours_tmp
            if l == 8:
                rate8 = Rate
                rate8hrs += Hours_tmp
            if l == 9:
                rate9 = Rate
                rate9hrs += Hours_tmp
            if l == 10:
                rate10 = Rate
                rate10hrs += Hours_tmp
            if l == 11:
                rate11 = Rate
                rate11hrs += Hours_tmp
            if l == 12:
                rate12 = Rate
                rate12hrs += Hours_tmp

    # End messiness


    # Reset values for new employee name
    if Name_nxt is None:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        Ncell += 1
        c1 = newsheet.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet.cell(row = Ncell, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet.cell(row = Ncell, column = 3)
        c3.value = hrsArmed
        c4 = newsheet.cell(row = Ncell, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet.cell(row = Ncell, column = 5)
        c5.value = hrsOT
        c6 = newsheet.cell(row = Ncell, column = 6)
        c6.value = hrsTrain
        c7 = newsheet.cell(row = Ncell, column = 7)
        c7.value = hrsSick
        c8 = newsheet.cell(row = Ncell, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet.cell(row = Ncell, column = 9)
        c9.value = hrsTotal
        c10 = newsheet.cell(row = Ncell, column = 10)
        c10.value = Net
        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        print("rate1:",rate1,"hrs1:",rate1hrs,"rate2:",rate2,"hrs2:",rate2hrs,"rate3:",rate3,"hrs3:",rate3hrs,"rate4:",rate4,"hrs4:",rate4hrs,"rate5:",rate5,"hrs5:",rate5hrs,"rate6:",rate6,"hrs6:",rate6hrs,"rate7:",rate7,"hrs7:",rate7hrs,"rate8:",rate8,"hrs8:",rate8hrs)
        break
    if Name_nxt not in Name:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        Ncell += 1
        c1 = newsheet.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet.cell(row = Ncell, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet.cell(row = Ncell, column = 3)
        c3.value = hrsArmed
        c4 = newsheet.cell(row = Ncell, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet.cell(row = Ncell, column = 5)
        c5.value = hrsOT
        c6 = newsheet.cell(row = Ncell, column = 6)
        c6.value = hrsTrain
        c7 = newsheet.cell(row = Ncell, column = 7)
        c7.value = hrsSick
        c8 = newsheet.cell(row = Ncell, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet.cell(row = Ncell, column = 9)
        c9.value = hrsTotal
        c10 = newsheet.cell(row = Ncell, column = 10)
        c10.value = Net
        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        print("rate1:",rate1,"hrs1:",rate1hrs,"rate2:",rate2,"hrs2:",rate2hrs,"rate3:",rate3,"hrs3:",rate3hrs,"rate4:",rate4,"hrs4:",rate4hrs,"rate5:",rate5,"hrs5:",rate5hrs,"rate6:",rate6,"hrs6:",rate6hrs,"rate7:",rate7,"hrs7:",rate7hrs,"rate8:",rate8,"hrs8:",rate8hrs)

        # Clear values for next employee
        Name = Name_nxt # next employee
        Net = 0
        hrsUnarmed = 0
        hrsArmed = 0
        hrsAdmin = 0
        hrsOT = 0
        hrsTrain = 0
        hrsSick = 0
        hrsCOVID = 0
        hrsTotal = 0
        nrowsEmp = 0
        rate1 = 0
        rate2 = 0
        rate3 = 0
        rate4 = 0
        rate5 = 0
        rate6 = 0
        rate7 = 0
        rate8 = 0
        rate9 = 0
        rate10 = 0
        rate11 = 0
        rate12 = 0
        rate1hrs = 0
        rate2hrs = 0
        rate3hrs = 0
        rate4hrs = 0
        rate5hrs = 0
        rate6hrs = 0
        rate7hrs = 0
        rate8hrs = 0
        rate9hrs = 0
        rate10hrs = 0
        rate11hrs = 0
        rate12hrs = 0

GrandHrs = str(round(GrandHrs, 2))
GrandTot = str(round(GrandTot, 2))
GrandReim = str(round(GrandReim, 2))
GrandNet = str(round(GrandNet, 2))

c0 = newsheet.cell(row = Ncell+2, column = 1)
c0.value = "Total hours ($)"
c0 = newsheet.cell(row = Ncell+2, column = 2)
c0.value = "Total pay ($)"
c0 = newsheet.cell(row = Ncell+2, column = 3)
c0.value = "Reimbursement ($)"
c0 = newsheet.cell(row = Ncell+2, column = 4)
c0.value = "Net pay ($)"

cGrandHrs = newsheet.cell(row = Ncell+3, column = 1)
cGrandHrs.value = GrandHrs
cGrandTot = newsheet.cell(row = Ncell+3, column = 2)
cGrandTot.value = GrandTot
cGrandReim = newsheet.cell(row = Ncell+3, column = 3)
cGrandReim.value = GrandReim
cGrandNet = newsheet.cell(row = Ncell+3, column = 4)
cGrandNet.value = GrandNet

c0 = newsheet.cell(row = Ncell+5, column = 1)
c0.value = "Total unarmed"
c0 = newsheet.cell(row = Ncell+5, column = 2)
c0.value = "Armed"
c0 = newsheet.cell(row = Ncell+5, column = 3)
c0.value = "Admin"
c0 = newsheet.cell(row = Ncell+5, column = 4)
c0.value = "OT"
c0 = newsheet.cell(row = Ncell+5, column = 5)
c0.value = "Training"
c0 = newsheet.cell(row = Ncell+5, column = 6)
c0.value = "Sick pay"
c0 = newsheet.cell(row = Ncell+5, column = 7)
c0.value = "COVID"

cGrandUnarmed = newsheet.cell(row = Ncell+6, column = 1)
cGrandUnarmed.value = GrandUnarmed
cGrandArmed = newsheet.cell(row = Ncell+6, column = 2)
cGrandArmed.value = GrandArmed
cGrandAdmin = newsheet.cell(row = Ncell+6, column = 3)
cGrandAdmin.value = GrandAdmin
cGrandOT = newsheet.cell(row = Ncell+6, column = 4)
cGrandOT.value = GrandOT
cGrandTrain = newsheet.cell(row = Ncell+6, column = 5)
cGrandTrain.value = GrandTrain
cGrandSick = newsheet.cell(row = Ncell+6, column = 6)
cGrandSick.value = GrandSick
cGrandCovid = newsheet.cell(row = Ncell+6, column = 7)
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


### Next iterration of this code will use arrays to make everything simpler, I promise.

### Code for reading in a payroll spreadsheet and organizing net payments ###

# Source files
import openpyxl
import array as arr
from openpyxl import load_workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# Payroll file location
workbook = load_workbook('payroll.xlsx')

# Payroll output
newbook0 = openpyxl.Workbook()
newbook1 = openpyxl.Workbook()
newsheet0 = newbook0.active
newsheet1 = newbook1.active

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

c0 = newsheet0.cell(row = 1, column = 1)
c0.value = "Employee name"
c0 = newsheet0.cell(row = 1, column = 2)
c0.value = "Unarmed"
c0 = newsheet0.cell(row = 1, column = 3)
c0.value = "Armed"
c0 = newsheet0.cell(row = 1, column = 4)
c0.value = "Admin"
c0 = newsheet0.cell(row = 1, column = 5)
c0.value = "OT"
c0 = newsheet0.cell(row = 1, column = 6)
c0.value = "Training"
c0 = newsheet0.cell(row = 1, column = 7)
c0.value = "Sick"
c0 = newsheet0.cell(row = 1, column = 8)
c0.value = "COVID"
c0 = newsheet0.cell(row = 1, column = 9)
c0.value = "Total hours"
c0 = newsheet0.cell(row = 1, column = 10)
c0.value = "Net pay"

c0 = newsheet1.cell(row = 1, column = 1)
c0.value = "Employee name"
c0 = newsheet1.cell(row = 1, column = 2)
c0.value = "rate 1"
c0 = newsheet1.cell(row = 1, column = 3)
c0.value = "rate 1 hrs"
c0 = newsheet1.cell(row = 1, column = 4)
c0.value = "rate 2"
c0 = newsheet1.cell(row = 1, column = 5)
c0.value = "rate 2 hrs"
c0 = newsheet1.cell(row = 1, column = 6)
c0.value = "rate 3"
c0 = newsheet1.cell(row = 1, column = 7)
c0.value = "rate 3 hrs"
c0 = newsheet1.cell(row = 1, column = 8)
c0.value = "rate 4"
c0 = newsheet1.cell(row = 1, column = 9)
c0.value = "rate 4 hrs"
c0 = newsheet1.cell(row = 1, column = 10)
c0.value = "rate 5"
c0 = newsheet1.cell(row = 1, column = 11)
c0.value = "rate 5 hrs"
c0 = newsheet1.cell(row = 1, column = 12)
c0.value = "rate 6"
c0 = newsheet1.cell(row = 1, column = 13)
c0.value = "rate 6 hrs"
c0 = newsheet1.cell(row = 1, column = 14)
c0.value = "rate 7"
c0 = newsheet1.cell(row = 1, column = 15)
c0.value = "rate 7 hrs"

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
rate1unarmed = 0
rate2unarmed = 0
rate3unarmed = 0
rate4unarmed = 0
rate5unarmed = 0
rate6unarmed = 0
rate7unarmed = 0
rate8unarmed = 0
rate9unarmed = 0
rate10unarmed = 0
rate11unarmed = 0
rate12unarmed = 0
rate1armed = 0
rate2armed = 0
rate3armed = 0
rate4armed = 0
rate5armed = 0
rate6armed = 0
rate7armed = 0
rate8armed = 0
rate9armed = 0
rate10armed = 0
rate11armed = 0
rate12armed = 0
rate1admin = 0
rate2admin = 0
rate3admin = 0
rate4admin = 0
rate5admin = 0
rate6admin = 0
rate7admin = 0
rate8admin = 0
rate9admin = 0
rate10admin = 0
rate11admin = 0
rate12admin = 0
rate1OT = 0
rate2OT = 0
rate3OT = 0
rate4OT = 0
rate5OT = 0
rate6OT = 0
rate7OT = 0
rate8OT = 0
rate9OT = 0
rate10OT = 0
rate11OT = 0
rate12OT = 0
rate1train = 0
rate2train = 0
rate3train = 0
rate4train = 0
rate5train = 0
rate6train = 0
rate7train = 0
rate8train = 0
rate9train = 0
rate10train = 0
rate11train = 0
rate12train = 0
rate1sick = 0
rate2sick = 0
rate3sick = 0
rate4sick = 0
rate5sick = 0
rate6sick = 0
rate7sick = 0
rate8sick = 0
rate9sick = 0
rate10sick = 0
rate11sick = 0
rate12sick = 0
rate1covid = 0
rate2covid = 0
rate3covid = 0
rate4covid = 0
rate5covid = 0
rate6covid = 0
rate7covid = 0
rate8covid = 0
rate9covid = 0
rate10covid = 0
rate11covid = 0
rate12covid = 0

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
    # Note to self: learn how to use arrays and goto commands in python
    if Name_nxt is not Name:
        k = nrowsEmp
        #print(k,i,i-k+1)
        l = 0
        for j in range(i-k+1, i+1):
            valRate = sheet.cell(row=j,column=4)
            Rate = valRate.value
            valHours_tmp = sheet.cell(row=j,column=3)
            Hours_tmp = valHours_tmp.value
            valCat_tmp = sheet.cell(row=j,column=2)
            Cat = valCat_tmp.value
            #print(i-k+1,j,i+1,Rate,Hours_tmp)
            l += 1
            if l == 1:
                rate1 = Rate
                rate1hrs += Hours_tmp
                if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                    rate1unarmed += Hours_tmp
                elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                    rate1armed += Hours_tmp
                elif Cat.__contains__("Admin"):
                    rate1admin += Hours_tmp
                elif Cat.__contains__("OT"):
                    rate1OT += Hours_tmp
                elif Cat.__contains__("Train"):
                    rate1train += Hours_tmp
                elif Cat.__contains__("Sick"):
                    rate1sick += Hours_tmp
                elif Cat.__contains__("Covid"):
                    rate1covid += Hours_tmp
                else:
                    print("Unknown category for hours!")
            if l == 2:
                rate2 = Rate
                if rate2 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 3:
                rate3 = Rate
                if rate3 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate3 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 4:
                rate4 = Rate
                if rate4 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate4 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate4 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 5:
                rate5 = Rate
                if rate5 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate5 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate5 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate5 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 6:
                rate6 = Rate
                if rate6 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate6 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate6 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate6 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate6 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 7:
                rate7 = Rate
                if rate7 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate7 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate7 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate7 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate7 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate7 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 8:
                rate8 = Rate
                if rate8 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                       rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate8 == rate7:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate8hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate8unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate8armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate8admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate8OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate8train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate8sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate8covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 9:
                rate9 = Rate
                if rate9 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate7:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate9 == rate8:
                    rate8hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate8unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate8armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate8admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate8OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate8train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate8sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate8covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate9hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate9unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate9armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate9admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate9OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate9train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate9sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate9covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 10:
                rate10 = Rate
                if rate10 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate7:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate8:
                    rate8hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate8unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate8armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate8admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate8OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate8train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate8sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate8covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate10 == rate9:
                    rate9hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate9unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate9armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate9admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate9OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate9train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate9sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate9covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate10hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate10unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate10armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate10admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate10OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate10train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate10sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate10covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 11:
                rate11 = Rate
                if rate11 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate7:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate8:
                    rate8hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate8unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate8armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate8admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate8OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate8train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate8sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate8covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate9:
                    rate9hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate9unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate9armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate9admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate9OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate9train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate9sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate9covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate11 == rate10:
                    rate10hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate10unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate10armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate10admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate10OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate10train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate10sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate10covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate11hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate11unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate11armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate11admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate11OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate11train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate11sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate11covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
            if l == 12:
                rate12 = Rate
                if rate12 == rate1:
                    rate1hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate1unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate1armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate1admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate1OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate1train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate1sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate1covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate2:
                    rate2hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate2unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate2armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate2admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate2OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate2train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate2sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate2covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate3:
                    rate3hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate3unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate3armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate3admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate3OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate3train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate3sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate3covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate4:
                    rate4hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate4unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate4armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate4admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate4OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate4train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate4sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate4covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate5:
                    rate5hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate5unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate5armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate5admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate5OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate5train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate5sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate5covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate6:
                    rate6hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate6unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate6armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate6admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate6OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate6train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate6sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate6covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate7:
                    rate7hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate7unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate7armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate7admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate7OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate7train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate7sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate7covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate8:
                    rate8hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate8unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate8armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate8admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate8OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate8train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate8sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate8covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate9:
                    rate9hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate9unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate9armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate9admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate9OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate9train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate9sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate9covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate10:
                    rate10hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate10unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate10armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate10admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate10OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate10train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate10sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate10covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                elif rate12 == rate11:
                    rate11hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate11unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate11armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate11admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate11OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate11train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate11sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate11covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")
                else:
                    rate12hrs += Hours_tmp
                    if Cat.__contains__("Un") and not Cat.__contains__("OT"):
                        rate12unarmed += Hours_tmp
                    elif Cat.__contains__("Armed") and not Cat.__contains__("OT"):
                        rate12armed += Hours_tmp
                    elif Cat.__contains__("Admin"):
                        rate12admin += Hours_tmp
                    elif Cat.__contains__("OT"):
                        rate12OT += Hours_tmp
                    elif Cat.__contains__("Train"):
                        rate12train += Hours_tmp
                    elif Cat.__contains__("Sick"):
                        rate12sick += Hours_tmp
                    elif Cat.__contains__("Covid"):
                        rate12covid += Hours_tmp
                    else:
                        print("Unknown category for hours!")

            if rate2hrs == 0:
                rate2 = 0
            if rate3hrs == 0:
                rate3 = 0
            if rate4hrs == 0:
                rate4 = 0
            if rate5hrs == 0:
                rate5 = 0
            if rate6hrs == 0:
                rate6 = 0
            if rate7hrs == 0:
                rate7 = 0
            if rate8hrs == 0:
                rate8 = 0
            if rate9hrs == 0:
                rate9 = 0
            if rate10hrs == 0:
                rate10 = 0
            if rate11hrs == 0:
                rate11 = 0
            if rate12hrs == 0:
                rate12 = 0
    # End messiness


    # Reset values for new employee name
    if Name_nxt is None:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        Ncell += 1
        c1 = newsheet0.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet0.cell(row = Ncell, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet0.cell(row = Ncell, column = 3)
        c3.value = hrsArmed
        c4 = newsheet0.cell(row = Ncell, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet0.cell(row = Ncell, column = 5)
        c5.value = hrsOT
        c6 = newsheet0.cell(row = Ncell, column = 6)
        c6.value = hrsTrain
        c7 = newsheet0.cell(row = Ncell, column = 7)
        c7.value = hrsSick
        c8 = newsheet0.cell(row = Ncell, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet0.cell(row = Ncell, column = 9)
        c9.value = hrsTotal
        c10 = newsheet0.cell(row = Ncell, column = 10)
        c10.value = Net

        c1 = newsheet1.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet1.cell(row = Ncell, column = 2)
        c2.value = rate1
        c3 = newsheet1.cell(row = Ncell, column = 3)
        c3.value = rate1hrs
        c4 = newsheet1.cell(row = Ncell, column = 4)
        c4.value = rate2
        if rate2 == 0:
            c4.value = None
        c5 = newsheet1.cell(row = Ncell, column = 5)
        c5.value = rate2hrs
        if rate2hrs == 0:
            c5.value = None
        c6 = newsheet1.cell(row = Ncell, column = 6)
        c6.value = rate3
        if rate3 == 0:
            c6.value = None
        c7 = newsheet1.cell(row = Ncell, column = 7)
        c7.value = rate3hrs
        if rate3hrs == 0:
            c7.value = None
        c8 = newsheet1.cell(row = Ncell, column = 8)
        c8.value = rate4
        if rate4 == 0:
            c8.value = None
        c9 = newsheet1.cell(row = Ncell, column = 9)
        c9.value = rate4hrs
        if rate4hrs == 0:
            c9.value = None
        c10 = newsheet1.cell(row = Ncell, column = 10)
        c10.value = rate5
        if rate5 == 0:
            c10.value = None
        c11 = newsheet1.cell(row = Ncell, column = 11)
        c11.value = rate5hrs
        if rate5hrs == 0:
            c11.value = None
        c12 = newsheet1.cell(row = Ncell, column = 12)
        c12.value = rate6
        if rate6 == 0:
            c12.value = None
        c13 = newsheet1.cell(row = Ncell, column = 13)
        c13.value = rate6hrs
        if rate6hrs == 0:
            c13.value = None

        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        print("\trate1:",rate1,"hrs1:",rate1hrs,"rate2:",rate2,"hrs2:",rate2hrs,"rate3:",rate3,"hrs3:",rate3hrs,"rate4:",rate4,"hrs4:",rate4hrs,"rate5:",rate5,"hrs5:",rate5hrs,"rate6:",rate6,"hrs6:",rate6hrs,"rate7:",rate7,"hrs7:",rate7hrs,"rate8:",rate8,"hrs8:",rate8hrs,"rate9:",rate9,"hrs9:",rate9hrs,"hrs10:",rate10,"hrs10:",rate10hrs,"rate11:",rate11hrs,"rate12:",rate12,"hrs12:",rate12hrs)
        break
    if Name_nxt not in Name:
        hrsTotal = float(round(hrsTotal,2))
        Net = str(round(Net, 2))
        #Ncell += 1
        Ncell += k
        c1 = newsheet0.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet0.cell(row = Ncell, column = 2)
        c2.value = hrsUnarmed
        c3 = newsheet0.cell(row = Ncell, column = 3)
        c3.value = hrsArmed
        c4 = newsheet0.cell(row = Ncell, column = 4)
        c4.value = hrsAdmin
        c5 = newsheet0.cell(row = Ncell, column = 5)
        c5.value = hrsOT
        c6 = newsheet0.cell(row = Ncell, column = 6)
        c6.value = hrsTrain
        c7 = newsheet0.cell(row = Ncell, column = 7)
        c7.value = hrsSick
        c8 = newsheet0.cell(row = Ncell, column = 8)
        c8.value = hrsCOVID
        c9 = newsheet0.cell(row = Ncell, column = 9)
        c9.value = hrsTotal
        c10 = newsheet0.cell(row = Ncell, column = 10)
        c10.value = Net
        '''
        c1 = newsheet1.cell(row = Ncell, column = 1)
        c1.value = Name
        c2 = newsheet1.cell(row = Ncell, column = 2)
        c2.value = rate1
        c3 = newsheet1.cell(row = Ncell, column = 3)
        c3.value = rate1hrs
        c4 = newsheet1.cell(row = Ncell, column = 4)
        c4.value = rate2
        if rate2 == 0:
            c4.value = None
        c5 = newsheet1.cell(row = Ncell, column = 5)
        c5.value = rate2hrs
        if rate2hrs == 0:
            c5.value = None
        c6 = newsheet1.cell(row = Ncell, column = 6)
        c6.value = rate3
        if rate3 == 0:
            c6.value = None
        c7 = newsheet1.cell(row = Ncell, column = 7)
        c7.value = rate3hrs
        if rate3hrs == 0:
            c7.value = None
        c8 = newsheet1.cell(row = Ncell, column = 8)
        c8.value = rate4
        if rate4 == 0:
            c8.value = None
        c9 = newsheet1.cell(row = Ncell, column = 9)
        c9.value = rate4hrs
        if rate4hrs == 0:
            c9.value = None
        c10 = newsheet1.cell(row = Ncell, column = 10)
        c10.value = rate5
        if rate5 == 0:
            c10.value = None
        c11 = newsheet1.cell(row = Ncell, column = 11)
        c11.value = rate5hrs
        if rate5hrs == 0:
            c11.value = None
        c12 = newsheet1.cell(row = Ncell, column = 12)
        c12.value = rate6
        if rate6 == 0:
            c12.value = None
        c13 = newsheet1.cell(row = Ncell, column = 13)
        c13.value = rate6hrs
        if rate6hrs == 0:
            c13.value = None
        '''
        c1 = newsheet1.cell(row = Ncell-k+1, column = 1)
        c1.value = Name
        for j in range(Ncell, Ncell+k):
            jrow = j-k
            print(j, Ncell, k,Ncell+k, jrow, rate1,rate1unarmed,rate1armed,rate1admin,rate1OT,"\t",rate2,rate2unarmed,rate2armed,rate2admin,rate2OT)
            c1 = newsheet1.cell(row = j-k+1, column = 2)
            c1.value = rate1
            c1 = newsheet1.cell(row = j-k+1, column = 3)
            c1.value = "Unarmed:"
            c1 = newsheet1.cell(row = j-k+1, column = 4)
            c1.value = rate1unarmed
            c1 = newsheet1.cell(row = j-k+1, column = 5)
            c1.value = "Armed:"
            c1 = newsheet1.cell(row = j-k+1, column = 6)
            c1.value = rate1armed
            c1 = newsheet1.cell(row = j-k+1, column = 7)
            c1.value = "Admin:"
            c1 = newsheet1.cell(row = j-k+1, column = 8)
            c1.value = rate1admin
            c1 = newsheet1.cell(row = j-k+1, column = 9)
            c1.value = "OT:"
            c1 = newsheet1.cell(row = j-k+1, column = 10)
            c1.value = rate1OT
            c1 = newsheet1.cell(row = j-k+1, column = 11)
            c1.value = "Train:"
            c1 = newsheet1.cell(row = j-k+1, column = 12)
            c1.value = rate1train
            c1 = newsheet1.cell(row = j-k+1, column = 13)
            c1.value = "Sick:"
            c1 = newsheet1.cell(row = j-k+1, column = 14)
            c1.value = rate1sick
            c1 = newsheet1.cell(row = j-k+1, column = 15)
            c1.value = "COVID:"
            c1 = newsheet1.cell(row = j-k+1, column = 16)
            c1.value = rate1covid

        print("Employee:",Name,"  Hours unarmed: ",hrsUnarmed," armed:",hrsArmed," admin:",hrsAdmin," OT:",hrsOT," training:",hrsTrain," sick pay:",hrsSick," COVID:",hrsCOVID," --- Total hours:",hrsTotal," Total pay:",Net)
        print("\trate1:",rate1,"hrs1:",rate1hrs,"rate2:",rate2,"hrs2:",rate2hrs,"rate3:",rate3,"hrs3:",rate3hrs,"rate4:",rate4,"hrs4:",rate4hrs,"rate5:",rate5,"hrs5:",rate5hrs,"rate6:",rate6,"hrs6:",rate6hrs,"rate7:",rate7,"hrs7:",rate7hrs,"rate8:",rate8,"hrs8:",rate8hrs,"rate9:",rate9,"hrs9:",rate9hrs,"hrs10:",rate10,"hrs10:",rate10hrs,"rate11:",rate11hrs,"rate12:",rate12,"hrs12:",rate12hrs)

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
        rate1unarmed = 0
        rate2unarmed = 0
        rate3unarmed = 0
        rate4unarmed = 0
        rate5unarmed = 0
        rate6unarmed = 0
        rate7unarmed = 0
        rate8unarmed = 0
        rate9unarmed = 0
        rate10unarmed = 0
        rate11unarmed = 0
        rate12unarmed = 0
        rate1armed = 0
        rate2armed = 0
        rate3armed = 0
        rate4armed = 0
        rate5armed = 0
        rate6armed = 0
        rate7armed = 0
        rate8armed = 0
        rate9armed = 0
        rate10armed = 0
        rate11armed = 0
        rate12armed = 0
        rate1admin = 0
        rate2admin = 0
        rate3admin = 0
        rate4admin = 0
        rate5admin = 0
        rate6admin = 0
        rate7admin = 0
        rate8admin = 0
        rate9admin = 0
        rate10admin = 0
        rate11admin = 0
        rate12admin = 0
        rate1OT = 0
        rate2OT = 0
        rate3OT = 0
        rate4OT = 0
        rate5OT = 0
        rate6OT = 0
        rate7OT = 0
        rate8OT = 0
        rate9OT = 0
        rate10OT = 0
        rate11OT = 0
        rate12OT = 0
        rate1train = 0
        rate2train = 0
        rate3train = 0
        rate4train = 0
        rate5train = 0
        rate6train = 0
        rate7train = 0
        rate8train = 0
        rate9train = 0
        rate10train = 0
        rate11train = 0
        rate12train = 0
        rate1sick = 0
        rate2sick = 0
        rate3sick = 0
        rate4sick = 0
        rate5sick = 0
        rate6sick = 0
        rate7sick = 0
        rate8sick = 0
        rate9sick = 0
        rate10sick = 0
        rate11sick = 0
        rate12sick = 0
        rate1covid = 0
        rate2covid = 0
        rate3covid = 0
        rate4covid = 0
        rate5covid = 0
        rate6covid = 0
        rate7covid = 0
        rate8covid = 0
        rate9covid = 0
        rate10covid = 0
        rate11covid = 0
        rate12covid = 0

# Excel spreadhseet options
newsheet0.column_dimensions["A"].width = 25
for row in newsheet0[2:newsheet0.max_row]:
    cell = row[9]
    cell.alignment = Alignment(horizontal='right')
'''
newsheet1.column_dimensions["A"].width = 25
newsheet1.column_dimensions["B"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'B{r}'].number_format = '##.## "$/hr"'
newsheet1.column_dimensions["D"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'D{r}'].number_format = '##.## "$/hr"'
newsheet1.column_dimensions["F"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'F{r}'].number_format = '##.## "$/hr"'
newsheet1.column_dimensions["H"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'H{r}'].number_format = '##.## "$/hr"'
newsheet1.column_dimensions["J"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'J{r}'].number_format = '##.## "$/hr"'
newsheet1.column_dimensions["L"].width = 10
for r in range(2, Ncell+1):
    newsheet1[f'L{r}'].number_format = '##.## "$/hr"'
'''

GrandHrs = str(round(GrandHrs, 2))
GrandTot = str(round(GrandTot, 2))
GrandReim = str(round(GrandReim, 2))
GrandNet = str(round(GrandNet, 2))

c0 = newsheet0.cell(row = Ncell+2, column = 2)
c0.value = "Total hours"
c0 = newsheet0.cell(row = Ncell+2, column = 3)
c0.value = "Total pay"
c0 = newsheet0.cell(row = Ncell+2, column = 4)
c0.value = "Reimbursement"
c0 = newsheet0.cell(row = Ncell+2, column = 5)
c0.value = "Net pay"

cGrandHrs = newsheet0.cell(row = Ncell+3, column = 2)
cGrandHrs.value = GrandHrs
cGrandTot = newsheet0.cell(row = Ncell+3, column = 3)
cGrandTot.value = GrandTot
cGrandReim = newsheet0.cell(row = Ncell+3, column = 4)
cGrandReim.value = GrandReim
cGrandNet = newsheet0.cell(row = Ncell+3, column = 5)
cGrandNet.value = GrandNet

c0 = newsheet0.cell(row = Ncell+5, column = 2)
c0.value = "Unarmed"
c0 = newsheet0.cell(row = Ncell+5, column = 3)
c0.value = "Armed"
c0 = newsheet0.cell(row = Ncell+5, column = 4)
c0.value = "Admin"
c0 = newsheet0.cell(row = Ncell+5, column = 5)
c0.value = "OT"
c0 = newsheet0.cell(row = Ncell+5, column = 6)
c0.value = "Training"
c0 = newsheet0.cell(row = Ncell+5, column = 7)
c0.value = "Sick pay"
c0 = newsheet0.cell(row = Ncell+5, column = 8)
c0.value = "COVID"

cGrandUnarmed = newsheet0.cell(row = Ncell+6, column = 2)
cGrandUnarmed.value = GrandUnarmed
cGrandArmed = newsheet0.cell(row = Ncell+6, column = 3)
cGrandArmed.value = GrandArmed
cGrandAdmin = newsheet0.cell(row = Ncell+6, column = 4)
cGrandAdmin.value = GrandAdmin
cGrandOT = newsheet0.cell(row = Ncell+6, column = 5)
cGrandOT.value = GrandOT
cGrandTrain = newsheet0.cell(row = Ncell+6, column = 6)
cGrandTrain.value = GrandTrain
cGrandSick = newsheet0.cell(row = Ncell+6, column = 7)
cGrandSick.value = GrandSick
cGrandCovid = newsheet0.cell(row = Ncell+6, column = 8)
cGrandCovid.value = GrandCovid

newbook0.save("/mnt/c/Users/Jacob/macros/test/output.xlsx")
newbook1.save("/mnt/c/Users/Jacob/macros/test/output_rate.xlsx")

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

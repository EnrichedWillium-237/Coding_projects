# include <fstream>
# include <iostream>
# include <iomanip>

using namespace std;

class employee {
public:
    double salary, hourlyRate, taxRate, taxMed, taxAmount, taxMedAmount, grossPay, netPay, otPay;
    int hours, otHours;

    int payStat;
    int employeeID;
    string firstName;
    string lastName;

public:
    void setVariables( int empID, string fName, string lName, int stat,  int rate, int hrs ) {
        employeeID = empID;
        firstName = fName;
        lastName = lName;
        payStat = stat;

        if (payStat == 1) {
            hourlyRate = rate;
        } else {
            salary = rate;
        }
        hours = hrs;
    }

public:
    virtual double calculateGrossPay() = 0;

    double calculateTaxAmount() {
        taxRate = 0.12;
        taxAmount = grossPay*taxRate;
        return taxAmount;
    }

    double calculateMedicadeAmount() {
        taxMed = 0.02;
        taxMedAmount = grossPay*taxMed;
        return taxMedAmount;
    }

    double calculateNetPay() {
        netPay = grossPay - taxAmount;
        return netPay;
    }

    void printData() {
        cout << setprecision(2) << setiosflags(ios::fixed | ios::showpoint);
        // cout << firstName << " " << setw(6) << lastName << setw(6) << employeeID << setw(10) << hours << setw(3)
        //      << otHours << setw(8) << grossPay << setw(8) << netPay << setw(8) << otPay << endl;
        //-------------
        cout << firstName << " " << lastName << "\t" << employeeID << "\t" << hours << "\t"
        << otHours << "\t" << grossPay << "\t" << netPay << "\t" << otPay << endl;
    }

};

class employeeHourly : public employee {
public:
    double calculateGrossPay()
    {
        const double regPay = (40*hourlyRate);
        if (hours > 40) {
            otHours = (hours - 40);
            otPay = (otHours * hourlyRate * 1.5);
            grossPay = (regPay + otPay);
        } else {
            otHours = 0;
            otPay = 0;
            grossPay = regPay;
        }
        return grossPay;
    }
};

class employeeSalary : public employee {
public:
    double calculateGrossPay() {
        double regPay = hours*hourlyRate;
        double hourlyRate = ((salary/52)/40);
        otHours = 0;
        otPay = 0;
        grossPay = regPay;
        // if (hours > 40) {
        //     otHours = (hours - 40);
        //     otPay = (otHours*hourlyRate);
        //     grossPay = (regPay + otPay);
        // } else if (hours <= 40) {
        //     otHours = 0;
        //     otPay = 0;
        //     grossPay = regPay;
        // }
        return grossPay;
    }
};

class employeeSalaryOT : public employee {
public:
    double calculateGrossPay() {
        double regPay = hours*hourlyRate;
        double hourlyRate = ((salary/52)/40);
        if (hours > 40) {
            otHours = (hours - 40);
            otPay = (otHours*hourlyRate);
            grossPay = (regPay + otPay);
        } else if (hours <= 40) {
            otHours = 0;
            otPay = 0;
            grossPay = regPay;
        }
        return grossPay;
    }
};

void payroll() {
    int employeeCounter = 0;
    int totalEmployeeCount = 0;
    string fName, lName;
    int empID = 0;
    int stat = 0;
    float rate = 0;
    float hrs = 0;

    cout << "Enter number of emplyees to process: ";
    cin >> totalEmployeeCount;

    employee * employee[100];

    while (employeeCounter<totalEmployeeCount) {
        cout << "Is employee " << employeeCounter+1 << " hourly or salary? (enter 1 for hourly, 2 for salary):";
        cin >> stat;

        if (stat == 1) {
            cout << "Initializing an HOURLY employee object inherited from base class employee" << endl << endl;

            cout << "Enter employee's ID: ";
            cin >> empID;
            cout << "Employee's first name: ";
            cin >> fName;
            cout << "Employee's last name: ";
            cin >> lName;
            cout << "Employee's hourly wage: ";
            cin >> rate;
            cout << "Employee's hours: ";
            cin >> hrs;

            employee[employeeCounter] = new employeeHourly();
            employee[employeeCounter]->setVariables( empID, fName, lName, stat, rate, hrs );
            employee[employeeCounter]->calculateGrossPay();
            employee[employeeCounter]->calculateTaxAmount();
            employee[employeeCounter]->calculateNetPay();
            employeeCounter++;
            cout << "\n" << endl;
        } else {
             cout << "Initializing a SALARY employee object in herited from base class employee" << endl << endl;

             cout << "Enter employee's ID: ";
             cin >> empID;
             cout << "Employee's first name: ";
             cin >> fName;
             cout << "Employee's last name: ";
             cin >> lName;
             cout << "Employeer's annual salary: ";
             cin >> rate;
             cout << "Employee's hours: ";
             cin >> hrs;

             employee[employeeCounter] = new employeeSalary();
             employee[employeeCounter]->setVariables( empID, fName, lName, stat, rate, hrs );
             employee[employeeCounter]->calculateGrossPay();
             employee[employeeCounter]->calculateTaxAmount();
             employee[employeeCounter]->calculateNetPay();
             employeeCounter++;
             cout << "\n" << endl;
         }
    }
    cout << "----------------------------------\n";
    cout << "Name             Emp ID       Hours   Hours OT   Gross pay   Net pay   OT pay" << endl;
    for (int i = 0; i<employeeCounter; i++) {
        employee[i]->printData();
    }
    cout << "----------------------------------\n";

    return;
}

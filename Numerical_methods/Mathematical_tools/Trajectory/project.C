// project.C
//
// A basic program for calculating the trajectory of a projectile under constant acceleration.
// Assumes no wind resistance or aerodynamic effects (will add those soon.)

//# include "NumMeth.h"
# include "TCanvas.h"
# include "TH1.h"
# include "TMath.h"

# include <cmath>
# include <fstream>
# include <iostream>

using namespace std;

const int num = 2;
Double_t y, speed, theta;
Double_t r1[num], v1[num], r[num], v[num], a[num];
const double pi = TMath::Pi();
const double pi2 = (double)2. * TMath::Pi();

// set physical parameters
double Cd = 0.35;       // Drag coefficient
double area =  4.0e-3;  // Cross-section of projectile (m/2)
double mass = 0.12;     // Projectile mass (kg)
double grav = 9.8;      // Gravitational acceleration (m/s^2)
double AD = 1.2;        // Air density (kg/m^3)
double airFlag, rho;

void project() {

    // Set initial conditions
    cout << "Enter initial height (m): "; cin >> y;
    r1[0] = 0;
    r1[1] = y; // initial vector position
    cout << "Enter initial speed (m/s): "; cin >> speed;
    cout << "Enter initial theta (degrees)"; cin >> theta;
    v1[0] = speed * TMath::Cos(theta*pi/180); // initial x-velocity
    v1[1] = speed * TMath::Sin(theta*pi/180); // initial y-velocity
    r[0] = r1[0];
    r[1] = r1[1];
    v[0] = v1[0];
    v[1] = v1[1];

    cout << "Account for air resistance? (Yes:1, No:0): "; cin >> airFlag;
    if (airFlag == 0) {
        rho = 0;
    } else {
        rho = AD;
    }
    double air_const = -0.5*Cd*rho*area/mass;
    double tau;
    cout << "Timestep (seconds): "; cin >> tau;
    int iStep, maxStep = 1e5;
    double *xplot, *yplot, *xNoAir, *yNoAir;
    xplot  = new double [maxStep];
    yplot  = new double [maxStep];
    xNoAir = new double [maxStep];
    yNoAir = new double [maxStep];

    // main loop
    for (iStep = 0; iStep<maxStep; iStep++) {
        xplot[iStep] = r[0];
        yplot[iStep] = r[1];

        double t = iStep*tau;
        xNoAir[iStep] = r1[0] + v1[0]*t;
        yNoAir[iStep] = r1[1] + v1[1]*t - 0.5*grav*t*t;

        double normV = sqrt( pow(v1[0],2) + pow(v1[1],2) ); // air resistance
        a[0] = air_const*normV*v[0];
        a[1] = air_const*normV*v[1];
        a[1] -= grav;

        // Euler method
        r[0] += tau*v[0];
        r[1] += tau*v[1];
        v[0] += tau*tau*a[0];
        v[1] += tau*tau*a[1];
        if (r[1] < 0) {
            xplot[iStep] = r[0];
            yplot[iStep] = r[1];
            break;
        }
    }
    cout << "Maximum range = " << r[0] << " meters" << endl;
    cout << "Time of flight = " << iStep*tau << " seconds" << endl;

    ofstream xplotOut("xplot.txt");
    ofstream yplotOut("yplot.txt");
    ofstream xNoAirOut("xNoAir.txt");
    ofstream yNoAirOut("yNoAir.txt");

    xplotOut << "Initial height (m): " << y << endl;
    xplotOut << "Initial speed (m/s): " << speed << endl;
    xplotOut << "Initial angle (degrees): " << theta << endl;
    xplotOut << "Timestep : " << tau << "\n" << endl;

    yplotOut << "Initial height (m): " << y << endl;
    yplotOut << "Initial speed (m/s): " << speed << endl;
    yplotOut << "Initial angle (degrees): " << theta << endl;
    yplotOut << "Timestep : " << tau << "\n" << endl;

    xNoAirOut << "Initial height (m): " << y << endl;
    xNoAirOut << "Initial speed (m/s): " << speed << endl;
    xNoAirOut << "Initial angle (degrees): " << theta << endl;
    xNoAirOut << "Timestep : " << tau << "\n" << endl;

    yNoAirOut << "Initial height (m): " << y << endl;
    yNoAirOut << "Initial speed (m/s): " << speed << endl;
    yNoAirOut << "Initial angle (degrees): " << theta << endl;
    yNoAirOut << "Timestep : " << tau << "\n" << endl;

    for (int i = 0; i<iStep; i++) {
        xplotOut << xplot[i] << endl;
        yplotOut << yplot[i] << endl;
        xNoAirOut << xNoAir[i] << endl;
        yNoAirOut << yNoAir[i] << endl;
    }

    delete[] xplot, yplot, xNoAir, yNoAir;

}

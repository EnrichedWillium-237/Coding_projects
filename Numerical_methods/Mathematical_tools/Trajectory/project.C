// project.C
//
// A basic program for calculating the trajectory of a projectile under constant acceleration.
// Assumes no wind resistance or aerodynamic effects (will add those soon.)

# include "TCanvas.h"
# include "TGraph.h"
# include "TH1.h"
# include "TLegend.h"
# include "TMath.h"

# include <cmath>
# include <fstream>
# include <iostream>

# include "style.h"

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
    int cnt = 0;
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
        cnt++;
    }
    cout << "Maximum range = " << r[0] << " meters" << endl;
    cout << "Time of flight = " << iStep*tau << " seconds" << endl;

    ofstream dataOut("data.txt");
    dataOut << "Initial height (m): " << y << endl;
    dataOut << "Initial speed (m/s): " << speed << endl;
    dataOut << "Initial angle (degrees): " << theta << endl;
    dataOut << "Timestep : " << tau << "\n" << endl;
    dataOut << "Maximum range = " << r[0] << " meters" << endl;
    dataOut << "Time of flight = " << iStep*tau << " seconds\n" << endl;
    dataOut << "xplot \t yplot \t xNoAir \t yNoAir" << endl;

    for (int i = 0; i<iStep; i++) {
        dataOut << xplot[i] << "\t " << yplot[i] << "\t " << xNoAir[i] << "\t " << yNoAir[i] << endl;
    }

    // plotting
    TCanvas * c0 = new TCanvas("c0", "c0", 600, 500);
    c0->cd();
    TGraph * g0 = new TGraph(cnt, xplot, yplot);
    g0->GetXaxis()->SetTitle("Range (m)");
    g0->GetYaxis()->SetTitle("Height (m)");
    g0->SetMarkerStyle(21);
    g0->SetMarkerSize(1.3);
    g0->SetMarkerColor(kBlue);
    g0->SetLineColor(kBlue);
    g0->Draw("APL");
    TGraph * g1 = new TGraph(cnt, xNoAir, yNoAir);
    g1->SetMarkerStyle(20);
    g1->SetMarkerSize(1.4);
    g1->SetMarkerColor(kRed);
    g1->SetLineColor(kRed);
    g1->Draw("same PL");
    TLegend * leg0 = new TLegend(0.44, 0.19, 0.69, 0.32);
    SetLegend(leg0, 21);
    leg0->AddEntry(g0, "No air", "lp");
    leg0->AddEntry(g1, "Air resistance", "lp");
    leg0->Draw();
    c0->Print("plot.pdf","pdf");

    delete[] xplot, yplot, xNoAir, yNoAir;

}

// Exercise for the calculation of a simple pendulum.
// Computes motion using either the Euler or Verlet methods.
// Adaptation from A. Garcia's "Numerical Methods for Physics".

# include "TCanvas.h"
# include "TGraph.h"
# include "TGraphErrors.h"
# include "TMath.h"

# include <fstream>
# include <iostream>

double theta0, theta, thetaLast, thetaNew;
double w = 0;
double gL = 1.0;
double t = 0.0;
double tlast;
int rev = 0;
int nT = 0;
int nStep;
double tau;
double a;

void pend() {

    cout << "Choose calculation method 1) Euler, 2) Verlet: ";
    int method;
    cin >> method;

    // Initial conditions
    cout << "Starting angle (degrees): "; cin >> theta0;
    theta = theta0*TMath::Pi()/180;
    w = 0.0;
    cout << "Enter time step: "; cin >> tau;

    // Verlet method
    a = -gL*TMath::Sin(theta);
    thetaLast = theta - w*tau + 0.5*pow(tau, 2)*a;
    cout << "Enter number of time steps: "; cin >> nStep;
    double *t_plot, * th_plot, *T;
    t_plot = new double[nStep+1];
    th_plot = new double[nStep+1];
    T = new double[nStep+1];

    // Main loop
    for (int iStep = 0; iStep<nStep; iStep++) {
        t_plot[iStep] = t;
        th_plot[iStep] = theta*180/TMath::Pi();
        t += tau;
        a = -gL*TMath::Sin(theta);
        if (method == 1) {
            // Euler method
            thetaLast = theta;
            theta += tau*w;
            w += tau*a;
        } else {
            // Varlet method
            thetaNew = 2*theta - thetaLast + pow(tau,2)*a;
            thetaLast = theta;
            theta = thetaNew;
        }
        if (theta*thetaLast<0) {
            // cout << "Passes zero at time: " << t << endl;
            if (rev == 0) {
                tlast = t;
            } else {
                T[rev] = 2*(t - tlast);
                tlast = t;
            }
            rev++;
        }
    }
    nT = rev - 1; // count of oscillations

    double TAve = 0., Terr = 0.;
    for (int i = 0; i<nT; i++) TAve += T[i];
    TAve /= nT;
    for (int i = 0; i<nT; i++) Terr += (T[i] - TAve)*(T[i] - TAve);
    Terr = sqrt(Terr/(nT*(nT-1)));
    cout << "Average period: " << TAve << " +/- " << Terr << endl;
    ofstream plotOut("data.txt");
    plotOut << "Method 1) Euler, 2) Verlet: " << method << endl;
    plotOut << "Starting angle (degrees): " << theta0 << endl;
    plotOut << "Time step: " << tau << endl;
    plotOut << "Number of time steps: " << nStep << endl;
    plotOut << "Average period: " << TAve << " +/- " << Terr << endl;
    plotOut << "time\t theta " << endl;
    for (int i = 0; i<nStep; i++) {
        plotOut << t_plot[i] << "\t " << th_plot[i] << endl;
    }

    TCanvas * c0 = new TCanvas("c0", "c0", 600, 500);
    c0->cd();
    TGraph * g0 = new TGraph(rev, t_plot, th_plot);
    g0->Draw("APL");

    delete[] t_plot, th_plot, T;

}

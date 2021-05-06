// corbit.C  Program for computing the trajectory of a comet

# include "TCanvas.h"
# include "TMath.h"

# include <fstream>
# include <iostream>

void gravrk( double x[], double t, double param[], double deriv[] );
void rk4( double x[], int nX, double t, double tau,
    void (*derivsRK)(double x[], double t, double param[], double deriv[]),
    double param[]);
void rka( double x[], int nX, double& t, double& tau, double err,
    void (*derivsRK)(double x[], double t, double param[], double deriv[]),
    double param[]);

double r0, v0;
double r[2], v[2], state[4], a[2];
int nState = 4;
// Set physical constants
double GM = 4*pow(TMath::Pi(),2);
double param[2];
double mass = 1.0;          // comet mass
double adaptErr = 1.0e-3;   // Runge-Kutta error parameter

void corbit() {

    cout << "Enter initial distance (AU): "; cin >> r0;
    cout << "Enter initial tangental velocity (AU/year): "; cin >> v0;
    r[0] = r0;
    r[1] = 0;
    v[0] = 0;
    v[1] = 0;
    state[0] = r[0];
    state[1] = r[1];
    state[2] = v[0];
    state[3] = v[1];
    param[0] = GM;  // normalized central force constant
    double time = 0;

    // main event loop

}

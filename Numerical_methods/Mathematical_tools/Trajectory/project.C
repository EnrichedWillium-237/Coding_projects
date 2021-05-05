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

const int num = 3;
Double_t y, speed, theta;
Double_t r1[num], v1[num], r[num], v[num], a[num];
const double pi = TMath::Pi();
const double pi2 = (double)2. * TMath::Pi();

void project() {

    // Set initial conditions
    cout << "Enter initial height (m): "; cin >> y;
    r1[1] = 0;
    r1[2] = y; // initial vector position
    cout << "Enter initial speed (m/s): "; cin >> speed;
    cout << "Enter initial theta (degrees)"; cin >> theta;
    v1[1] = speed * TMath::Cos(theta*pi/180); // initial x-velocity
    v1[2] = speed * TMath::Sin(theta*pi/180); // initial y-velocity
    r[1] = r1[1];
    r[2] = r1[2];
    v[1] = v1[1];
    v[2] = v1[2];

}

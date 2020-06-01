// interp.C
// Uses Lagrange polynomials to fit quadratic function to three data points

#include <assert.h>
#include <iostream>
#include <fstream>
#include "math.h"
#include "Matrix.h"

using namespace std;

void interp() {

    // Setup quadratic fit
    Double_t x[3+1], y[3+1];
    Double_t xmin, xmax;
    cout << "Enter data points: " << endl;
    Int_t i;
    for (i = 1; i<=3; i++) {
        x[i] = i; // get rid of after debugging
        y[i] = i; // get rid of after debugging
        // cout << "x[" << i << "] = ";
        // cin >> x[i];
        // cout << "y[" << i << "] = ";
        // cin >> y[i];
    }

    xmin = -1; xmax = 10; // get rid of after debugging
    // cout << "Minimum x: "; cin >> xmin;
    // cout << "Maximum x: "; cin >> xmax;
    cout << "code has run succesfully" << endl;
}

Double_t intrpf( Double_t xi, Double_t x[], double y[] );

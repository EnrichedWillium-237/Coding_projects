// interp.C
// Uses Lagrange polynomials to fit quadratic function to three data points

# include "math.h"
# include "Matrix.h"
# include "TCanvas.h"
# include "TGraphErrors.h"

# include <assert.h>
# include <iostream>
# include <fstream>

using namespace std;

Double_t xmin, xmax;
Double_t *xi, *yi;
Int_t nplot = 100;
TGraphErrors * g0;

Double_t intrpf( Double_t xi, Double_t x[], Double_t y[] );

void interp() {

    Double_t x[3+1], y[3+1];

    cout << "Enter data points: " << endl;
    Int_t i;
    for (i=0; i<3; i++) {
        cout << "x[" << i << "] = ";
        cin >> x[i];
        cout << "y[" << i << "] = ";
        cin >> y[i];
    }

    cout << "Minimum x: "; cin >> xmin;
    cout << "Maximum x: "; cin >> xmax;

    xi = new Double_t[nplot+1];
    yi = new Double_t[nplot+1];
    for (i=0; i<nplot; i++) {
        xi[i] = xmin + (xmax - xmin)*(i - 1)/(nplot - 1);
        yi[i] = intrpf(xi[i], x, y);
    }

    // Plot results
    TCanvas * c0 = new TCanvas("c0", "c0", 600, 500);
    c0->cd();
    g0 = new TGraphErrors(nplot, xi, yi, 0, 0);
    g0->SetMarkerColor(kBlue);
    g0->SetLineColor(kBlue);
    g0->Draw("ap");


    // Print out plotting variables
    ofstream xOut("x.txt"), yOut("y.txt"), xiOut("xi.txt"), yiOut("yi.txt");
    for (i=0; i<=2; i++) {
        xOut << x[i] << endl;
        yOut << y[i] << endl;
    }
    for (i=0; i<nplot; i++) {
        xiOut << xi[i] << endl;
        yiOut << yi[i] << endl;
    }
    delete [] xi, yi;

}

Double_t intrpf( Double_t xi, Double_t x[], Double_t y[] ) {

    Double_t yi = (xi-x[2])*(xi-x[3])/((x[1]-x[2])*(x[1]-x[3]))*y[1]
    + (xi-x[1])*(xi-x[3])/((x[2]-x[1])*(x[2]-x[3]))*y[2]
    + (xi-x[1])*(xi-x[2])/((x[3]-x[1])*(x[3]-x[2]))*y[3];

    return (yi);

}

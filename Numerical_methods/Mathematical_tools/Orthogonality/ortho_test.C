// ortho_test.C
//
// A simple program to test if two vectors are orthogonal to each other.
// Assumes vectors are in a Euclidean three-space.

# include "TAxis.h"
# include "TCanvas.h"
# include "TH1.h"
# include "TLine.h"

# include <iostream>

using namespace std;

void ortho_test(  ) {

    double a[3 + 1];
    double b[3 + 1];
    int i, j;

    cout << "\nInput vector coordinates: " << endl;

    for (i = 1; i<=3; i++) {
        cout << " a[" << i << "] = ";
        cin >> a[i];
    }
    for (j = 1; j<=3; j++) {
        cout << " b[" << j << "] = ";
        cin >> b[j];
    }

    // Calculate the dot product of the two vectors
    double a_dot_b = 0.;
    for (i=1; i<=3; i++) a_dot_b += a[i] * b[i];

    if (a_dot_b == 0.) {
        cout << "Vectors are orthogonal" << endl;
    }
    else {
        cout << "Vectors are not orthogonal" << endl;
        cout << "Dot product = " << a_dot_b << endl;
    }

    // Plotting vectors (just in the x-y plane)
    TH1D * h0;
    double x1 = 0;
    double x2 = 0;
    double y1 = 0;
    double y2 = 0;
    x1 = a[1] - 1.1*a[1];
    x2 = b[1] + 1.1*b[1];

    y1 = a[2] - 1.1*a[2];
    y2 = b[2] + 1.1*b[2];

    cout << "x1 = " << x1 << "\tx2 = " << x2 << endl;
    cout << "y1 = " << y1 << "\ty2 = " << y2 << endl;
    TCanvas * c0 = new TCanvas("c0", "c0", 600, 500);
    c0->cd();
    h0 = new TH1D("h0", "", 100, x1, x2);
    h0->SetMinimum(y1);
    h0->SetMaximum(y2);
    h0->Draw("");
    //TLine * l1 = new TLine(a[1], a[2], b[1], b[2]);
    TLine * l1 = new TLine(a[1], a[2], b[1], b[2]);
    l1->SetLineColor(kBlack);
    l1->SetLineWidth(2);
    l1->Draw("l");

}

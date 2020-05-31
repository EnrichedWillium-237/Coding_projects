// ortho_test.C
//
// A simple program to test if two vectors are orthogonal to each other.
// Assumes vectors are in a Euclidean three-space.

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

}

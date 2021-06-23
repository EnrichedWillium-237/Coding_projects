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

void corbit() {

    double r0, v0;
    double r[3], v[3], state[5], a[3];
    int nState = 4;
    // Set physical constants
    double GM = 4*pow(TMath::Pi(),2);
    double param[2];
    double mass = 1.0;          // comet mass
    double adaptErr = 1.0e-3;   // Runge-Kutta error parameter

    // cout << "Enter initial distance (AU): "; cin >> r0;
    // cout << "Enter initial tangental velocity (AU/year): "; cin >> v0;
    ////////
    r0 = 1.;
    v0 = 2.*TMath::Pi();
    int nStep = 1000.;
    double tau = 0.1;
    int method = 1;
    ///////
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

    // cout << "Enter number of steps: "; int nStep; cin >> nStep;
    // cout << "Enter time step (years): "; double tau; cin >> tau;
    // cout << "1) Euler, 2) Euler-Cromer, 3) Runge-Kutta, 4) Adaptive Runge-Kutta: ";
    // int method; cin >> method;

    // plotting variables
    double *rplot, *thplot, *tplot, *kinetic, *potential;
    rplot     = new double [nStep+1];
    thplot    = new double [nStep+1];
    tplot     = new double [nStep+1];
    kinetic   = new double [nStep+1];
    potential = new double [nStep+1];

    // main event loop
    for (int iStep = 0; iStep<nStep; iStep++) {
        double normR = sqrt( r[0]*r[0] + r[1]*r[1] );
        double normV = sqrt( v[0]*v[0] + v[1]*v[1] );
        rplot[iStep] = normR;               // Record position for plotting
        thplot[iStep] = TMath::ATan2(r[1], r[0]);
        tplot[iStep] = time;
        kinetic[iStep] = 0.5*mass*pow(normV,2);   // Record energies
        potential[iStep] = - GM*mass/normR;

        if (method == 1) { // Euler
            a[0] = -GM*r[0]/pow(normR,3);
            a[1] = -GM*r[1]/pow(normR,3);
            r[0] += tau*v[0];             // Euler step
            r[1] += tau*v[1];
            cout<<r[1]<<"\t"<<tau<<"\t"<<v[1]<<endl;
            v[0] += tau*a[0];
            v[1] += tau*a[1];
            time += tau;
            //cout <<r[0]<<"\t"<<r[1]<<"\t"<<v[0]<<"\t"<< v[1] << "\t" << a[0]<<"\t"<<a[1] << endl;
        } else if (method == 2) { // Euler-Cromer
            a[0] = -1.*GM*r[0]/pow(normR, 3);
            a[1] = -1.*GM*r[1]/pow(normR, 3);
            v[0] += tau*a[0];
            v[1] += tau*a[1];
            r[0] += tau*v[0];
            r[1] =+ tau*v[1];
            time += tau;
        } else if (method == 3) { // rk4
            rk4( state, nState, time, tau, gravrk, param );
            r[0] = state[0];
            r[1] = state[1];
            v[0] = state[2];
            v[1] = state[3];
            time += tau;
        } else {
            rka( state, nState, time, tau, adaptErr, gravrk, param );
            r[0] = state[0];
            r[1] = state[1];
            v[0] = state[2];
            v[1] = state[3];
        }
    }

    ofstream plotOut("data.txt");
    plotOut << "theta: \tradius: \tperiod: \tPE: \tKE: " << endl;
    for (int i = 0; i<nStep; i++) {
        plotOut << thplot[i] << "\t" << rplot[i] << "\t" << tplot[i] << "\t" << potential[i] << "\t" << kinetic[i] << endl;
    }

    // plotting


    delete [] rplot, thplot, tplot, potential, kinetic;

}

void gravrk( double x[], double t, double param[], double deriv[] ) {
    // Calculate acceleration
    double GM = param[0];
    double r1 = x[0];
    double r2 = x[1];
    double v1 = x[2];
    double v2 = x[3];
    double normR = sqrt( pow(r1, 2) + pow(r2, 2) );
    double a1 = -1.*GM*r1/(pow(normR, 3));
    double a2 = -1.*GM*r2/(pow(normR, 3));

    deriv[0] = v1;
    deriv[1] = v2;
    deriv[2] = a1;
    deriv[3] = a2;
}

void rk4( double x[], int nX, double t, double tau,
    void (*derivsRK) (double x[], double t, double param[], double deriv[]), double param[]) {

    double *F1, *F2, *F3, *F4, *xtemp;
    F1 = new double [nX];
    F2 = new double [nX];
    F3 = new double [nX];
    F4 = new double [nX];
    xtemp = new double [nX];

    // Calculate F1 = f(x,t)
    (derivsRK)( x, t, param, F1 );

    // Calculate F2 = f( x+tau*F1/2, t+tau/2 )
    double t_half = t + 0.5*tau;
    for (int i = 0; i<nX; i++) {
        xtemp[i] = x[i] + 0.5*tau*F1[i];
    }
    (*derivsRK) ( xtemp, t_half, param, F2 );

    // Calculate F3 = f( x+tau*F2/2, t+tau/2 )
    for (int i = 0; i<nX; i++) {
        xtemp[i] = x[i] + 0.5*tau*F2[i];
    }
    (*derivsRK) ( xtemp, t_half, param, F3 );

    // Calculate F4 = f( x+tau*F3, t+tau )
    double t_full = t + tau;
    for (int i = 1; i<nX; i++) {
        xtemp[i] = x[i] + tau*F3[i];
    }
    (*derivsRK) ( xtemp, t_full, param, F4 );

    for (int i = 1; i<nX; i++) {
        x[i] += tau/6.*(F1[i] + F4[i] + 2.*(F2[i] + F3[i]));
    }

    delete [] F1, F2, F3, F4, xtemp;
}

void rka( double x[], int nX, double t, double tau, double err,
    void (*derivsRK) (double x[], double t, double param[], double deriv[]), double param[] ) {

    double tSave = t;
    double safe1 = 0.9;
    double safe2 = 4.0;
    double *xSmall, *xBig;
    xSmall = new double [nX];
    xBig = new double [nX];
    int imax = 100;
    for (int i = 0; i<imax; i++) {
        // take two small time steps
        double halft = 0.5*tau;
        for (int j = 0; j<nX; j++) {
            xSmall[j] = x[j];
        }
        rk4( xSmall, nX, tSave, halft, derivsRK, param );
        t = tSave + halft;
        rk4( xSmall, nX, t, halft, derivsRK, param );

        // take big time step
        t = tSave + tau;
        for (int j = 0; j<nX; j++) {
            xBig[j] = x[j];
        }
        rk4( xBig, nX, tSave, tau, derivsRK, param );

        // compute error
        double errRat = 0.;
        double eps = 1.e-16;
        for (int j = 0; j<nX; j++) {
            double scale = err*(fabs(xSmall[j]) + fabs(xBig[j]))/2.;
            double xDiff = xSmall[j] - xBig[j];
            double ratio = fabs(xDiff)/(scale + eps);
            errRat = (errRat>ratio) ? errRat:ratio;
        }

        double oldt = tau;
        tau = safe1*oldt*pow(errRat, -0.2);
        tau = (tau > oldt/safe2) ? tau:oldt/safe2;
        tau = (tau <= safe2/oldt) ? tau:safe2*oldt;

        if (errRat < 1) {
            for (int j = 0; j<nX; j++) {
                x[j] = xSmall[j];
            }
            return;
        }
    }

    cout << "Error: Adaptive Runge-Kutta method failed... " << endl;

}

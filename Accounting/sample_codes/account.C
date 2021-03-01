// account.C
// A simple code for sorting accounts
// Reads in data.txt

# include "TMath.h"
# include "TRandom3.h"
# include <cmath>
# include <cstdlib>
# include <fstream>
# include <iostream>

using namespace std;

int main() {
    int acctnum[1000];
    char name[1000];
    double balance[1000];

    char input;
    int search_acct;
    int search_name;

    cout << "Please enter the name of the file: \n";

    char filename[50];
    ifstream random;
    cin.getline(filename, 50);
    random.open(filename);

    
}

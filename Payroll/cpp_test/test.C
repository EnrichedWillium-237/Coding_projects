#include <fstream>
#include <sstream>
#include <iostream>
#include <iomanip>
#include <vector>

using namespace std;

int test() {
    string filename{"input.csv"};
    ifstream input{filename};

    if (!input.is_open()) {
      cerr << "Couldn't read file: " << filename << "\n";
      return 1;
    }
    cout<<"test"<<endl;

    vector<vector<string>> csvRows;

    for (string line; std::getline(input, line);) {
      istringstream ss(move(line));
      vector<string> row;
      if (!csvRows.empty()) {
         // We expect each row to be as big as the first row
        row.reserve(csvRows.front().size());
      }
      // std::getline can split on other characters, here we use ','
      for (string value; getline(ss, value, ',');) {
        row.push_back(move(value));
      }
      csvRows.push_back(move(row));
    }

    // Print out our table
    for (const vector<string>& row : csvRows) {
      for (const string& value : row) {
        cout << setw(10) << value;
      }
      cout << "\n";
    }
}

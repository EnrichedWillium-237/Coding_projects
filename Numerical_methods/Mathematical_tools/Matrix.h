#include <assert.h>  // Defines the assert function.

class Matrix {

public:
// Default Constructor. Creates a 1 by 1 matrix and sets value to zero.
Matrix() {
    nRow_ = 1;
    nCol_ = 1;
    data_ = new double [1]; // Allocate memory
    set(0.0); // Set value of data_[0] to 0.0
}

// Create an nR by nC matrix and sets values to zero.
Matrix( int nR, int nC = 1 ) {
    assert(nR > 0 && nC > 0); // Check that nC and nR both > 0.
    nRow_ = nR; nCol_ = nC;
    data_ = new double [nR*nC]; // Allocate memory
    assert(data_ != 0); // Check that memory was allocated
    set(0.0); // Set values of data_[] to 0.0
}

// Copy Constructor.
// Used when a copy of an object is produced
// (e.g., passing to a function by value)
Matrix( const Matrix& mat ) {
    this->copy(mat); // Call private copy function.
}

// Destructor. Called when a Matrix object goes out of scope.
~Matrix() {
    delete [] data_; // Release allocated memory
}

// Assignment operator function.
// Overloads the equal sign operator to work with
// Matrix objects.
Matrix& operator = (const Matrix& mat) {
    if (this == &mat) return *this; // If two sides equal, do nothing.
    delete [] data_;                // Delete data on left hand side
    this->copy(mat);                // Copy right hand side to l.h.s.
    return *this;
}

// Simple "get" functions. Return number of rows or columns.
int nRow() const { return nRow_; }
int nCol() const { return nCol_; }

// Parenthesis operator function.
double& operator() (int i, int j = 1) {
    assert(i > 0 && i <= nRow_);
    assert(j > 0 && j <= nCol_);
    return data_[ nCol_*(i-1) + (j-1) ];
}

// Parenthesis operator function (const version).
const double& operator() (int i, int j = 1) const{
    assert(i > 0 && i <= nRow_);
    assert(j > 0 && j <= nCol_);
    return data_[ nCol_*(i-1) + (j-1) ];
}

// Set function. Sets all elements of a matrix to a given value.
void set(double value) {
    int i, iData = nRow_*nCol_;
    for (i=0; i<iData; i++) data_[i] = value;
}

//*********************************************************************
private:

// Matrix data.
int nRow_, nCol_;  // Number of rows, columns
double* data_;

// Private copy function.
// Copies values from one Matrix object to another.
void copy(const Matrix& mat) {
    nRow_ = mat.nRow_;
    nCol_ = mat.nCol_;
    int i, iData = nRow_*nCol_;
    data_ = new double [iData];
    for(i = 0; i<iData; i++ ) data_[i] = mat.data_[i];
}

};

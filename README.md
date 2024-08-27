# Gaussian Elimination Numerical Method Implementation in Python

This repository contains a Python implementation of the Gaussian Elimination method for solving systems of linear equations. The code reads coefficients from an Excel file (`read.xls`), performs Gaussian Elimination to transform the system into an upper triangular matrix, and then back-substitutes to find the solution. The results are saved back into the Excel file.

### Table of Contents
- [Gaussian Elimination Theory](#gaussian-elimination-theory)
- [Dependencies](#dependencies)
- [Installation](#installation)
- [Usage](#usage)
- [Code Explanation](#code-explanation)
- [Example](#example)
- [Files in the Repository](#files-in-the-repository)
- [Input Parameters](#input-parameters)
- [Troubleshooting](#troubleshooting)
- [Author](#author)

### Gaussian Elimination Theory
Gaussian Elimination is a method for solving linear systems of equations. It involves two main steps:

1. **Forward Elimination**: Transform the system into an upper triangular matrix.
2. **Back Substitution**: Solve the equations starting from the last row upwards.

By systematically performing row operations, the matrix is reduced to a simpler form from which the solutions can be easily obtained.

### Dependencies
To run this code, you need the following libraries:
- `numpy`
- `xlrd`
- `xlutils`

### Installation
To install the required libraries, you can use `pip`:
```sh
pip install numpy xlrd xlutils
```

### Usage
1. Clone the repository.
2. Ensure the script and the Excel file (`read.xls`) are in the same directory.
3. Run the script using Python:
    ```sh
    python gaussian_elimination.py
    ```
4. The script will read the coefficients from the Excel file, perform Gaussian Elimination, and write the results back into the Excel file.

### Code Explanation
The code begins by importing the necessary libraries and reading the matrix of coefficients and constants from the Excel file. It then performs Gaussian Elimination to transform the system into an upper triangular matrix, and back-substitutes to find the solutions. The results are written into a new sheet in the same Excel file.

Below is a snippet from the code illustrating the main logic:

```python
import xlrd
import numpy as np
from xlrd import open_workbook
from xlutils.copy import copy

workbook = xlrd.open_workbook('read.xls')
sheet = workbook.sheet_by_index(0)

arr = np.zeros((sheet.nrows, sheet.ncols))
arr2 = np.zeros((sheet.nrows + 1, sheet.ncols + 1))

# Taking excel data into numpy array
for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        arr[i][j] = sheet.cell_value(i, j)
        arr2[i][j] = sheet.cell_value(i, j)

# Calculate Echolon Matrix
for k in range(0, sheet.ncols):
    for i in range(k + 1, sheet.nrows):
        temp = arr2[k][k] / arr2[i][k]
        for j in range(k, sheet.ncols):
            arr2[i][j] = (temp) * arr2[i][j]
            arr2[i][j] = arr2[i][j] - arr2[k][j]

rb = open_workbook("read.xls")
wb = copy(rb)
sheet1 = wb.get_sheet(1)

# Clearing all data of excel
for i in range(100):
    for j in range(100):
        sheet1.write(i, j, '')

# Backstep Calculation
for k in range(sheet.nrows - 1, -1, -1):
    temp = 0
    for j in range(k + 1, sheet.nrows):
        temp = temp + (arr2[k][j] * arr2[j][sheet.ncols])
    arr2[k][sheet.ncols] = ((arr2[k][sheet.ncols - 1]) - (temp)) / arr2[k][k]

for i in range(sheet.nrows):
    for j in range(sheet.ncols + 1):
        sheet1.write(i, j, arr2[i][j])

wb.save('read.xls')
```

The code completes by saving the final results into the Excel file `read.xls`.

### Example
Below is an example of how to use the script:

1. Prepare the `read.xls` file with the system of equations in matrix form.
2. **Run the script**:
    ```sh
    python gaussian_elimination.py
    ```

3. **Output**:
    - The script will compute the results using the Gaussian Elimination method and write the intermediate and final results into the Excel file (`read.xls`).

### Files in the Repository
- `gaussian_elimination.py`: The main script for performing the Gaussian Elimination method.
- `read.xls`: Excel file from which the matrix data is read and into which the results are written.

### Input Parameters
The initial input data is expected to be in the form of a matrix within the `read.xls` file. Each row represents coefficients of the variables in the equations along with the constants.

### Troubleshooting
1. **Excel File**: Ensure that the input matrix is correct and placed in the `read.xls` file.
2. **Matrix Format**: Confirm that the matrix is complete and correctly formatted.
3. **Excel File Creation**: Ensure you have write permissions in the directory where the script is run to save the Excel file.
4. **Python Version**: This script is compatible with Python 3. Ensure you have Python 3 installed.

## Author
Script created by sudipto3331.

---

This documentation should guide you through understanding, installing, and using the Gaussian Elimination method script. For further issues or feature requests, please open an issue in the repository on GitHub. Feel free to contribute by creating issues and submitting pull requests. Happy coding!

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Tue May  2 15:26:14 2023

@author: sudipto3331
"""

import xlrd
import numpy as np
from xlrd import open_workbook
from xlutils.copy import copy


workbook = xlrd.open_workbook('read.xls')
sheet = workbook.sheet_by_index(0)

arr = np.zeros((sheet.nrows, sheet.ncols))
arr2 = np.zeros((sheet.nrows+1, sheet.ncols+1))


# Taking excel data into numpy array

for i in range(sheet.nrows):
    for j in range(sheet.ncols):
        arr[i][j] = sheet.cell_value(i,j)
        arr2[i][j] = sheet.cell_value(i,j)
        
# Calculate Echolon Matrix

for k in range(0, sheet.ncols):
    for i in range(k+1, sheet.nrows):
        temp = arr2[k][k]/arr2[i][k];
        for j in range(k, sheet.ncols):
            arr2[i][j] = (temp)*arr2[i][j]
            arr2[i][j] = arr2[i][j] - arr2[k][j]
            # print(arr2[i][j])


rb = open_workbook("read.xls")
wb = copy(rb)
sheet1 = wb.get_sheet(1)

# Clearing all data of excel

for i in range(100):
    for j in range(100):
        sheet1.write(i,j, '');

# Backstep Calculation

for k in range(sheet.nrows-1, -1, -1):
    temp = 0;
    for j in range(k+1, sheet.nrows):
        temp = temp + (arr2[k][j]*arr2[j][sheet.ncols]);
    arr2[k][sheet.ncols] = ((arr2[k][sheet.ncols-1]) - (temp))/arr2[k][k];
    

for i in range(sheet.nrows):
    for j in range(sheet.ncols+1):
        sheet1.write(i,j,arr2[i][j])

    
wb.save('read.xls')

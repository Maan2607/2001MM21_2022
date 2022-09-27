def octant_transition_count(mod=5000):
 from openpyxl import load_workbook
 import os
 import numpy as np
 os.system('cls')

 time1 = [] # make a new list
 v1 = []    # make a new list
 u1 = []    # make a new list
 w1 = []    # make a new list
 v2 = []    # make a list for difference 
 u2 = []    # make a list for difference
 w2 = []    # make a list for difference
 oct = []   # make a list for octant storing

 wb=load_workbook("input_octant_transition_identify.xlsx")
 sheet=wb["sheet1"] 
 n=sheet.max_row
 for r in range(2,n+1):
  time1.append(sheet.cell(row=r,column=1).value)
  u1.append(sheet.cell(row=r,column=2).value)
  v1.append(sheet.cell(row=r,column=3).value)
  w1.append(sheet.cell(row=r,column=4).value)

 v_mean=np.mean(v1, dtype=np.float64) # calculate the mean of v1
 u_mean=np.mean(u1, dtype=np.float64) # mean of u1
 w_mean=np.mean(w1, dtype=np.float64) # mean of w1
 


mod=5000
octant_transition_count(mod)
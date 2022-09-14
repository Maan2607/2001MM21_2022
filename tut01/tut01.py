def octact_identification(mod):
 import csv

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
 with open('octant_input.csv', 'r') as file:
  reader = csv.reader(file)
  j = 0
  for row in reader:
   if (j == 0):
    j = j+1 #continue only for j==0
   else:
    time1.append(row[0])
    u1.append(row[1])
    v1.append(row[2])
    w1.append(row[3])
  v_mean=np.mean(v1, dtype=np.float64) # it will give mean of v1
  u_mean=np.mean(u1, dtype=np.float64) # it will give mean of u1
  w_mean=np.mean(w1, dtype=np.float64) # it will give mean of w1

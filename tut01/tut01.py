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
 with open('octant_input.csv', 'r') as file: # again opening the input file
  reader = csv.reader(file)
  m=0
  for row in reader:
   if (m==0):
    m=m+1
   else: # it will not work only if m=0
    u2.append(float(row[1])-u_mean) # pushing the differences 
    v2.append(float(row[2])-v_mean) # pushing the differences
    w2.append(float(row[3])-w_mean)	# pushing the differences

 with open('octant_input.csv', 'r') as file: # again opening the input file for counting rows 
  reader = csv.reader(file)
  i=0
  for row in reader:
   i=i+1
  n=i #n here is number of rows
  n=n-1 # updating the  n with no. of entries
 with open('octant_input.csv', 'r') as file: # again opening the input file for octant 
  reader = csv.reader(file)  
  for i in range(n):# using this for loop we will get octant
   if (i<n-1):
    if ((u2[i]>=0) and (v2[i]>=0) and (w2[i]>=0 )):
     oct.append(1)
    elif((u2[i]>=0) and (v2[i]>=0) and (w2[i]<0 )):
     oct.append(-1)
    elif((u2[i]<0) and (v2[i]>=0) and (w2[i]>=0 )):
     oct.append(2)
    elif((u2[i]<0) and (v2[i]>=0) and (w2[i]<0 )):
     oct.append(-2)
    elif((u2[i]>=0) and (v2[i]<0) and (w2[i]>=0 )):
     oct.append(4)
    elif((u2[i]>=0) and (v2[i]<0) and (w2[i]<0 )):
     oct.append(-4)    
    elif((u2[i]<0) and (v2[i]<0) and (w2[i]>=0 )):
     oct.append(3)
    else:
     oct.append(-3) 

 

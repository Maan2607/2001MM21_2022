
from datetime import datetime
from unicodedata import name
start_time = datetime.now()

def attendance_report():
###Code
 import csv
 import os
 import numpy as np
 os.system('cls')
 roll_no=[]
 stu_name=[]

 with open('input_registered_students.csv', 'r') as file:
  reader = csv.reader(file)
  n=0
  for row in reader:
   if n!=0:
    roll_no.append(row[0])
    stu_name.append(row[1])
   n=n+1 
  n=n-1
 dat =["28/07","01/08","04/08","08/08","11/08","18/08","22/08","25/08","29/08","01/09","05/09","08/09","12/09"]

 real_count={}
 fake_count={}
 actual_count={}

 for x in range(0,n):
  real_count[roll_no[x]]=0 
  fake_count[roll_no[x]]=0
  actual_count[roll_no[x]]=0
 with open('input_attendance.csv', 'r') as file:
  reader = csv.reader(file) 
  x=0
  for row in reader: 
   if x==0:
    x = x+1
    continue
   else:
    if (row[0][0:5] in dat) and row[0][11:13]=="14" :
     if row[1][0:8] in roll_no:  
      real_count.update({row[1][0:8]: real_count.get(row[1][0:8])+1})
    elif row[1][0:8] in roll_no :
     fake_count.update({row[1][0:8]: fake_count.get(row[1][0:8])+1})
 
  for x in roll_no:
   for j in dat:
    with open('input_attendance.csv', 'r') as file:
     reader = csv.reader(file)
     for row in reader:
      if j==row[0][0:5] and row[1][0:8]==x and row[0][11:13]=="14" : 
       actual_count.update({row[1][0:8]: actual_count.get(row[1][0:8])+1})
       break 


 print(fake_count["2001CB02"])
 print(actual_count["2001CB02"])
attendance_report()




#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))

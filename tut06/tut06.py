

## THIS CODE WILL TAKE A BIT LONGER TIME APPROX. = 1 MINUTE
from datetime import datetime
from unicodedata import name
start_time = datetime.now()

def attendance_report():
###Code
 import csv #importing csv
 import os  #importing os 
 import numpy as np # importing numpy
 os.system('cls')
 roll_no=[] #declaring a list of registered roll NO.
 stu_name=[] #declaring a list of registered Name

 with open('input_registered_students.csv', 'r') as file:
  reader = csv.reader(file)
  n=0
  for row in reader:
   if n!=0:
    roll_no.append(row[0])  #making a list of registered roll NO.
    stu_name.append(row[1])   #making a list of registered Name
   n=n+1 
  n=n-1
 dat =["28/07","01/08","04/08","08/08","11/08","15/08","18/08","22/08","25/08","29/08","01/09","05/09","08/09","12/09","15/09","26/09","29/09"]
 # dat list will give all the dates on which classes were conducted
 s=len(dat) 


 with open('input_attendance.csv', 'r') as file:
  reader = csv.reader(file) 

  m=[] # This list will give full data of all Students
  c=[] # This list will give full data of a particular Student
  d=[] # This list will give full data of a particular Student on a particular date 
  for x in roll_no:
   c=[] # Initializing c
   for j in dat:
    d=[] # Initializing d
    a=0
    b=0
    e=0
    with open('input_attendance.csv', 'r') as file:
     reader = csv.reader(file)
     for row in reader:
      if j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00" ) and a==0 : 
       a=a+1 # counting real Attendance 
      elif j==row[0][0:5] and row[1][0:8]==x and (row[0][11:13]=="14" or row[0][11:16]=="15:00") and a>0 : 
       b=b+1 # counting duplicate Attendance 
      elif j==row[0][0:5] and row[1][0:8]==x :
       e=e+1 # counting fake Attendance 
     d.append(a+b+e)  
     d.append(a)
     d.append(b)
     d.append(e)
     if a==0:
      d.append(1)
     else:
      d.append(0) 
    c.append(d)
   m.append(c)
 
 if os.path.exists("output"):
  for f in os.listdir("output"):
    os.remove(os.path.join("output",f)) # removing all the prebuild files in output folder

 os.chdir("output")
 from openpyxl import Workbook
 for i in range(0,n): 
  book=Workbook()
  sheet= book.active    
  rows=[] # Making of list of rows of a particular student of all dates
  rows.append(["Date","Roll No.","Name","Total attendance count","Real","Duplicate","Invalid","absent"])
  rows.append(["",roll_no[i],stu_name[i],"","","","",""])
  for q in range(0,s):
   rows.append([dat[q],"","",m[i][q][0],m[i][q][1],m[i][q][2],m[i][q][3],m[i][q][4]]) # Appending all types of Attendance in row
  for w in rows:
   sheet.append(w)
  book.save( roll_no[i] + ".xlsx") # Saving a file with the name of roll no. for each rgistered student
   
  dic={0:"A",1:"P"} # Dictionary for present and absent
  book=Workbook()
  sheet= book.active    
  rows=[]
  z=["Roll No.","Name"] # Initializing 1st row
  for i in dat: # Initializing 1st row
   z.append(i)
  z.append("Total Lecture taken")
  z.append("Total Real")
  z.append("% Attendance")
  rows.append(z) 

  z=["(Sorted by roll no.)","","Atleast one real P"] # Initializing 2nd row
  for i in range(0,s-1): # using this loop we are making 2nd row
   z.append("")
  z.append("(=Total Mon+Thur dynamic count")
  z.append("")
  z.append("Real/Actual Lecture taken")
  rows.append(z) 

  for i in range(0,n):  # using this loop i am making full data of all the students
   z=[roll_no[i],stu_name[i]]
   c=0
   for q in range(0,s):
    z.append(dic[m[i][q][1]])
    if dic[m[i][q][1]]=="P":
     c=c+1
   z.append(s)
   z.append(c)
   l=(c/s)*100
   z.append("{:.2f}".format(l))
   rows.append(z)

  for w in rows:
   sheet.append(w)
  book.save( "Attendance_report_consolidated" + ".xlsx")  # Making of a full report of all the students

attendance_report()


#This shall be the last lines of the code.
end_time = datetime.now()
print('Duration of Program Execution: {}'.format(end_time - start_time))

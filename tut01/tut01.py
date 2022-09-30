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


    ### NOW WE WILL WRITE THE FILE


 with open('octant_input.csv', 'r') as file: # again opening the input file for overall count of octant 
  reader = csv.reader(file)  
  i=0
  c1=0 # overall count of octant 1
  c2=0 # overall count of octant -1
  c3=0 # overall count of octant 2
  c4=0 # overall count of octant -2
  c5=0 # overall count of octant 3
  c6=0 # overall count of octant -3
  c7=0 # overall count of octant 4
  c8=0 # overall count of octant -4	 
  for i in range(n):# using this for loop we will get octant
   if (i<n-1):
    if (oct[i]==1):
     c1=c1+1
    elif(oct[i]==-1):
     c2=c2+1
    elif(oct[i]==2):
     c3=c3+1
    elif(oct[i]==-2):
     c4=c4+1
    elif(oct[i]==3):
     c5=c5+1
    elif(oct[i]==-3):
     c6=c6+1    
    elif(oct[i]==4):
     c7=c7+1
    else:
     c8=c8+1
    i=i+1  
 o1=[] #make a list of octant count in mod range 
 o2=[] #make a list of octant count in mod range
 o3=[] #make a list of octant count in mod range
 o4=[] #make a list of octant count in mod range
 o5=[] #make a list of octant count in mod range
 o6=[] #make a list of octant count in mod range
 o7=[] #make a list of octant count in mod range
 o8=[] #make a list of octant count in mod range
 with open('octant_input.csv', 'r') as file: # again opening the input file for count of octant in mod range
  reader = csv.reader(file)
  i=0
  j=0
  oct1=0
  oct2=0
  oct3=0
  oct4=0
  oct5=0
  oct6=0
  oct7=0
  oct8=0
  m= int((n-1)/mod) +1
  for i in range(m):
   s=mod*i
   for j in range(mod):	
    if(j+s<n-1):	                                    
     if (oct[j+s]==1):
      oct1=oct1+1
     elif(oct[j+s]==-1):
      oct2=oct2+1
     elif(oct[j+s]==2):
      oct3=oct3+1
     elif(oct[j+s]==-2):
      oct4=oct4+1
     elif(oct[j+s]==3):
      oct5=oct5+1
     elif(oct[j+s]==-3):
      oct6=oct6+1    
     elif(oct[j+s]==4):
      oct7=oct7+1
     elif(oct[j+s]==-4):
      oct8=oct8+1 
   o1.append(oct1)
   o2.append(oct2)
   o3.append(oct3)
   o4.append(oct4)
   o5.append(oct5)
   o6.append(oct6)
   o7.append(oct7)
   o8.append(oct8)
   oct1=0
   oct2=0
   oct3=0
   oct4=0
   oct5=0
   oct6=0
   oct7=0
   oct8=0


 with open('octant_output.csv','w',newline="") as file: # opening the file to write the output
  writer=csv.writer(file)
  writer.writerow(["Time",'U','V','W','Uavg','Vavg','Wavg',"U'=U-Uavg","V'=V-Vavg","W'=W-Wavg","Octant","","OctantID","+1","-1","+2","-2","+3","-3","+4","-4"])
  j=0
  for x in range(n-1):
   if(x==0):   
    writer.writerow([time1[x],u1[x],v1[x],w1[x],u_mean,v_mean,w_mean,u2[x],v2[x],w2[x],oct[x],"","Overall count",c1,c2,c3,c4,c5,c6,c7,c8])
   elif(x==1):
    s="mod "+str(mod)		
    writer.writerow([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"User input",s,"","","","","","","",""])
   elif(x>=2 and x<2+m):
    if(x==1+m):# it will work ont=ly if x=1+m
     z=j*mod 	
     y=(j+1)*mod
     s=str(z)+"-"+str(n)		 
     writer.writerow([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"",s,o1[x-2],o2[x-2],o3[x-2],o4[x-2],o5[x-2],o6[x-2],o7[x-2],o8[x-2]])	 
     j=j+1
    else:
     z=j*mod	
     y=(j+1)*mod-1
     s=str(z)+"-"+str(y)		 
     writer.writerow([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"",s,o1[x-2],o2[x-2],o3[x-2],o4[x-2],o5[x-2],o6[x-2],o7[x-2],o8[x-2]])	 
     j=j+1
   else:
    writer.writerow([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"","","","","","","","","",""])			 
mod = 5000 # given input
octact_identification(mod) # given function

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
 sheet=wb["Sheet1"] 
 n=sheet.max_row
 for r in range(2,n+1):
  time1.append(sheet.cell(row=r,column=1).value)
  u1.append(sheet.cell(row=r,column=2).value)
  v1.append(sheet.cell(row=r,column=3).value)
  w1.append(sheet.cell(row=r,column=4).value)

 v_mean=np.mean(v1, dtype=np.float64) # calculate the mean of v1
 u_mean=np.mean(u1, dtype=np.float64) # mean of u1
 w_mean=np.mean(w1, dtype=np.float64) # mean of w1
 
 for r in range(2,n+1):
  u2.append(float(sheet.cell(row=r,column=2).value)-u_mean) # pushing the differences 
  v2.append(float(sheet.cell(row=r,column=3).value)-v_mean) # pushing the differences
  w2.append(float(sheet.cell(row=r,column=4).value)-w_mean)	# pushing the differences  

 for i in range(0,n-1):# using this for loop we will get octant
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
     
 c1=0 # overall count of octant 1
 c2=0 # overall count of octant -1
 c3=0 # overall count of octant 2
 c4=0 # overall count of octant -2
 c5=0 # overall count of octant 3
 c6=0 # overall count of octant -3
 c7=0 # overall count of octant 4
 c8=0 # overall count of octant -4	

 for i in range(0,n-1):# using this for loop we will get octant
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

 o1=[] #make a list of octant count in mod range 
 o2=[] #make a list of octant count in mod range
 o3=[] #make a list of octant count in mod range
 o4=[] #make a list of octant count in mod range
 o5=[] #make a list of octant count in mod range
 o6=[] #make a list of octant count in mod range
 o7=[] #make a list of octant count in mod range
 o8=[] #make a list of octant count in mod range

 oct1=0 #making variables for using in function 
 oct2=0 #making variables for using in function
 oct3=0 #making variables for using in function
 oct4=0 #making variables for using in function
 oct5=0 #making variables for using in function
 oct6=0 #making variables for using in function
 oct7=0 #making variables for using in function
 oct8=0 #making variables for using in function

 m= int((n-2)/mod) +1

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

 r=[]
 req=[]
 for i in range(0,m) :
  p=i*mod
  a1=[0]*8
  a2=[0]*8
  a3=[0]*8
  a4=[0]*8
  a5=[0]*8
  a6=[0]*8
  a7=[0]*8
  a8=[0]*8 
  for j in range(0,mod):
   if j+p<=n-3:
    if oct[j+p]==1 and oct[j+p+1]==1 :
      a1[0]=a1[0]+1
    elif oct[j+p]==1 and oct[j+p+1]==-1:
      a1[1]=a1[1]+1    
    elif oct[j+p]==1 and oct[j+p+1]==2:
      a1[2]=a1[2]+1
    elif oct[j+p]==1 and oct[j+p+1]==-2:
      a1[3]=a1[3]+1
    elif oct[j+p]==1 and oct[j+1+p]==3:
      a1[4]=a1[4]+1 
    elif oct[j+p]==1 and oct[j+p+1]==-3:
      a1[5]=a1[5]+1
    elif oct[j+p]==1 and oct[j+1+p]==4:
      a1[6]=a1[6]+1
    elif oct[j+p]==1 and oct[j+p+1]==-4:
      a1[7]=a1[7]+1
  
    elif oct[j+p]==-1 and oct[j+p+1]==1 :
      a2[0]=a2[0]+1
    elif oct[j+p]==-1 and oct[j+p+1]==-1:
      a2[1]=a2[1]+1    
    elif oct[j+p]==-1 and oct[j+p+1]==2:
      a2[2]=a2[2]+1
    elif oct[j+p]==-1 and oct[j+p+1]==-2:
      a2[3]=a2[3]+1
    elif oct[j+p]==-1 and oct[j+p+1]==3:
      a2[4]=a2[4]+1 
    elif oct[j+p]==-1 and oct[j+1+p]==-3:
      a2[5]=a2[5]+1
    elif oct[j+p]==-1 and oct[j+p+1]==4:
      a2[6]=a2[6]+1
    elif oct[j+p]==-1 and oct[j+1+p]==-4:
      a2[7]=a2[7]+1
    
    elif oct[j+p]==2 and oct[j+p+1]==1 :
      a3[0]=a3[0]+1
    elif oct[j+p]==2 and oct[j+p+1]==-1:
      a3[1]=a3[1]+1    
    elif oct[j+p]==2 and oct[j+p+1]==2:
      a3[2]=a3[2]+1
    elif oct[j+p]==2 and oct[j+p+1]==-2:
      a3[3]=a3[3]+1
    elif oct[j+p]==2 and oct[j+p+1]==3:
      a3[4]=a3[4]+1 
    elif oct[j+p]==2 and oct[j+1+p]==-3:
      a3[5]=a3[5]+1
    elif oct[j+p]==2 and oct[j+p+1]==4:
      a3[6]=a3[6]+1
    elif oct[j+p]==2 and oct[j+1+p]==-4:
      a3[7]=a3[7]+1 

    elif oct[j+p]==-2 and oct[j+p+1]==1 :
      a4[0]=a4[0]+1
    elif oct[j+p]==-2 and oct[j+p+1]==-1:
      a4[1]=a4[1]+1    
    elif oct[j+p]==-2 and oct[j+p+1]==2:
      a4[2]=a4[2]+1
    elif oct[j+p]==-2 and oct[j+p+1]==-2:
      a4[3]=a4[3]+1
    elif oct[j+p]==-2 and oct[j+p+1]==3:
      a4[4]=a4[4]+1 
    elif oct[j+p]==-2 and oct[j+1+p]==-3:
      a4[5]=a4[5]+1
    elif oct[j+p]==-2 and oct[j+p+1]==4:
      a4[6]=a4[6]+1
    elif oct[j+p]==-2 and oct[j+1+p]==-4:
      a4[7]=a4[7]+1  
  
    elif oct[j+p]==3 and oct[j+p+1]==1 :
      a5[0]=a5[0]+1
    elif oct[j+p]==3 and oct[j+p+1]==-1:
      a5[1]=a5[1]+1    
    elif oct[j+p]==3 and oct[j+p+1]==2:
      a5[2]=a5[2]+1
    elif oct[j+p]==3 and oct[j+p+1]==-2:
      a5[3]=a5[3]+1
    elif oct[j+p]==3 and oct[j+p+1]==3:
      a5[4]=a5[4]+1 
    elif oct[j+p]==3 and oct[j+1+p]==-3:
      a5[5]=a5[5]+1
    elif oct[j+p]==3 and oct[j+p+1]==4:
      a5[6]=a5[6]+1
    elif oct[j+p]==3 and oct[j+1+p]==-4:
      a5[7]=a5[7]+1 
    
    elif oct[j+p]==-3 and oct[j+p+1]==1 :
      a6[0]=a6[0]+1
    elif oct[j+p]==-3 and oct[j+p+1]==-1:
      a6[1]=a6[1]+1    
    elif oct[j+p]==-3 and oct[j+p+1]==2:
      a6[2]=a6[2]+1
    elif oct[j+p]==-3 and oct[j+p+1]==-2:
      a6[3]=a6[3]+1
    elif oct[j+p]==-3 and oct[j+p+1]==3:
      a6[4]=a6[4]+1 
    elif oct[j+p]==-3 and oct[j+1+p]==-3:
      a6[5]=a6[5]+1
    elif oct[j+p]==-3 and oct[j+p+1]==4:
      a6[6]=a6[6]+1
    elif oct[j+p]==-3 and oct[j+1+p]==-4:
      a6[7]=a6[7]+1

    elif oct[j+p]==4 and oct[j+p+1]==1 :
      a7[0]=a7[0]+1
    elif oct[j+p]==4 and oct[j+p+1]==-1:
      a7[1]=a7[1]+1    
    elif oct[j+p]==4 and oct[j+p+1]==2:
      a7[2]=a7[2]+1
    elif oct[j+p]==4 and oct[j+p+1]==-2:
      a7[3]=a7[3]+1
    elif oct[j+p]==4 and oct[j+p+1]==3:
      a7[4]=a7[4]+1
    elif oct[j+p]==4 and oct[j+1+p]==-3:
      a7[5]=a7[5]+1
    elif oct[j+p]==4 and oct[j+p+1]==4:
      a7[6]=a7[6]+1
    elif oct[j+p]==4 and oct[j+1+p]==-4:
      a7[7]=a7[7]+1 

    elif oct[j+p]==-4 and oct[j+p+1]==1 :
      a8[0]=a8[0]+1
    elif oct[j+p]==-4 and oct[j+p+1]==-1:
      a8[1]=a8[1]+1    
    elif oct[j+p]==-4 and oct[j+p+1]==2:
      a8[2]=a8[2]+1
    elif oct[j+p]==-4 and oct[j+p+1]==-2:
      a8[3]=a8[3]+1
    elif oct[j+p]==-4 and oct[j+p+1]==3:
      a8[4]=a8[4]+1 
    elif oct[j+p]==-4 and oct[j+1+p]==-3:
      a8[5]=a8[5]+1
    elif oct[j+p]==-4 and oct[j+p+1]==4:
      a8[6]=a8[6]+1
    elif oct[j+p]==-4 and oct[j+1+p]==-4:
      a8[7]=a8[7]+1      
  req.append(a1)  
  req.append(a2) 
  req.append(a3)
  req.append(a4)
  req.append(a5)
  req.append(a6)
  req.append(a7)
  req.append(a8)
  c=req.copy()
  r.append(c)
  req.clear()
 
 from openpyxl import Workbook 
 book=Workbook()
 sheet= book.active    
 rows=[] 
 rows.append(["Time",'U','V','W','Uavg','Vavg','Wavg',"U'=U-Uavg","V'=V-Vavg","W'=W-Wavg","Octant","","OctantID","+1","-1","+2","-2","+3","-3","+4","-4"])
 j=0
 a=0
 b=0
 print(r)
 for x in range(n-2):
  if(x==0):
   rows.append([time1[x],u1[x],v1[x],w1[x],u_mean,v_mean,w_mean,u2[x],v2[x],w2[x],oct[x],"","Overall count",c1,c2,c3,c4,c5,c6,c7,c8])
  elif(x==1):
   s="mod "+str(mod)		
   rows.append([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"User input",s,"","","","","","","",""])
  elif(x>=2 and x<2+m):
   if(x==1+m):# it will work ont=ly if x=1+m
    z=j*mod 	
    y=(j+1)*mod
    s=str(z)+"-"+str(n)		 
    rows.append([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"",s,o1[x-2],o2[x-2],o3[x-2],o4[x-2],o5[x-2],o6[x-2],o7[x-2],o8[x-2]])	 
    j=j+1
   else:
    z=j*mod	
    y=(j+1)*mod-1
    s=str(z)+"-"+str(y)		 
    rows.append([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"",s,o1[x-2],o2[x-2],o3[x-2],o4[x-2],o5[x-2],o6[x-2],o7[x-2],o8[x-2]])	 
    j=j+1
    # elif x<30:
    #  rows.append([time1[x],u1[x],v1[x],w1[x],"","","",u2[x],v2[x],w2[x],oct[x],"",r[a][b][0],r[a][b][1],r[a][b][2],r[a][b][3],r[a][b][4],r[a][b][5],r[a][b][6],r[a][b][7]])	
    #  b = b+1
    #  if b==8:
    #   b=0
    #   if a<m:
    #   a = a+1

 

 for i in rows:
  sheet.append(i)
 book.save("output_octant_transition_identify1.xlsx")
 
      
 
 
print("ho")

mod=5000
octant_transition_count(mod)
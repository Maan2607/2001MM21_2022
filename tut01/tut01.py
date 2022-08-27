def factorial(x):
 fa =1
 if x==0:
   print("Factorial = 1")
 elif x<0:
   print("Invalid Number")
 else:
   for i in range (1,x+1):
      fa=fa*i     
 print("factorial = ",fa,)  

x=int(input("Enter the number whose factorial is to be found"))
factorial(x)

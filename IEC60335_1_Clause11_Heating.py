#/bin/python
#coding:utf-8

def clause11():
    R1=float(input("Please input R1:"))
    R2=float(input("Please input R2:"))
    t1=float(input("Please input t1:"))
    t2=float(input("Please input t2:"))
    print("k is equal to the following:\n(a) 225 for aluminium windings and copper/aluminium windings with an aluminium content >=85%\n(b) 229,75 for copper/aluminium windings with an copper content between 15% and 85%\n(c) 234,5 for copper windings and copper/aluminium windings with an copper content >=85%")
    k=float(input("Please input k:"))
    result=(R2-R1)/R1*(k+t1)-(t2-t1)
    print("The temperature rise:"+str(result))

while True:
    try:
        clause11()
        flag=input("Press ENTER to continue! Others to EXIT!")
        if flag!="":
            break
        else:
            print("================another calculation================")
    except:
        print("\nWarning: An error occurs! Please contact author.")
        break
    

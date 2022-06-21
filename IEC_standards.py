#!/bin/python
import os

Version="Name: IEC-search V0.1"
Description="Description: This software is used to searh the IEC standards.The database will be updated as soon as new standards publish."
line="-------------------------------------"
while True:
    print(line+"INFO"+line)
    print(Version)
    print(Description)
    print(line+"----"+line)
    try:
        para=input("Please input the standard:")
        cmd='find "'+para+'" .\Database\std.db'
        test=os.system(cmd)
        flag=input("Press Enter to continue! Others to EXIT!")
        if flag!="":
            break
        else:
            os.system("cls")
    except:
        print("\nWarning: There's an error! Please contact author.")
        break

#!/bin/python
import os

Version="Name: G-search V0.1"
Description="Description: This software is used to searh the GSO standards.The database will be updated as soon as new standards publish."
line="-------------------------------------"
while True:
    print(line+"INFO"+line)
    print(Version)
    print(Description)
    print(line+"----"+line)
    try:
        para=raw_input("Please input the standard:")
        cmd='find "'+para+'" .\Database\GSO.db'
        test=os.system(cmd)
        flag=raw_input("Press Enter to continue! Others to EXIT!")
        if flag!="":
            break
        else:
            os.system("cls")
    except:
        print("\nWarning: There's an error! Please contact author.")
        break

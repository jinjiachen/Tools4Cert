#!/bin/python
import os

Version="Name: EC-search V0.1"
Description="Description: This software is used to search the equipment lists quickly.The database will be synchronized with the lists."
while True:
    try:
        para=input("Please input the EC number:")
        cmd='find "'+para+'" .\Database\Equipment_List.db'
        os.system(cmd)
        flag=input("Press Enter to continue! Others to EXIT!")
        if flag!="":
            break
        else:
            os.system("cls")
    except:
        print("\nWarning: An error occurs! Please contact author.")
        break

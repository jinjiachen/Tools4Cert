#!/bin/python
import os

Version="Name: EC-search V0.1"
Description="Description: This software is used to search the equipment lists quickly.The database will be synchronized with the lists."
while True:
    try:
        para=raw_input("Please input the EC number:")
        cmd='find "'+para+'" .\Database\Equipment_List.db'
        os.system(cmd)
        flag=raw_input("Press Enter to exit! Others to continue!")
        if flag!="":
            break
        else:
            os.system("cls")
    except:
        print("\nWarning: An error occurs! Please contact author.")
        break

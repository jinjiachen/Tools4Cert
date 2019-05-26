#!/bin/python
import os

while True:
    para=raw_input("Please input the EC number:")
    cmd='find "'+para+'" .\Database\Equipment_List.db'
#    print cmd
    os.system(cmd)
    input("Press Enter to exit! Others to continue!")
    os.system("cls")

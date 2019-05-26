#!/bin/python
import os

while True:
    para=raw_input("Please input the standard:")
    cmd='find "'+para+'" E:\Python27\Tools4Cert-master\GSO.txt'
#    print cmd
    os.system(cmd)
    input("Press Enter to exit! Others to continue!")
    os.system("cls")

#!/bin/env python
#coding=utf8

Tc_Input=("26.7/19.4","35/23.9","26.7/19.4","23.9/35","-","-","21.1","15.6/7.2","25","21.1/14.7","21.1/14.7","-")
Tc_Temp_Pressure=("40/26.7","40/26.7","40/26.7","26.7/37.8","21.1/14.7","21.1/14.7","21.1","21.1/12.8","25","21.1/14.7","21.1/14.7","25")
#print len(Tc_Temp_Pressure)

print "=============Menu============="
print "1.cooling"
print "2.Reverse cycle heating"
print "3.Resistance heat only" 
print "4.Combination reverse cycle-resistance heat"
print "5.Steam or hot water"
#print "Please choose the number:1-5:"
choice1=raw_input("Please choose the number:1-5\n")

print "=============Result============"
if choice1=="1":
    print "Air cooling or water cooling?"
    choice2=raw_input("a for Air cooling;w for water cooling,please choose:\n")
    if choice2=="a":
        print "Input Test:"
        print "Indoor air(DB/WB):",Tc_Input[0]
        print "Outdoor air(DB/WB):",Tc_Input[1]
        print "Temperature and pressure test:"
        print "Indoor air(DB/WB):",Tc_Temp_Pressure[0]
        print "Outdoor air(DB/WB):",Tc_Temp_Pressure[1]
    if choice2=="w":
        print "Input Test:"
        print "Indoor air(DB/WB):",Tc_Input[2]
        print "Outdoor air(DB/WB):",Tc_Input[3]
        print "Temperature and pressure test:"
        print "Indoor air(DB/WB):",Tc_Temp_Pressure[2]
        print "Outdoor air(DB/WB):",Tc_Temp_Pressure[3]
elif choice1=="2":
    choice2=raw_input("a for Air cooling;w for water cooling,please choose:\n")
    if choice2=="a":
        print "Input Test:"
        print "Indoor air(DB/WB):",Tc_Input[4]
        print "Outdoor air(DB/WB):",Tc_Input[5]
        print "Temperature and pressure test:"
        print "Indoor air(DB/WB):",Tc_Temp_Pressure[4]
        print "Outdoor air(DB/WB):",Tc_Temp_Pressure[5]
    if choice2=="w":
        print "Input Test:"
        print "Indoor air(DB/WB):",Tc_Input[6]
        print "Outdoor air(DB/WB):",Tc_Input[7]
        print "Temperature and pressure test:"
        print "Indoor air(DB/WB):",Tc_Temp_Pressure[6]
        print "Outdoor air(DB/WB):",Tc_Temp_Pressure[7]
elif choice1=="3":
    print "Input Test:"
    print "Indoor air(DB/WB):",Tc_Input[8]
    print "Outdoor air(DB/WB):",Tc_Input[9]
    print "Temperature and pressure test:"
    print "Indoor air(DB/WB):",Tc_Temp_Pressure[8]
    print "Outdoor air(DB/WB):",Tc_Temp_Pressure[9]
elif choice1=="4":
    print "Input Test:"
    print "Indoor air(DB/WB):",Tc_Input[10]
    print "Outdoor air(DB/WB):",Tc_Input[11]
    print "Temperature and pressure test:"
    print "Indoor air(DB/WB):",Tc_Temp_Pressure[10]
    print "Outdoor air(DB/WB):",Tc_Temp_Pressure[11]
elif choice1=="5":
    print "Input Test:"
    print "Indoor air(DB/WB):",Tc_Input[12]
    print "Outdoor air(DB/WB):",Tc_Input[13]
    print "Temperature and pressure test:"
    print "Indoor air(DB/WB):",Tc_Temp_Pressure[12]
    print "Outdoor air(DB/WB):",Tc_Temp_Pressure[13]
print "==============The end=============="



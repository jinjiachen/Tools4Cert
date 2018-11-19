#!/bin/python
#coding:utf8

class LeakageCurrent(object):
    print "0: Class 0 appliance\n 01: Class 0I applicance\n 1: Class I appliance\n 2: Class II appliance\n 3: Class III appliance"
    choice=input("Please input 0-3:")
    if choice==2:
        print "Leakage Current < 0.35mA peak"
    elif choice==0 or choice==3:
        print "Leakage Current < 0.7mA peak"
    elif choice==01:
        print "Leakage Current < 0.5mA"
    elif choice==1:
        print "portable or stationary?"
        subchoice=raw_input("p or s")


if __name__=='__main__':
    LeakageCurrent()

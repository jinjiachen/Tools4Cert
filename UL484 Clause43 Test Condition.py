#!/bin/env python
#coding=utf8

Volt_Input=[115,208,230,'Rated','Rated','Rated']
Volt_Others=[120,208,240,277,480,600]
RateVolt=input("Please input the rated voltage:")
RateVolt=int(RateVolt)
if RateVolt>=110 and RateVolt<=120:
    print("Input test:",Volt_Input[0])
    print("All other tests:",Volt_Others[0])
elif RateVolt>=200 and RateVolt<=208:
    print("Input test:",Volt_Input[1])
    print("All other tests:",Volt_Others[1])
elif RateVolt>=220 and RateVolt<=240:
    print("Input test:",Volt_Input[2])
    print("All other tests:",Volt_Others[2])
elif RateVolt>=254 and RateVolt<=277:
    print("Input test:",Volt_Input[3])
    print("All other tests:",Volt_Others[3])
elif RateVolt>=440 and RateVolt<=480:
    print("Input test:",Volt_Input[4])
    print("All other tests:",Volt_Others[4])
elif RateVolt>=550 and RateVolt<=600:
    print("Input test:",Volt_Input[5])
    print("All other tests:",Volt_Others[5])

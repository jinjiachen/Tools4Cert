#!/bin/env python
#coding=utf8

import wx
import math

class TestVoltage(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,'UL484 Clause43 Test Voltage')
        self.panel=wx.Panel(self,-1)
#        self.catagory=('Class II appliances','Class 0 and III appliances','Class 0I appliances','Portable class I appliances','Stationary class I motor-operated appliances','Stationary class I heating appliances')
        wx.StaticText(self.panel,-1,"Please input the rated voltage:",pos=(50,50))
        self.text1=wx.TextCtrl(self.panel,-1,pos=(50,80))
#        wx.StaticText(self.panel,-1,"Please choose the appliance catagory:",pos=(50,50))
#        self.choice=wx.Choice(self.panel,-1,pos=(50,80),choices=self.catagory)
        self.Button=wx.Button(self.panel,-1,"Verdict",pos=(50,130))
#        self.Bind(wx.EVT_CHOICE,self.add,self.choice)
        self.Bind(wx.EVT_BUTTON,self.compute,self.Button)
#        self.stext1=wx.StaticText(self.panel,-1,"Input Test(V):",pos=(50,180))
#        self.stext2=wx.StaticText(self.panel,-1,"All other test(V):",pos=(50,210))
            
    def compute(self,event):
        Volt_Input=['115','208','230','Rated','Rated','Rated']
        Volt_Others=['120','208','240','277','480','600']
        voltage=float(self.text1.GetValue())#Rated Voltage
        if voltage>=110 and voltage<=120:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[0],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[0],pos=(50,210))
        elif voltage>=200 and voltage<=208:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[1],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[1],pos=(50,210))
        elif voltage>=220 and voltage<=240:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[2],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[2],pos=(50,210))
        elif voltage>=254 and voltage<=277:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[3],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[3],pos=(50,210))
        elif voltage>=440 and voltage<=480:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[4],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[4],pos=(50,210))
        elif voltage>=550 and voltage<=600:
            self.stext1=wx.StaticText(self.panel,-1,"Input Test(V): "+Volt_Input[5],pos=(50,180))
            self.stext2=wx.StaticText(self.panel,-1,"All other test(V): "+Volt_Others[5],pos=(50,210))

        

        


if __name__=='__main__':
    myapp=wx.PySimpleApp()
    myframe=TestVoltage()
    myframe.Show()
    myapp.MainLoop()


#!/bin/env python
#coding=utf8

import wx

class TestVoltageFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,'UL484 Clause43 Test Voltage')
        self.panel=wx.Panel(self,-1)
        wx.stext1=wx.StaticText(self.panel,-1,"Rated Voltage(V):",pos=(50,50))
        self.text1=wx.TextCtrl(self.panel,-1,pos=(150,50))
        self.button=wx.Button(self.panel,-1,"Result",pos=(50,80))
        self.Bind(wx.EVT_BUTTON,self.result,self.button)

    def result(self,event):
        RateVolt=float(self.text1.GetValue())
        Volt_Input=['115','208','230','Rated','Rated','Rated']
        Volt_Others=['120','208','240','277','480','600']
#        self.stext2=wx.StaticText(self.panel,-1,pos=(50,120))
#        self.stext3=wx.StaticText(self.panel,-1,pos=(50,150))
#        self.stext2.Show(False)
#        self.stext3.Show(False)
        if RateVolt>=110 and RateVolt<=120:
            res_Input=Volt_Input[0]
            res_Other=Volt_Others[0]
        elif RateVolt>=200 and RateVolt<=208:
            res_Input=Volt_Input[1]
            res_Other=Volt_Others[1]
        elif RateVolt>=220 and RateVolt<=240:
            res_Input=Volt_Input[2]
            res_Other=Volt_Others[2]
        elif RateVolt>=254 and RateVolt<=277:
            res_Input=Volt_Input[3]
            res_Other=Volt_Others[3]
        elif RateVolt>=440 and RateVolt<=480:
            res_Input=Volt_Input[4]
            res_Other=Volt_Others[4]
        elif RateVolt>=550 and RateVolt<=600:
            res_Input=Volt_Input[5]
            res_Other=Volt_Others[5]
        self.stext2=wx.StaticText(self.panel,-1,"Input test(V): "+res_Input,pos=(50,120))
        self.stext3=wx.StaticText(self.panel,-1,"Other tests(V): "+res_Other,pos=(50,150))

if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=TestVoltageFrame()
    myframe.Show()
    myapp.MainLoop()

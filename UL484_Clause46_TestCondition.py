#!/bin/env python
#coding=utf8

import wx

class TestConditionFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,"UL484 Clause46 Test Condition")
        self.panel=wx.Panel(self,-1)
        mode=[
                "Cooling",
                "Reverse cycle heating",
                "Resistance heat (only)",
                "Combination reverse cycle-resistance heat",
                "Steam or hot water"]
        submode=[
                "Air cooled unit",
                "Water cooled unit"
                ]
        test=[
                "Input test",
                "Temperature and pressure test"
             ]
        self.cho1=wx.RadioBox(self.panel,-1,"Mode:",choices=mode,style=wx.RA_SPECIFY_COLS,majorDimension=1)
        self.cho2=wx.RadioBox(self.panel,-1,"Submode",choices=submode,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(350,0))
        self.cho3=wx.RadioBox(self.panel,-1,"Test",choices=test,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(350,80))
        self.button=wx.Button(self.panel,-1,"Result",pos=(10,160))
        self.Bind(wx.EVT_RADIOBOX,self.li,self.cho1)
        self.Bind(wx.EVT_BUTTON,self.res,self.button)

    def li(self,event):
        if self.cho1.GetStringSelection()=="Steam or hot water":
            self.cho2.Show(False)
        elif self.cho1.GetStringSelection()=="Combination reverse cycle-resistance heat":
            self.cho2.Show(False)
            self.cho2.Show()
            self.cho2.ShowItem(1,False)
        else:
            self.cho2.Show(False)
            self.cho2.Show()

    def res(self,event):
        Tc_Input=("26.7/19.4","35/23.9","26.7/19.4","23.9/35","-","-","21.1","15.6/7.2","25","21.1/14.7","21.1/14.7","-")
        Tc_Temp_Pressure=("40/26.7","40/26.7","40/26.7","26.7/37.8","21.1/14.7","21.1/14.7","21.1","21.1/12.8","25","21.1/14.7","21.1/14.7","25")


if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=TestConditionFrame()
    myframe.Show()
    myapp.MainLoop()

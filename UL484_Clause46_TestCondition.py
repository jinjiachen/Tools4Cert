#!/bin/env python
#coding=utf8

import wx

class TestConditionFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,"UL484 Clause46 Test Condition",size=(600,350))
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
        elif self.cho1.GetStringSelection()=="Resistance heat (only)":
            self.cho2.Show(False)
        else:
            self.cho2.Show(False)
            self.cho2.Show()

    def res(self,event):
        Tc_Input=("26.7/19.4","35/23.9","26.7/19.4","23.9/35","-","-","21.1","15.6/7.2","25","21.1/14.7","21.1/14.7","-")
        Tc_Temp_Pressure=("40/26.7","40/26.7","40/26.7","26.7/37.8","21.1/14.7","21.1/14.7","21.1","21.1/12.8","25","21.1/14.7","21.1/14.7","25")
#Cooling mode
        if self.cho1.GetStringSelection()=="Cooling":
            if self.cho2.GetStringSelection()=="Air cooled unit":
                if self.cho3.GetStringSelection()=="Input test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Input[0]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Input[1]+"\t\t",pos=(10,220))
                elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Temp_Pressure[0]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Temp_Pressure[1]+"\t\t",pos=(10,220))
            if self.cho2.GetStringSelection()=="Water cooled unit":
                if self.cho3.GetStringSelection()=="Input test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Input[2]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Input[3]+"\t\t",pos=(10,220))
                elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Temp_Pressure[2]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Temp_Pressure[3]+"\t\t",pos=(10,220))
        #Reverse cycle heating
        elif self.cho1.GetStringSelection()=="Reverse cycle heating":
            if self.cho2.GetStringSelection()=="Air cooled unit":
                if self.cho3.GetStringSelection()=="Input test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Input[4]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Input[5]+"\t\t",pos=(10,220))
                elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Temp_Pressure[4]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Temp_Pressure[5]+"\t\t",pos=(10,220))
            if self.cho2.GetStringSelection()=="Water cooled unit":
                if self.cho3.GetStringSelection()=="Input test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Input[6]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Input[7]+"\t\t",pos=(10,220))
                elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                    wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Temp_Pressure[6]+"\t\t",pos=(10,200))
                    wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Temp_Pressure[7]+"\t\t",pos=(10,220))
        #Resistance heat (only)
        elif self.cho1.GetStringSelection()=="Resistance heat (only)":
            if self.cho3.GetStringSelection()=="Input test":
                wx.StaticText(self.panel,-1,'Air temperature,DB/WB:'+Tc_Input[8]+"\t",pos=(10,200))
            elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                wx.StaticText(self.panel,-1,'Air temperature,DB/WB:'+Tc_Temp_Pressure[8]+"\t",pos=(10,200))
        #Combination reverse cycle-resistance heat
        elif self.cho1.GetStringSelection()=="Combination reverse cycle-resistance heat":
            if self.cho3.GetStringSelection()=="Input test":
                wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Input[9]+"\t",pos=(10,200))
                wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Input[10]+"\t",pos=(10,220))
            elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                wx.StaticText(self.panel,-1,'Indoor air,DB/WB:'+Tc_Temp_Pressure[9]+"\t",pos=(10,200))
                wx.StaticText(self.panel,-1,'Outdoor air,DB/WB:'+Tc_Temp_Pressure[10]+"\t",pos=(10,220))
        #Steam or hot water
        elif self.cho1.GetStringSelection()=="Steam or hot water":
            if self.cho3.GetStringSelection()=="Input test":
                wx.StaticText(self.panel,-1,'Air temperature,DB/WB:'+Tc_Input[11]+"\t",pos=(10,200))
            elif self.cho3.GetStringSelection()=="Temperature and pressure test":
                wx.StaticText(self.panel,-1,'Air temperature,DB/WB:'+Tc_Temp_Pressure[11]+"\t",pos=(10,200))


if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=TestConditionFrame()
    myframe.Show()
    myapp.MainLoop()

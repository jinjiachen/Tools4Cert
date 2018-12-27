#!/bin/env python
#coding=utf8

import wx

class HVCircuitsFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,"UL484 Section33 High-Voltage Circuits",size=(600,350))
        self.panel=wx.Panel(self,-1)
        VA=[
                "2000 or less",
                "More than 2000"
            ]
        Volts1=[
                "300 or less",
                "301 - 600",
                ]
        Volts2=[
                "150 or less",
                "151 - 300",
                "301 - 600"
                ]
        Spacing=[
                "Through air",
                "Over surface",
                "To enclosure"
                ]
        self.ThroughAir=[
                "3.2",
                "9.5",
                "3.2",
                "6.4",
                "9.5"
             ]
        self.OverSurface=[
                "6.4",
                "12.7",
                "6.4",
                "9.5",
                "12.7"
             ]
        self.ToEnclosure=[
                "6.4",
                "12.7",
                "12.7",
                "12.7",
                "12.7"
             ]
        self.cho1=wx.RadioBox(self.panel,-1,"Volt-Amperes:",choices=VA,style=wx.RA_SPECIFY_COLS,majorDimension=1)
        self.cho2=wx.RadioBox(self.panel,-1,"Volts:",choices=Volts1,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(150,0))
        self.cho3=wx.RadioBox(self.panel,-1,"Volts:",choices=Volts2,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(150,0))
        self.cho3.Show(False) #initialize
        self.cho4=wx.RadioBox(self.panel,-1,"Minimum spacing:",choices=Spacing,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(250,0))
        self.button=wx.Button(self.panel,-1,"Search",pos=(10,100))
        self.Bind(wx.EVT_RADIOBOX,self.li,self.cho1)
        self.Bind(wx.EVT_BUTTON,self.res,self.button)

    def li(self,event):
        if self.cho1.GetStringSelection()=="2000 or less":
            self.cho2.Show(True)
            self.cho3.Show(False)
        else:
            self.cho2.Show(False)
            self.cho3.Show(True)

    def res(self,event):
        if self.cho1.GetStringSelection()=="2000 or less":
            if self.cho2.GetStringSelection()=="300 or less":
                if self.cho4.GetStringSelection()=="Through air":
                    result=self.ThroughAir[0]
                elif self.cho4.GetStringSelection()=="Over surface":
                    result=self.OverSurface[0]
                elif self.cho4.GetStringSelection()=="To enclosure":
                    result=self.ToEnclosure[0]
            elif self.cho2.GetStringSelection()=="301 - 600":
                if self.cho4.GetStringSelection()=="Through air":
                    result=self.ThroughAir[1]
                elif self.cho4.GetStringSelection()=="Over surface":
                    result=self.OverSurface[1]
                elif self.cho4.GetStringSelection()=="To enclosure":
                    result=self.ToEnclosure[1]
        elif self.cho1.GetStringSelection()=="More than 2000":
            if self.cho3.GetStringSelection()=="150 or less":
                if self.cho4.GetStringSelection()=="Through air":
                    result=self.ThroughAir[2]
                elif self.cho4.GetStringSelection()=="Over surface":
                    result=self.OverSurface[2]
                elif self.cho4.GetStringSelection()=="To enclosure":
                    result=self.ToEnclosure[2]
            elif self.cho3.GetStringSelection()=="151 - 300":
                if self.cho4.GetStringSelection()=="Through air":
                    result=self.ThroughAir[3]
                elif self.cho4.GetStringSelection()=="Over surface":
                    result=self.OverSurface[3]
                elif self.cho4.GetStringSelection()=="To enclosure":
                    result=self.ToEnclosure[3]
            elif self.cho3.GetStringSelection()=="301 - 600":
                if self.cho4.GetStringSelection()=="Through air":
                    result=self.ThroughAir[4]
                elif self.cho4.GetStringSelection()=="Over surface":
                    result=self.OverSurface[4]
                elif self.cho4.GetStringSelection()=="To enclosure":
                    result=self.ToEnclosure[4]
        wx.StaticText(self.panel,-1,result+"     ",pos=(10,150))


if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=HVCircuitsFrame()
    myframe.Show()
    myapp.MainLoop()

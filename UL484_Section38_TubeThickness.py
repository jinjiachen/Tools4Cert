#!/bin/env python
#coding=utf8

import wx

class TubeThicknessFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,"UL484 Section38 Refrigerant Tubing and Fittings",size=(600,350))
        self.panel=wx.Panel(self,-1)
        Material=[
                "Copper",
                "Steel",
                "Aluminum"
            ]
        sub=[
                "Protected",
                "Unprotected"
                ]
        OD=[
                "3/16",
                "1/4",
                "5/16",
                "3/8",
                "1/2",
                "5/8",
                "3/4",
                "7/8",
                "1",
                "1-1/8",
                "1-1/4",
                "1-3/8",
                "1-1/2",
                "1-5/8",
                "2-1/8",
                "2-5/8"
                ]
        self.CopperPro=[
                "0.622",
                "0.622",
                "0.622",
                "0.622",
                "0.622",
                "0.800",
                "0.800",
                "1.041",
                "1.168",
                "1.168",
                "1.283",
                "1.283",
                "1.410",
                "1.410",
                "1.626",
                "1.880"
                ]
        self.CopperUnpro=[
                "0.673",
                "0.673",
                "0.673",
                "0.673",
                "0.724",
                "0.800",
                "0.980",
                "1.041",
                "1.168",
                "1.168",
                "1.283",
                "1.283",
                "1.410",
                "1.410",
                "1.626",
                "1.880"
                ]
        self.Steel=[
                "0.64",
                "0.64",
                "0.64",
                "0.64",
                "0.64",
                "0.81",
                "0.81",
                "1.17",
                "-",
                "1.17",
                "1.17",
                "-",
                "1.57",
                "-",
                "-",
                "-"
                ]
        self.Al=[
                "0.89",
                "0.89",
                "0.89",
                "0.89",
                "0.89",
                "1.24",
                "1.24",
                "1.65",
                "1.83",
                "-",
                "-",
                "-",
                "-",
                "-",
                "-",
                "-",
                ]

                


        self.cho1=wx.RadioBox(self.panel,-1,"Material:",choices=Material,style=wx.RA_SPECIFY_COLS,majorDimension=1)
        self.cho2=wx.RadioBox(self.panel,-1,"Protected or unprotected:",choices=sub,style=wx.RA_SPECIFY_COLS,majorDimension=1,pos=(150,0))
        self.cho3=wx.StaticText(self.panel,-1,"Please choose outside diameter:",pos=(10,100))
        self.cho4=wx.Choice(self.panel,-1,choices=OD,pos=(10,120))
        self.button=wx.Button(self.panel,-1,"Search",pos=(10,150))
        self.Bind(wx.EVT_RADIOBOX,self.li,self.cho1)
        self.Bind(wx.EVT_BUTTON,self.res,self.button)

    def li(self,event):
        if self.cho1.GetStringSelection()=="Copper":
            self.cho2.Show(True)
        else:
            self.cho2.Show(False)

    def res(self,event):
        if self.cho1.GetStringSelection()=="Copper":
            if self.cho2.GetStringSelection()=="Protected":
                result=self.CopperPro[self.cho4.GetSelection()]
            elif self.cho2.GetStringSelection()=="Unprotected":
                result=self.CopperUnpro[self.cho4.GetSelection()]
        elif self.cho1.GetStringSelection()=="Steel":
                result=self.Steel[self.cho4.GetSelection()]
        elif self.cho1.GetStringSelection()=="Aluminum":
                result=self.Al[self.cho4.GetSelection()]
        wx.StaticText(self.panel,-1,result+"          ",pos=(10,180))
        Note="Exception: Copper or steel capillary tubing protected against mechanical damage \n by the cabinet or assembly shall have a wll thickness not less than 0.51mm"
        wx.StaticText(self.panel,-1,Note,pos=(10,210))


if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=TubeThicknessFrame()
    myframe.Show()
    myapp.MainLoop()

#!/bin/env python
#coding=utf8

import wx
import math

class TestCondition(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,'UL484 Clause46 Test Condition')
        self.panel=wx.Panel(self,-1)
#        self.catagory=('Class II appliances','Class 0 and III appliances','Class 0I appliances','Portable class I appliances','Stationary class I motor-operated appliances','Stationary class I heating appliances')
        wx.StaticText(self.panel,-1,"Please input the rated voltage:",pos=(50,150))
        self.text1=wx.TextCtrl(self.panel,-1,pos=(50,180))
        wx.StaticText(self.panel,-1,"Please choose the appliance catagory:",pos=(50,50))
        self.choice=wx.Choice(self.panel,-1,pos=(50,80),choices=self.catagory)
        self.Button=wx.Button(self.panel,-1,"Verdict",pos=(50,300))
        self.Bind(wx.EVT_CHOICE,self.add,self.choice)
        self.Bind(wx.EVT_BUTTON,self.compare,self.Button)

    def add(self,event):
        if self.choice.GetStringSelection()==self.catagory[5]:
            self.stext1=wx.StaticText(self.panel,-1,"Please input the power input(kW):",pos=(50,220))
            self.text2=wx.TextCtrl(self.panel,-1,pos=(50,250))
            
        else:
            self.text2.Show(False)
            self.stext1.Show(False)


    def compare(self,event):
        value_LC=float(self.text1.GetValue())
        value_Limited=[0.35,0.7,0.5,0.75,3.5]
        if self.choice.GetStringSelection()==self.catagory[5]:
            power=float(self.text2.GetValue())
            temp=min(power*0.75,5)
            max_LC=max(temp,0.75)
            value_Limited.append(max_LC)
#        print value_LC
#        print value_Limited[0]
#        print self.choice.GetSelection()
#        print self.catagory[0]
        for i in range(0,6):
            if self.choice.GetStringSelection()==self.catagory[i]:
                if value_LC<=value_Limited[i]:
                    wx.MessageBox("Pass! Leakage current complies.","Verdict",style=wx.OK)
                else:
                    wx.MessageBox("Fail! Leakage current does not complies.","Verdict",style=wx.OK)

        

        


if __name__=='__main__':
    myapp=wx.PySimpleApp()
    myframe=LeakageCurrentFrame()
    myframe.Show()
    myapp.MainLoop()


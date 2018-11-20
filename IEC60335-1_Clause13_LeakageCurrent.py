#!/bin/env python
#coding=utf8

import wx

class LeakageCurrentFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self,None,-1,'Leakage Current')
        panel=wx.Panel(self,-1)
        catagory=('Class II appliances','Class 0 and III appliances','Class 0I appliances','Portable class I appliances','Stationary class I motor-operated appliances','Stationary class I heating appliances')
        wx.StaticText(panel,-1,"Please choose the appliance catagory:",pos=(50,50))
        wx.Choice(panel,-1,pos=(50,80),choices=catagory)
        wx.StaticText(panel,-1,"Please input the measured leakage current:",pos=(50,150))
        self.LC=wx.TextCtrl(panel,-1,"test",pos=(50,180))

    def compare(self):
        self.LC.GetValue()
        


if __name__=='__main__':
    myapp=wx.PySimpleApp()
    myframe=LeakageCurrentFrame()
    myframe.Show()
    myapp.MainLoop()


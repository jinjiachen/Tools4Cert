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
        wx.RadioBox(self.panel,-1,"Mode:",choices=mode,style=wx.RA_SPECIFY_COLS,majorDimension=1)


if __name__=="__main__":
    myapp=wx.PySimpleApp()
    myframe=TestConditionFrame()
    myframe.Show()
    myapp.MainLoop()

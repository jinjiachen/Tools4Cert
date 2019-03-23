#!/bin/python
#coding=utf8

import wx
import IEC60335_1_Clause13_LeakageCurrent
import IEC60335_1_Clause16_LeakageCurrent
import UL484_Section33_HighVoltageCircuits
import UL484_Section38_TubeThickness
import UL484_Clause43_TestVoltage
import UL484_Clause46_TestCondition

class myframe(wx.Frame):    #Main frame

    def __init__(self):
        wx.Frame.__init__(self,None,-1,'Tool4Cert by Michael')
        panel=wx.Panel(self)
        mymenubar=wx.MenuBar()
        mymenu1=wx.Menu()   #File
        mymenu2=wx.Menu()   #CB
        mymenu3=wx.Menu()   #ETL
        mymenu4=wx.Menu()   #search
        mymenu5=wx.Menu()   #TDS
        Exit=mymenu1.Append(wx.NewId(),'&Exit')    
        text1=mymenu1.Append(-1,'Save')
        text2=mymenu2.Append(-1,'iEC60335-1:2010+A1:2013+A2:2016')
        text3=mymenu2.Append(-1,'iEC60335-2-40:2018')
        text4=mymenu3.Append(-1,'UL 484')
        text5=mymenu3.Append(-1,'UL 60335-2-40')
        text6=mymenu4.Append(-1,'UL')
        text7=mymenu4.Append(-1,'VDE')
        text8=mymenu4.Append(-1,'TUV')
        text9=mymenu5.Append(-1,'iEC60335-2-40:2018')
        self.SetMenuBar(mymenubar)
        mymenubar.Append(mymenu1,'File')
        mymenubar.Append(mymenu2,'CB')
        mymenubar.Append(mymenu3,'ETL')
        mymenubar.Append(mymenu4,'Search Online')
        mymenubar.Append(mymenu5,'TDS')
        self.Bind(wx.EVT_MENU,self.exit,Exit)
        self.Bind(wx.EVT_MENU,self.frame1,text2)
        self.Bind(wx.EVT_MENU,self.frame2,text3)
        self.Bind(wx.EVT_MENU,self.frame3,text4)
        

    def exit(self,event): #Event for Exit
        self.Close(True)

    def frame1(self,event): #Event for iEC60335-1:2010+A1:2013+A2:2016
        frame=myframe1()
        frame.Show()

    def frame2(self,event): #Event for iEC60335-2:2018
        frame=myframe2()
        frame.Show()

    def frame3(self,event): #Event for UL484
        frame=myframe3()
        frame.Show()

    
class myframe1(wx.Frame):   #Frame for iEC60335-1:2010+A1:2013+A2:2016

    def __init__(self):
        wx.Frame.__init__(self,None,-1,'iEC60335-1:2010+A1:2013+A2:2016 --by Tool4Cert by Michael')
        panel=wx.Panel(self)
        mymenubar=wx.MenuBar()
        mymenu1=wx.Menu()   #Exit
        Exit=mymenu1.Append(wx.NewId(),'&Exit')
        mymenubar.Append(mymenu1,'File')
        mymenu2=wx.Menu()
        mymenu3=wx.Menu()
        text1=mymenu2.Append(-1,'Clause13')
        text2=mymenu2.Append(-1,'Clause16')
        mymenubar.Append(mymenu2,'Clauses')
        self.SetMenuBar(mymenubar)
        self.Bind(wx.EVT_MENU,self.exit,Exit)
        self.Bind(wx.EVT_MENU,self.LC13,text1)
        self.Bind(wx.EVT_MENU,self.LC16,text2)

    def exit(self,event): #Event for Exit
        self.Close(True)

    def LC13(self,event): #Event for leakage current
        frame=IEC60335_1_Clause13_LeakageCurrent.LeakageCurrentFrame()
        frame.Show()

    def LC16(self,event): #Event for leakage current
        frame=IEC60335_1_Clause16_LeakageCurrent.LeakageCurrentFrame()
        frame.Show()

class myframe2(wx.Frame):   #Frame for iEC60335-2-40:2018

    def __init__(self):
        wx.Frame.__init__(self,None,-1,'iEC60335-2-40:2018 --by Tool4Cert by Michael')
        panel=wx.Panel(self)
        mymenubar=wx.MenuBar()
        mymenu1=wx.Menu()   #Exit
        Exit=mymenu1.Append(wx.NewId(),'&Exit')
        mymenubar.Append(mymenu1,'File')
#        mymenu2=wx.Menu()
#        mymenu3=wx.Menu()
#        text1=mymenu2.Append(-1,'Clause13')
#        text2=mymenu2.Append(-1,'Clause16')
#        mymenubar.Append(mymenu2,'Clauses')
        self.SetMenuBar(mymenubar)
        self.Bind(wx.EVT_MENU,self.exit,Exit)
#        self.Bind(wx.EVT_MENU,self.LC13,text1)
#        self.Bind(wx.EVT_MENU,self.LC16,text2)

    def exit(self,event): #Event for Exit
        self.Close(True)

class myframe3(wx.Frame):   #Frame for UL484

    def __init__(self):
        wx.Frame.__init__(self,None,-1,'UL484 --by Tool4Cert by Michael')
        panel=wx.Panel(self)
        mymenubar=wx.MenuBar()
        mymenu1=wx.Menu()   #Exit
        Exit=mymenu1.Append(wx.NewId(),'&Exit')
        mymenubar.Append(mymenu1,'File')
        mymenu2=wx.Menu()
        mymenu3=wx.Menu()
        text1=mymenu2.Append(-1,'Section33')
        text2=mymenu2.Append(-1,'Section38')
        text3=mymenu2.Append(-1,'Section43')
        text4=mymenu2.Append(-1,'Section46')
        mymenubar.Append(mymenu2,'Sections')
        self.SetMenuBar(mymenubar)
        self.Bind(wx.EVT_MENU,self.exit,Exit)
        self.Bind(wx.EVT_MENU,self.HVC,text1)
        self.Bind(wx.EVT_MENU,self.TubeThickness,text2)
        self.Bind(wx.EVT_MENU,self.TestVoltage,text3)
        self.Bind(wx.EVT_MENU,self.TestCondition,text4)

    def exit(self,event): #Event for Exit
        self.Close(True)

    def HVC(self,event):    #Event for HighVoltageCircuits
        frame=UL484_Section33_HighVoltageCircuits.HVCircuitsFrame()
        frame.Show()

    def TubeThickness(self,event):  #Event for TubeThickness
        frame=UL484_Section38_TubeThickness.TubeThicknessFrame()
        frame.Show()

    def TestVoltage(self,event):  #Event for TestVoltage
        frame=UL484_Clause43_TestVoltage.TestVoltage()
        frame.Show()

    def TestCondition(self,event):  #Event for TestCondition
        frame=UL484_Clause46_TestCondition.TestConditionFrame()
        frame.Show()
        
if __name__=="__main__":
    myapp=wx.PySimpleApp()
    frame=myframe()
    frame.Show()
    myapp.MainLoop()

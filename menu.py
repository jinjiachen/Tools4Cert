#!/bin/python
#coding=utf8

import wx

class myframe(wx.Frame):

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
        

    def exit(self,event): #Event for Exit
        self.Close(True)

    def frame1(self,event): #Event for iEC60335-1:2010+A1:2013+A2:2016
        frame1=myframe1()
        frame1.Show()

    
class myframe1(wx.Frame):

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

    def exit(self,event): #Event for Exit
        self.Close(True)

if __name__=="__main__":
    myapp=wx.PySimpleApp()
    frame=myframe()
    frame.Show()
    myapp.MainLoop()

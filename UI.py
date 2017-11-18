#!/usr/bin/env python 

import wx

class MainFrame(wx.Frame):
	def __init__(self,parent,title):
		wx.Frame.__init__(self,parent,title=title,size=(300,200))
		#self.control=wx.TextCtrl(self,style=wx.TE_MULTILINE)
		
		self.CreateStatusBar()
		filemenu=wx.Menu()
		
		
		menuItem = filemenu.Append(wx.ID_ABOUT,"&About","Designed By Sam")
		self.Bind(wx.EVT_MENU,self.OnAbout,menuItem)
		filemenu.AppendSeparator()
		filemenu.Append(wx.ID_EXIT,"&Exit","Exit")
		
		menuBar=wx.MenuBar()
		menuBar.Append(filemenu,"&File")
		self.SetMenuBar(menuBar)
		
		
		self.Show(True)
		
	def OnAbout(self,event):
		print("About event!")
		
		

app = wx.App(False)
frame=MainFrame(None,'Hey')
app.MainLoop()
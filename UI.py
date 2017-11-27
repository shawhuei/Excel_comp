#!/usr/bin/env python 

import wx

class MainFrame(wx.Frame):
	ContentsA = ''
	ContentsB = ''
	def __init__(self,parent,title):
		wx.Frame.__init__(self,parent,title=title,size=(1000,600))
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

		panel = wx.Panel(self,-1)
		#set button A
		self.button = wx.Button(panel,-1,"OpenFile A",pos=(50,20))
		self.Bind(wx.EVT_BUTTON,self.OnOpenAClick,self.button)
		self.button.SetDefault()
		self.ContentsA = wx.TextCtrl(panel,pos=(50,50),size=(360,40),style=wx.HSCROLL)

		#set button B
		self.button = wx.Button(panel,-1,"OpenFile B",pos=(50,120))
		self.Bind(wx.EVT_BUTTON,self.OnOpenBClick,self.button)
		self.button.SetDefault()
		self.ContentsB = wx.TextCtrl(panel,pos=(50,150),size=(360,40),style=wx.HSCROLL)

		#set button B
		self.button = wx.Button(panel,-1,"Start Compare",pos=(50,200))
		self.Bind(wx.EVT_BUTTON,self.OnCompareClick,self.button)
		self.button.SetDefault()

		self.Show(True)
		
	def OnAbout(self,event):
		print("About event!")
		#Pop a message
		dlg=wx.MessageDialog(None,"Designed by Sam!\n(shawhuei@126.com)","About",wx.YES_DEFAULT)
		result=dlg.ShowModal()
		dlg.Destroy()
		
	def OnOpenAClick(self,event):
		print("Click A!")
		dlg = wx.FileDialog(self,message="Choose a file",defaultFile="",wildcard="Excel files (*.xlsx)|*.xlsx")#,style=wx.CHANGE_DIR)#wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			tmp=''
			paths = dlg.GetPaths()
			for path in paths:
				tmp=tmp+path
			self.ContentsA.SetValue(tmp)
		dlg.Destroy()

	def OnOpenBClick(self,event):
		print("Click B!")
		dlg = wx.FileDialog(self,message="Choose a file",defaultFile="",wildcard="Excel files (*.xlsx)|*.xlsx")#,style=wx.CHANGE_DIR)#wx.OPEN | wx.MULTIPLE | wx.CHANGE_DIR)
		if dlg.ShowModal() == wx.ID_OK:
			tmp=''
			paths = dlg.GetPaths()
			for path in paths:
				tmp=tmp+path
			self.ContentsB.SetValue(tmp)
		dlg.Destroy()			

	def OnCompareClick(self,event):
		print("start compare!")
		fileA = self.ContentsA.GetValue()
		fileB = self.ContentsB.GetValue()
		print(fileA + fileB)
		

app = wx.App(False)
frame=MainFrame(None,'Excel Compare')
app.MainLoop()
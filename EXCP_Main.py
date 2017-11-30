#!/usr/bin/env python 

import wx

from excel import XLSX_class

VERSION_NO = "V0.1"

LOGO_DIR = "Res"

class MainFrame(wx.Frame):
	ContentsA = ''
	ContentsB = ''
	def __init__(self,parent,title):
		self.ContentsA = ''
		self.ContentsB = ''
		wx.Frame.__init__(self,parent,title=title,size=(800,400))
		#self.control=wx.TextCtrl(self,style=wx.TE_MULTILINE)
		
		self.CreateStatusBar()
		filemenu=wx.Menu()
		
		
		menuItem = filemenu.Append(wx.ID_ABOUT,"&About","Designed By Sam")
		self.Bind(wx.EVT_MENU,self.OnAbout,menuItem)
		filemenu.AppendSeparator()
		menuItem=filemenu.Append(wx.ID_EXIT,"&Exit","Exit")
		self.Bind(wx.EVT_MENU,self.OnExit,menuItem)
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

		#show logo
		image = wx.Image(self.res_path(os.path.join(LOGO_DIR,'LOGO.jpg')),wx.BITMAP_TYPE_JPEG)
		temp = image.ConvertToBitmap()
		size = temp.GetWidth(),temp.GetHeight()
		wx.StaticBitmap(parent=panel,bitmap=temp,pos=(500,50))

		self.Show(True)
	def res_path(self,relative):
		if hasattr(sys, "_MEIPASS"):
			return os.path.join(sys._MEIPASS, relative)
		return os.path.join(relative)
		
	def OnAbout(self,event):
		print("About event!")
		#Pop a message
		dlg=wx.MessageDialog(None,"Designed by Sam!\n(shawhuei@126.com)\nVersion: "+VERSION_NO,"About",wx.YES_DEFAULT)
		result=dlg.ShowModal()
		dlg.Destroy()
		
	def OnExit(self,event):
		print("Exit event!")
		#Pop a message
		wx.Exit()
		pass
				
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
		Diag=""
		fileA = self.ContentsA.GetValue()
		fileB = self.ContentsB.GetValue()
		print(fileA + fileB)
		compare=XLSX_class(fileA,fileB)
		ret=compare.fill_sheets()
		print("compare ret:%d" %ret)
		if(ret==1):
			Diag="Compare completed!"
		if(ret==0):
			Diag="Same files!"
		if(ret==-1):
			Diag="Compare Failed!"
		dlg=wx.MessageDialog(None,Diag,"Result!",wx.YES_DEFAULT)
		result=dlg.ShowModal()
		dlg.Destroy()
		
	




		
if __name__ == "__main__":
	app = wx.App(False)
	frame=MainFrame(None,'Excel Compare '+VERSION_NO)
	app.MainLoop()
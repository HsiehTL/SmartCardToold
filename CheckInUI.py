# -*- coding: utf-8 -*-

###########################################################################
## Python code generated with wxFormBuilder (version Jun  5 2014)
## http://www.wxformbuilder.org/
##
## PLEASE DO "NOT" EDIT THIS FILE!
###########################################################################

import wx
import wx.xrc
import time
import os
import openpyxl
import sys
from openpyxl import load_workbook



###########################################################################
## Class MyFrame1
###########################################################################

class MyFrame1 ( wx.Frame ):
	m_button5=""
	m_button6=""
	m_listBox3=""

	def __init__( self, parent ):
		wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 400,250 ), style = wx.DEFAULT_FRAME_STYLE & ~wx.MAXIMIZE_BOX ^ wx.RESIZE_BORDER )

		self.SetSizeHintsSz( wx.DefaultSize, wx.DefaultSize )

		bSizer1 = wx.BoxSizer( wx.VERTICAL )

		m_listBox3Choices = []
		self.m_listBox3 = wx.ListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.Size( 400,150 ), m_listBox3Choices, 0 )
		bSizer1.Add( self.m_listBox3, 0, wx.ALL, 5 )

		bSizer2 = wx.BoxSizer( wx.HORIZONTAL )

		self.m_button5 = wx.Button( self, wx.ID_ANY, u"CheckIn", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer2.Add( self.m_button5, 0, wx.ALL, 5 )

		self.m_button6 = wx.Button( self, wx.ID_ANY, u"AddSvn", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer2.Add( self.m_button6, 0, wx.ALL, 5 )

		self.m_button7 = wx.Button( self, wx.ID_ANY, u"Inf.", wx.DefaultPosition, wx.DefaultSize, 0 )
		bSizer2.Add( self.m_button7, 0, wx.ALL, 5 )


		bSizer1.Add( bSizer2, 1, wx.EXPAND, 5 )


		self.SetSizer( bSizer1 )
		self.Layout()

		self.Centre( wx.BOTH )

	def __del__( self ):
		pass







wildcard = "ScoreCard file (*.xls)|"


class CheckinFrame(MyFrame1):
	SVNPathFileTxt = ".\SVNPathsFile.txt_"
	SVNFilePathsFile = '.\ScoreCard.txt'
	def __init__(self, parent):
		MyFrame1.__init__(self, parent)
		self.OnInit()

	def OnInit(self):
		self.SetTitle("Auto Checkin ScoreCard Tool")
		#self.Show()
		ListBox3 = self.m_listBox3
		ListBox3.Clear()
		#ListBox3.Append("2")
		self.OpenDefaultFile(self.SVNPathFileTxt)
		SvnPaths = self.SVNFilePathsFile.readlines()
		ScoreCardCommitLogPath=self.SVNFilePathsFile
		Times = time.time()
		CheckDateStr = 'Check Date:' + time.strftime('%Y/%m/%d',time.localtime(Times))
		#==Button bind=====
		self.Bind(wx.EVT_BUTTON, self.Checkin, self.m_button5)
		self.Bind(wx.EVT_BUTTON, self.OpenSvnXlsPath, self.m_button6)
		self.Bind(wx.EVT_BUTTON, self.Info, self.m_button7)
		#===List SVN path====
		for PathList in SvnPaths:
			ListBox3.Append(PathList)

	def OpenDefaultFile(self, SVNFilaPathsFile):
		try:
			#.SVNFilePathsFile = ""
			self.SVNFilePathsFile = open(self.SVNPathFileTxt,'r')
		except:
			return True
		else:
			self.SVNFilePathsFile = open(self.SVNPathFileTxt,'a+')
			return False

	def OpenSvnXlsPath(self,evt):
		print os.getcwd()
		#print "Open"
		dlg = wx.FileDialog(self, "Open sketch file...", os.getcwd(),
                           style=wx.OPEN, wildcard=wildcard)
		if dlg.ShowModal() == wx.ID_OK:
			filename = dlg.GetPaths()
			#print str(filename)
			if self.CheckFileExist(filename):
				self.SVNFilePathsFile = open(self.SVNPathFileTxt,'a+')
				self.SVNFilePathsFile.writelines(filename)
				self.SVNFilePathsFile.writelines("\n")
				self.SVNFilePathsFile.close()
				self.OnInit()
			else:
				Wdlg = wx.MessageDialog(self, "Can't parse this xlxs file.", 'Warning', wx.OK | wx.ICON_WARNING)
				Wdlg.ShowModal()
				Wdlg.Destroy()
			dlg.Destroy()
			#__init__(None)

	def CheckFileExist(self,filepath):
		filepath_str = "".join(filepath)
		#print type(filepath)
		print filepath_str
		try:
			#print os.path.abspath(str(filepath))
			wb2 = openpyxl.load_workbook(os.path.abspath(filepath_str))
		except Exception as inst:
			return False
		else:
			print "Can boot"
			return True

	def Info(self,evt):
		Wdlg = wx.MessageDialog(self, "Smart Tool for Score(Stupid) Card.", 'Warning', wx.OK | wx.ICON_WARNING)
		Wdlg.ShowModal()
		Wdlg.Destroy()

	def Checkin(self,evt):
		print "check in"




if __name__ == '__main__':
	app = wx.App(False)
	FrameApp = CheckinFrame(None)
	FrameApp.Show()


	app.MainLoop()
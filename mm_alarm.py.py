#!/usr/bin/env python
import datetime
import sys
import threading
import wx

from wx.lib.masked import TimeCtrl

try:
	import win32com.client
except:
	print "Install PyWin32"

class MainWindow(wx.Frame):
	def __init__(self, parent, title, sdb):

		self.player = sdb.player

		wx.Frame.__init__(self, parent, title=title, size=(200,100))
		self.timeControl = TimeCtrl(self, -1, 
								value = '00:00:00',
								pos = wx.DefaultPosition,
								size = wx.DefaultSize,
								style = wx.TE_PROCESS_TAB,
								validator = wx.DefaultValidator,
								name = "time",
								format = 'HHMMSS',
								fmt24hr = True,
								displaySeconds = False,
								spinButton = None,
								min = None,
								max = None,
								limited = None,
								oob_color = "Yellow")

		self.sizer = wx.GridSizer(rows=3, cols=1, hgap=0, vgap=0)
		self.sizer.Add(self.timeControl, 0, wx.ALIGN_CENTER)

		# Button to set alarm
		self.setButton = wx.Button(self, -1, "Set Alarm")
		self.setButton.Bind(wx.EVT_BUTTON, self.OnAlarm)

		# Button to reset alarm
		self.resetButton = wx.ToggleButton(self, wx.ID_ANY, "ON")
		self.resetButton.SetBackgroundColour('GREEN')
		self.resetButton.SetValue(True)
		self.resetButton.Bind(wx.EVT_TOGGLEBUTTON, self.OnReset)

		self.sizer.Add(self.setButton, 0, wx.EXPAND)
		self.sizer.Add(self.resetButton, 0, wx.EXPAND)

		self.SetBackgroundColour('WHITE')

		self.CreateStatusBar()

		# File Menu Set Up
		filemenu = wx.Menu()

		menuAbout = filemenu.Append(wx.ID_ABOUT, "&About", 
			"An alarm clock that ties into Media Monkey and triggers\
			 a play on the now playing playlist.")
		self.Bind(wx.EVT_MENU, self.OnAbout, menuAbout)

		menuExit = filemenu.Append(wx.ID_EXIT, "&Exit", "Terminate the Program")
		self.Bind(wx.EVT_MENU, self.OnExit, menuExit)

		# Create the Menu Bar.
		menuBar = wx.MenuBar()
		menuBar.Append(filemenu, "&File")
		self.SetMenuBar(menuBar)

		#Layout sizers
		self.SetSizer(self.sizer)
		self.SetAutoLayout(1)
		self.sizer.Fit(self)

		# Set the window to be visible.
		self.Show(True);


	def OnAbout(self, e):
		# A message dialog with OK button.
		dlg = wx.MessageDialog( self, "An alarm clock that ties into Media Monkey and triggers\
				 a play on the now playing playlist. \n\n Created by Natasha Dalal (2014).", "About", wx.OK)
		dlg.ShowModal()
		dlg.Destroy() # Destroy the window when it is finished.
		print('resetting')

	def OnExit(self, e):
		self.Close(True)

	def OnReset(self, e):
		if not self.resetButton.GetValue():
			self.resetButton.SetBackgroundColour('RED')
			self.resetButton.SetLabel("OFF")
			try:
				self.timer.Stop()
			except:
				print "No timer currently set"
		else:
			self.resetButton.SetBackgroundColour('GREEN')
			self.resetButton.SetLabel("ON")

	def OnAlarm(self, e):
		if self.resetButton.GetValue():

			try:
				self.timer.Stop()
			except:
				print "No timer currently set"

			# Get the current set alarm time:
			atTime = self.timeControl.GetValue(as_wxDateTime=True)

			# Set the date to the right date.
			atTime.SetDay(wx.DateTime.Today().GetDay())
			atTime.SetMonth(wx.DateTime.Today().GetMonth())
			atTime.SetYear(wx.DateTime.Today().GetYear())

			# If the time specified has already passed, set it for
			# the provided time on the next day.
			now =  wx.DateTime.Now()
			if atTime.IsEarlierThan(now):
				# Number of Days in the current month
				dim = wx.DateTime.GetNumberOfDaysInMonth(atTime.GetMonth())
				if atTime.GetDay() == dim:
					atTime.SetDay(1)
					atTime.SetMonth(atTime.GetMonth() + 1)
				else:
					atTime.SetDay(atTime.GetDay() + 1)

			#At this point, At time should be the right time and date.

			# The sleep time in seconds
			sleepTime = atTime.GetTicks() - now.GetTicks()
			self.timer = wx.Timer(self, -1)
	        self.Bind(wx.EVT_TIMER, self.DoPlay)

	        self.timer.Start(sleepTime * 1000, wx.TIMER_ONE_SHOT)


	def DoPlay(x, y):
		print("Starting to Play")
		sdbPlayer.Play()


if __name__ == "__main__":

	# Make a connection to the player.
	try:
		SDB = win32com.client.Dispatch("SongsDB.SDBApplication")
		SDB.ShutdownAfterDisconnect = False
		sdbPlayer = SDB.player

		# Set up and open window.
		app = wx.App(False)
		frame = MainWindow(None, "Media Monkey Alarm", SDB)
		app.MainLoop()

	except:
		#TODO: create a window for when this fails.
		print "Is MediaMonkey up and running?"
		print sys.exc_info()
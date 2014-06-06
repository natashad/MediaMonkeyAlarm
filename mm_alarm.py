#!/usr/bin/env python
import datetime
import pythoncom
import sys
import threading
import time
import wx

from wx.lib.masked import TimeCtrl

try:
    import win32com.client
except:
    print "Install PyWin32"

NOW_PLAYING = "Now Playing: "

boolReps = ['F', 'T']   # hacky!
quit = False

class MainWindow(wx.Frame):
    def __init__(self, parent, title, sdb):

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

        self.sizer = wx.GridSizer(rows=4, cols=1, hgap=0, vgap=0)

        self.timeControlSizer = wx.GridSizer(rows=1, cols=1, hgap=0, vgap=0)
        self.timeControlSizer.Add(self.timeControl, 0, wx.EXPAND)
        self.sizer.Add(self.timeControlSizer, 0, wx.ALIGN_CENTER)

        # Button to set alarm
        self.setButton = wx.Button(self, -1, "Set Alarm")
        self.setButton.Bind(wx.EVT_BUTTON, self.OnAlarm)

        # Button to reset alarm
        self.resetButton = wx.ToggleButton(self, wx.ID_ANY, "OFF")
        self.resetButton.SetBackgroundColour('RED')
        self.resetButton.SetValue(False)
        self.resetButton.Bind(wx.EVT_TOGGLEBUTTON, self.OnReset)

        self.alarmControlSizer = wx.GridSizer(rows=1, cols=2, hgap=0, vgap=0)
        self.alarmControlSizer.Add(self.setButton, 0, wx.EXPAND)
        self.alarmControlSizer.Add(self.resetButton, 0, wx.EXPAND)
        self.sizer.Add(self.alarmControlSizer, 0, wx.EXPAND)

        # Now Playing
        self.nowPlaying = wx.TextCtrl(self, wx.ID_ANY, NOW_PLAYING)
        self.nowPlaying.SetEditable(False)
        self.sizer.Add(self.nowPlaying, 0, wx.EXPAND)
        self.SetNowPlaying(sdbPlayer.CurrentSong.ArtistName,
                                sdbPlayer.CurrentSong.Title)

        # Player Controls
        self.controlsSizer = wx.GridSizer(rows=1, cols=4, hgap=0, vgap=0)

        self.playButton = wx.Button(self, wx.ID_ANY, "> ||")
        self.playButton.Bind(wx.EVT_BUTTON, self.OnPressPlay)

        self.stopButton = wx.Button(self, wx.ID_ANY, "x")
        self.stopButton.Bind(wx.EVT_BUTTON, self.OnPressStop)

        self.nextButton = wx.Button(self, wx.ID_ANY, ">>")
        self.nextButton.Bind(wx.EVT_BUTTON, self.OnPressNext)

        self.prevButton = wx.Button(self, wx.ID_ANY, "<<")
        self.prevButton.Bind(wx.EVT_BUTTON, self.OnPressPrev)

        self.controlsSizer.Add(self.playButton, 0, wx.EXPAND)
        self.controlsSizer.Add(self.stopButton, 0, wx.EXPAND)
        self.controlsSizer.Add(self.prevButton, 0, wx.EXPAND)
        self.controlsSizer.Add(self.nextButton, 0, wx.EXPAND)

        self.sizer.Add(self.controlsSizer, 0, wx.EXPAND)


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
        self.ToggleReset()

    def ToggleReset(self):
        self.SetAlarmArmed(self.resetButton.GetValue())

    def SetAlarmArmed(self, arm):
        self.resetButton.SetValue(arm)
        if not arm:
            self.resetButton.SetBackgroundColour('RED')
            self.resetButton.SetLabel("OFF")
            try:
                self.timer.Stop()
                print "Disarming Alarm"
            except:
                print "No timer currently set"
        else:
            print "Arming Alarm"
            self.resetButton.SetBackgroundColour('GREEN')
            self.resetButton.SetLabel("ON")

    def OnAlarm(self, e):
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

        # At this point, atTime should be the right time and date.

        # The sleep time in seconds
        sleepTime = atTime.GetTicks() - now.GetTicks()
        self.timer = wx.Timer(self, -1)
        self.Bind(wx.EVT_TIMER, self.DoPlay)

        self.timer.Start(sleepTime * 1000, wx.TIMER_ONE_SHOT)
        self.SetAlarmArmed(True)


    def DoPlay(self, e):
        print("Starting to Play")
        sdbPlayer.Play()
        self.ToggleReset()

    def SetNowPlaying(self, artist, title):
        self.nowPlaying.SetValue(NOW_PLAYING + artist + " - " + title)

    # Player Control Handlers
    def OnPressPlay(self, e):
        if sdbPlayer.isPlaying:
            sdbPlayer.Pause()
        else:
            sdbPlayer.Play()

    def OnPressStop(self, e):
        sdbPlayer.Stop()

    def OnPressNext(self, e):
        sdbPlayer.Next()

    def OnPressPrev(self, e):
        sdbPlayer.Previous()



class MMEventHandlers():
    def __init__(self):
        self._play_events = 0
 
    def showMM(self):
        # note: MMEventHandlers instance includes all of SDBApplication members as well
        playing = self.Player.isPlaying
        paused = self.Player.isPaused
        isong = self.Player.CurrentSongIndex
        print 'Play', boolReps[playing], '; Pause', boolReps[paused], '; iSong', isong,
        if playing:
            print '>>', self.Player.CurrentSong.ArtistName[:40]
        else:
            print
        frame.SetNowPlaying(self.Player.CurrentSong.ArtistName,
                                self.Player.CurrentSong.Title)
 
    def OnShutdown(self):   #OK
        global quit
        print '>>> SHUTDOWN >>> buh-bye' 
        quit = True
    def OnPlay(self):       #OK
        self._play_events += 1
        print "PLAY #",
        self.showMM()
    def OnPause(self):      #OK
        print "PAUS #",
        self.showMM()
 
    def OnStop(self):
        print "STOP #",
        self.showMM()
    def OnTrackEnd(self):
        print "TRKE #",
        self.showMM()
    def OnPlaybackEnd(self):
        print "PLYE #",
        self.showMM()
    def OnCompletePlaybackEnd(self):
        print "LSTE #",
        self.showMM()
    def OnSeek(self):       #OK
        print "SEEK #",
        self.showMM()
    def OnNowPlayingModified(self):     #OK
        print "LIST #",
        self.showMM()
 
    # OnTrackSkipped gets an argument
    def OnTrackSkipped(self, track):  #OK (only when playing)
        print "SKIP #",
        self.showMM()
        # the type of any argument to an event is PyIDispatch
        # here, use PyIDispatch.Invoke() to query the 'Title' attribute for printing
        print '[', track.Invoke(3,0,2,True), ']'


if __name__ == "__main__":
    # Make a connection to the player.
    try:
        
        SDB = win32com.client.DispatchWithEvents('SongsDB.SDBApplication', MMEventHandlers)
        # SDB = win32com.client.Dispatch('SongsDB.SDBApplication')
        SDB.ShutdownAfterDisconnect = False
        sdbPlayer = SDB.Player

        # Set up and open window.
        app = wx.App(False)
        frame = MainWindow(None, "Media Monkey Alarm", SDB)
        app.MainLoop()
        while not quit:
            # required by this script because no other message loop running
            # if the app has its message loop (i.e., has a Windows UI), then
            # the events will arrive with no additional handling
            pythoncom.PumpWaitingMessages()
            time.sleep(0.2)

    except:
        #TODO: create a window for when this fails.
        print "Is MediaMonkey up and running?"
        print sys.exc_info()

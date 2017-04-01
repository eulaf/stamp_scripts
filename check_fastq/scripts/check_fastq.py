#!/usr/bin/env python

"""
Test FASTQ files for corruption. Report number of lines in each file
for comparison.
"""

import gzip
import os
import struct
import sys
import time
from collections import defaultdict
from argparse import ArgumentParser
import wx
import wx.richtext 

VERSION="1.0"
BUILD="170330"

#----common.py----------------------------------------------------------------

def getScriptPath():
  return os.path.dirname(os.path.realpath(sys.argv[0]))

def format_time_minsec(seconds):
  minutes = seconds // 60
  seconds %= 60
  return "%i:%02i" % (minutes, seconds)

#----fileops.py---------------------------------------------------------------

def get_file_size(filename):
  """Only works for gzipped files < 4Gb"""
  file_size = None
#  file_size = os.path.getsize(filename)
#  print "nl file size {}".format(file_size)
  if filename.endswith('.gz'):
    with open(filename, 'rb') as fo:
      fo.seek(-4, 2)
      r = fo.read(4)
    file_size = struct.unpack('<I', r)[0]
#    print "gz file size {}".format(file_size)
  else:
    file_size = os.path.getsize(filename)
#    print "normal file size {}".format(file_size)
  sys.stdout.flush()
  return file_size

def check_fastq(fqfile, progress=None):
  sys.stdout.write("\nFile: {}\n".format(fqfile))
  sys.stdout.flush()
  err = None
  openfunc = open
  if fqfile.endswith('.gz') or fqfile.endswith('.bgz'):
    openfunc = gzip.open
  numlines = 0
  starttime = time.time()
  start_pos=None
  # not using perc_done bc it is not accurate for large fastq files
  #perc_done = 0
  try:
    filesize = get_file_size(fqfile)
    sys.stdout.write("  file size: {}\n".format(filesize))
    with openfunc(fqfile, 'r') as fh:
      for line in fh:
        numlines += 1
        if progress and numlines % 750000==0: # report progress to GUI
          if progress.quit_flag:
            break
          #perc_done = fh.tell()*100.0/filesize if filesize else 0
          if start_pos:
            progress.rtc.Delete(wx.richtext.RichTextRange(start_pos,
            progress.rtc.current_pos))
            progress.rtc.current_pos = start_pos
          else:
            start_pos = progress.rtc.current_pos
          timeelapse = format_time_minsec(time.time()-starttime)
          progress.rtc.WriteFormattedText(
            normaltext="  {} sequences ({} time elapsed)".format(
            numlines/4, timeelapse))
          wx.Yield()
      #perc_done = fh.tell()*100.0/filesize if filesize else 0
  except Exception as e:
    err = "ERROR: {} {}".format(type(e).__name__, e)
#  endtime = time.time()
#  timeelapse = format_time_minsec(endtime-starttime)
  timeelapse = format_time_minsec(time.time()-starttime)
  sys.stdout.write("  {} lines\n".format(numlines))
  sys.stdout.write("  {} sequences\n".format(numlines/4))
  sys.stdout.write("  {} time elapsed (min:sec)\n".format(timeelapse))
  if err:
    sys.stdout.write("  {}\n".format(err))
  sys.stdout.flush()
  if progress and start_pos:
    progress.rtc.Delete(wx.richtext.RichTextRange(start_pos,
    progress.rtc.current_pos))
    progress.rtc.current_pos = start_pos
    progress.rtc.WriteFormattedText(
      normaltext="  {} sequences ({} time elapsed)".format(numlines/4, timeelapse))
  return numlines, err, timeelapse


#----gui.py-------------------------------------------------------------------

class TestFastQ_App(wx.App):
  def __init__(self, args, **kwargs):
    self.args = args
    wx.App.__init__(self, kwargs)

  def OnInit(self):
    self.frame = MainFrame(self.args)
    self.frame.Show()
    self.SetTopWindow(self.frame)
    return True

class MainRTC(wx.richtext.RichTextCtrl):
  def __init__(self, parent):
    wx.richtext.RichTextCtrl.__init__(self, parent, -1, "",
            style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
    self.Bind(wx.EVT_MOUSE_EVENTS, self.DoNothing)
    self.current_pos = 0
    self.previous_pos = 0

  def DoNothing(self, event):
    pass

  def AddIntroBlurb(self):
    intro_blurb = "Check gzipped FASTQ files for corruption and get a"
    intro_blurb += " count of the number of sequences in each file."
    intro_items = [
     " Drop 'Undetermined' FASTQ files here,"+\
     " then click 'Check FASTQ files' button.",
     " Each file takes 4-8 minutes to check.",
    ]
    self.BeginFontSize(10)
    self.Newline()
    self.Newline()
    self.WriteText(intro_blurb)
    self.Newline()
    self.BeginSymbolBullet('*', 25, 30)
    for descr in intro_items:
      self.WriteText(descr)
      self.Newline()
    self.EndSymbolBullet()
    self.EndFontSize()
    self.Newline()
    self.Newline()

  def ScrollWindow(self):
    pos = self.GetScrollRange(wx.VERTICAL)
    self.Scroll(0, pos)

  def WriteFormattedText(self, boldtext='', normaltext='', bullet=False,
               red='', newline=True, pos=None):
    self.previous_pos = self.current_pos
    if pos:
      self.SetCaretPosition(pos)
    elif self.current_pos:
#      self.MoveEnd()
      self.SetCaretPosition(self.current_pos)
    if bullet:
      self.BeginSymbolBullet('*', 25, 30)
    if boldtext:
      self.BeginBold()
      self.WriteText(boldtext)
      self.EndBold()
    if normaltext:
      self.WriteText(normaltext)
    if red:
      self.BeginTextColour((255, 0, 0))
      self.WriteText(red)
      self.EndTextColour()
    if newline: 
      self.Newline()
    if bullet:
      self.EndSymbolBullet()
    self.ScrollWindow()
    self.current_pos = self.GetCaretPosition()

class MainFrame(wx.Frame):
  def __init__(self, args):
    self.args = args
    self.quit_flag = False
    wx.Frame.__init__(self, None, size=(550,500),
              title="Check FASTQ v"+VERSION, )

    panel = wx.Panel(self)
    label = wx.StaticText(panel, -1, "Drop FASTQ files here:")
#    self.rtc = wx.richtext.RichTextCtrl(panel,-1, "",
#            style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
    self.rtc = MainRTC(panel)
    self.rtc.AddIntroBlurb()
    self.button_check = wx.Button(panel, -1, "Check FASTQ files", style=wx.BU_EXACTFIT)
    self.Bind(wx.EVT_BUTTON, self.CheckFastQ, self.button_check)
    self.button_check.Enable(False)
    self.button_stop = wx.Button(panel, -1, "Stop", style=wx.BU_EXACTFIT)
    self.button_stop.SetToolTip(wx.ToolTip("Stop counting and clear file list"))
    self.button_stop.Enable(False)
    self.Bind(wx.EVT_BUTTON, self.StopCounting, self.button_stop)
    button_quit = wx.Button(panel, -1, "Quit", style=wx.BU_EXACTFIT)
    button_quit.SetToolTip(wx.ToolTip("Quit application"))
    self.Bind(wx.EVT_BUTTON, self.OnCloseMe, button_quit)
    self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

    sizer = wx.BoxSizer(wx.VERTICAL)
    sizer.Add(label, 0, wx.ALL, 5)
    sizer.Add(self.rtc, 1, wx.EXPAND|wx.ALL, 5)

    button_sizer = wx.BoxSizer(wx.HORIZONTAL)
    button_sizer.Add(self.button_check, 0, wx.ALIGN_CENTER_VERTICAL)
    button_sizer.Add(self.button_stop, 0, wx.ALIGN_CENTER_VERTICAL)
    button_sizer.AddStretchSpacer()
    button_sizer.Add(button_quit, 0, wx.ALIGN_CENTER_VERTICAL)
    sizer.Add(button_sizer, 0, wx.ALL|wx.EXPAND, 5)
    panel.SetSizer(sizer)

    self.filedrop = FileDropProcessing(self, self.rtc, self.args)
    self.rtc.SetDropTarget(self.filedrop)

  def CheckFastQ(self, event):
    self.button_check.Enable(False)
    self.button_stop.Enable(True)
    try:
      filenames = self.filedrop.dropped_files
      self.rtc.WriteFormattedText('', '{} files to process'.format(self.filedrop.num_files))
      for fqfile, num in sorted(filenames.items(), key=lambda k: (k[1], k[0])):
        if self.quit_flag:
          break
        self.rtc.WriteFormattedText("File {}: ".format(num), os.path.basename(fqfile))
        wx.Yield()
        numlines, errmsg, timeelapse = check_fastq(fqfile, progress=self)
        if errmsg:
          self.rtc.WriteFormattedText(red=errmsg)
      self.rtc.WriteFormattedText(normaltext='Done.')
      self.filedrop.Reset()
      self.button_check.Enable(False)
      self.button_stop.Enable(False)
    except Exception as e:
      err = "ERROR: {} {}".format(type(e).__name__, e)
      self.rtc.WriteFormattedText(red=err)
      self.button_check.Enable(True)
      self.button_stop.Enable(False)
    self.rtc.WriteFormattedText(newline=True)


  def StopCounting(self, event):
    self.quit_flag = True
    self.filedrop.Reset()
    self.button_check.Enable(False)
    self.button_stop.Enable(False)
    self.rtc.WriteFormattedText('','Stopped.')

  def OnCloseMe(self, event):
    self.quit_flag = True
    self.Close(True)

  def OnCloseWindow(self, event):
    self.Destroy()
    
class FileDropProcessing(wx.FileDropTarget):
  def __init__(self, main, window, args):
    wx.FileDropTarget.__init__(self)
    self.main = main
    self.window = window
    self.args = args
    self.dropped_files = {}
    self.num_files = 0

  def OnDropFiles(self, x, y, filenames):
    for fqfile in filenames:
      if fqfile not in self.dropped_files:
        self.num_files += 1
        self.dropped_files[fqfile] = self.num_files
      self.window.WriteFormattedText("File {}: ".format(self.dropped_files[fqfile]), 
                              os.path.basename(fqfile))
    self.window.WriteFormattedText("", "{} file{} dropped\n".format(
                            self.num_files, '' if self.num_files==1 else 's'))
    if self.num_files>0:
      self.main.quit_flag = False
      self.main.button_check.Enable(True)

  def Reset(self):
    self.num_files = 0
    self.dropped_files = {}

def run_gui(args):
  app = TestFastQ_App(args)
  app.MainLoop()

#-----------------------------------------------------------------------------
if __name__=='__main__':
  descr = "Test gzipped FASTQ files for corruption and get a count"
  descr += " of the number of lines and sequences in each file."
  parser = ArgumentParser(description=descr)
  parser.add_argument("fastq", nargs="*",
            help="FASTQ files")
  parser.add_argument("--debug", default=False, action='store_true',
            help="Write debugging messages")

  args = parser.parse_args()
  if len(args.fastq)==0:
    run_gui(args)
  else:
    for fqfile in args.fastq:
      numlines, err, timeelapse = check_fastq(fqfile)



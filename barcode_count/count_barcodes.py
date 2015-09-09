#!/usr/bin/python
"""
Count the barcodes in a MiSeq index read FastQ file

By default, MiSeq does not output the index read file. This has to be enabled 
in the MiSeq Reporter software with the option 
'<add key="CreateFastqForIndexReads" value="1" />' 
in the file 'MiSeqReporter.exe.config'.

2014-10-03 <stehr@stanford.edu>

Revision 9/9/2015 <eula@stanford.edu>
Add GUI and optionally parse sample2barcode.txt to label barcodes of interest.
Add option to count barcodes in non-STAMP runs.  In this case, all 8 bases of
barcode are kept.  SampleSheet.csv must be included as input to trigger this
mode.
"""

import sys
import os
import re
import gzip
import time
from threading import Thread
import wx
import ParseFastQ as fq
from operator import itemgetter

LABELS = {
    'i1_fq': 'I1 fastq',
    'sample2barcode': 'Sample2Barcode',
    'samplesheetCSV': 'Sample Sheet',
}

LIMIT=12

#----gui.py-------------------------------------------------------------------

class BarcodeCount_App(wx.App):
    def __init__(self, **kwargs):
        wx.App.__init__(self, kwargs)

    def OnInit(self):
        self.frame = BCFrame()
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

class BCFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, title="Barcode counter", size=(725,550))
        self.cb_thread = None
        self.timer = wx.Timer(self, wx.ID_ANY)
        panel = wx.Panel(self)

        self.button_clear = wx.Button(panel, -1, "Clear")
        clear_tooltip = "Clear input files and stop any currently "+\
                        "running analysis."
        self.button_clear.SetToolTip(wx.ToolTip(clear_tooltip))

        self.button_count = wx.Button(panel, -1, "Count barcodes")
        count_tooltip = "Count barcodes in index fastq file." 
        self.button_count.SetToolTip(wx.ToolTip(count_tooltip))

        self.button_print = wx.Button(panel, -1, "Save report")
        print_tooltip = "Save barcode counts to text file."
        self.button_print.SetToolTip(wx.ToolTip(print_tooltip))

        button_quit = wx.Button(panel, -1, "Quit", style=wx.BU_EXACTFIT)

        label_in = wx.StaticText(panel, -1, "Drop index fastq file here:")
        self.text_in = wx.TextCtrl(panel,-1, "", 
                       style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
        drop_tooltip = "Drop index fastq file here. A sample2barcode.txt" +\
                 " file can optionally be included to label barcodes of " +\
                 "interest. For non-STAMP runs, include the SampleSheet.csv."
        self.text_in.SetToolTip(wx.ToolTip(drop_tooltip))
        self.filedrop = FileDrop(self.text_in, self)
        self.text_in.SetDropTarget(self.filedrop)
        label_out = wx.StaticText(panel, -1, "Barcode counts:")
        self.text_out = wx.TextCtrl(panel,-1, "", size=(500,275),
                        style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
        self.gauge = wx.Gauge(panel, -1, 100, size=(500,12))

        self.Bind(wx.EVT_BUTTON, self.Reset, self.button_clear)
        self.Bind(wx.EVT_BUTTON, self.CountBarcodes, self.button_count)
        self.Bind(wx.EVT_BUTTON, self.SaveReport, self.button_print)
        self.Bind(wx.EVT_BUTTON, self.OnCloseMe, button_quit)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

        top_sizer = wx.BoxSizer(wx.HORIZONTAL)
        top_sizer.Add(label_in, 0, wx.ALIGN_CENTER_VERTICAL)
        top_sizer.AddStretchSpacer()
        top_sizer.Add(self.button_clear, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(top_sizer, 0, wx.ALL|wx.EXPAND, 5)
        sizer.Add(self.text_in, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(label_out, 0, wx.ALL, 5)
        sizer.Add(self.text_out, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(self.gauge, 0, wx.EXPAND|wx.ALL, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.Add(self.button_count, 0, wx.ALIGN_CENTER_VERTICAL)
        button_sizer.AddStretchSpacer()
        button_sizer.Add(self.button_print, 0, wx.ALIGN_CENTER_VERTICAL)
        button_sizer.Add(button_quit, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(button_sizer, 0, wx.ALL|wx.EXPAND, 5)
        panel.SetSizer(sizer)
        self.Reset(None)


    def CountBarcodes(self, event):
        infiles = self.filedrop.infiles
        if not infiles['i1_fq']:
            self.text_out.AppendText("No index file to process.\n")
            return
        err = check_infile(infiles['i1_fq'])
        if err:
            self.text_out.AppendText(err)
        else:
            self.button_count.Enable(False)
            self.is_stamp=True
            self.barcode2sample = {}
            if infiles['sample2barcode']:
                sfile = infiles['sample2barcode']
                self.barcode2sample = parse_sample2barcode(sfile)
            elif infiles['samplesheetCSV']:
                sfile = infiles['samplesheetCSV']
                self.barcode2sample = parse_samplesheetCSV(sfile)
                self.is_stamp=False
            self.cb_thread = CountBarcodeThread(infiles['i1_fq'], self,
                        callback=self.DoneCountingBarcodes, cbargs=['i1_fq',])
            self.cb_thread.start()
            self.Bind(wx.EVT_TIMER, self.WriteResults, self.timer)
            self.timer.Start(6000)
            time.sleep(3)
            self.WriteResults()

    def DoneCountingBarcodes(self, fqtype):
        self.WriteResults()
        if self.filedrop.infiles['i1_fq']:
            self.button_count.Enable(True)
        if self.barcode_counts:
            self.button_print.Enable(True)
        self.gauge.SetValue(0)
        self.timer.Stop()

    def WriteResults(self, to_window=True, ofh=sys.stderr, limit=LIMIT):
        totfilesize = 0
        progress = 0
        numrds = self.num_reads
        file_size = self.parser.file_size if self.parser else 0
        loc = self.parser._file.tell()
        perc_done = loc*100.0/file_size if file_size else 0
        perc_format = " ({:.2f}%)".format(perc_done) if loc != file_size \
                      else ''
        msg = "Number of reads: {:,d}{}\n".format(numrds, perc_format)
        msg += self.BarcodeCountStats(limit=limit)
        if ofh:
            ofh.write(msg)
            ofh.flush()
        if to_window:
            self.text_out.SetValue(msg)
            if perc_done:
                self.gauge.SetValue(int(perc_done))

    def BarcodeCountStats(self, limit=LIMIT):
        labels = self.barcode2sample
        sorted_barcode_counts = dict_items_by_val(self.barcode_counts)
        numreads = self.num_reads
        numbarcodes = len(sorted_barcode_counts)
        num = 0
        msg = "Number of unique barcodes: {}\n\n".format(numbarcodes)
        for (key,val) in sorted_barcode_counts:
            perc = val*100.0/numreads if numreads else 0
            l = "\t"+labels[key] if key in labels else ''
            barcode = pad_with_ns(key) if self.is_stamp else key
            msg += "{}\t{:8d}\t{:.2f}%{}\n".format(barcode,val, perc, l)
            num += 1
            if limit and num == limit: break
        return msg

    def SaveReport(self, event):
        self.text_out.AppendText("\nSaving report.\n")
        if not self.barcode_counts:
            self.text_out.AppendText("No results to save.\n\n")
        else:
            saveFileDialog = wx.FileDialog(self, 
                             "Save barcode counts to file", "", "", "*.txt", 
                             wx.FD_SAVE|wx.FD_OVERWRITE_PROMPT)
            if saveFileDialog.ShowModal() == wx.ID_CANCEL:
                self.text.AppendText("\nSave cancelled.\n")
            else:
                with open(saveFileDialog.GetPath(), 'w') as ofh:
                    self.WriteResults(to_window=False, ofh=ofh, limit=None)
                self.text.AppendText("\nBarcode counts saved to {}.\n".\
                    format(saveFileDialog.GetPath()))

    def Reset(self, event):
        self.StopThreads()
        self.barcode_counts = {}
        self.num_reads = 0
        self.parser = None
        self.filedrop.ClearInput(None)
        self.button_count.Enable(False)
        self.button_print.Enable(False)

    def StopThreads(self):
        if self.cb_thread:
            self.cb_thread.stop = True
            self.cb_thread = None
        if self.timer:
            self.timer.Stop()
        self.gauge.SetValue(0)


    def OnCloseMe(self, event):
        self.Close(True)

    def OnCloseWindow(self, event):
        self.StopThreads()
        self.Destroy()
        

class FileDrop(wx.FileDropTarget):
    def __init__(self, window, parent):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.parent = parent
        self.infiles = {}

    def ClearInput(self, event):
        self.infiles = {'i1_fq': None, 'r1_fq': None,'r2_fq': None, 
                        'sample2barcode': None, 'samplesheetCSV': None,
                        'unknown': [], }
        self.UpdateInputText()

    def UpdateInputText(self):
        msg = ''
        for k in sorted(LABELS.keys()):
            if self.infiles[k]:
                msg += "{:<12s}\t {}\n".format(LABELS[k].upper()+':', 
                                           self.infiles[k])
        if self.infiles['unknown']:
            while self.infiles['unknown']:
                unknown = self.infiles['unknown'].pop(0)
                msg += "Unrecogized file: {}\n".format(unknown)
        self.window.SetValue(msg)

    def OnDropFiles(self, x, y, filenames):
        counts = {}
        unknown = []
        for dropfile in filenames:
            if '_I1_' in dropfile:
                self.infiles['i1_fq'] = dropfile
                self.parent.button_count.Enable(True)
            elif 'sample2barcode' in dropfile:
                self.infiles['sample2barcode'] = dropfile
            elif '.csv' in dropfile:
                self.infiles['samplesheetCSV'] = dropfile
            else:
                self.infiles['unknown'].append(dropfile)
        self.UpdateInputText()

class CountBarcodeThread(Thread):
    def __init__(self, i1file, results, callback=None, cbargs=[]):
        Thread.__init__(self)
        self.setDaemon(True)
        self.i1file = i1file
        self.results = results
        self.callback = callback
        self.callback_args = cbargs
        self.stop = False

    def run(self):
        i1p = fq.FastQParser(self.i1file)
        self.results.parser = i1p
        self.results.num_reads = 0
        self.results.barcode_counts = {}
        for index in i1p:
            barcode = index.seq
            bc = get_sub_barcode(barcode) if self.results.is_stamp else barcode
            self.results.barcode_counts[bc] = \
                self.results.barcode_counts.get(bc,0) + 1
            self.results.num_reads += 1
            if self.stop: break
        if self.callback:
            wx.CallAfter(self.callback, *self.callback_args)


def run_gui():
    app = BarcodeCount_App()
    app.MainLoop()

# ----------- Functions -----------

"""from an 8-base standard illumina barcode read from the index reads, return the substring that we use for demultiplexing"""
def get_sub_barcode(barcode):
  return barcode[4:8]

"""corresponding to get_sub_barcode(), fill a sub-barcode with N's to get 8 bases"""
def pad_with_ns(sub_barcode):
  return "NNNN%s" % sub_barcode

"""returns the items in a dictionary as a list of tuples sorted by descending value"""
def dict_items_by_val(d):
  return sorted(d.iteritems(), key=itemgetter(1), reverse=True)

"""print error and exit if file is not readable"""
def check_infile(f):
  msg = None
  if not f:
      msg = "No index file to process.\n"
      sys.stderr.write(msg)
  elif not (os.path.isfile(f) and os.access(f, os.R_OK)):
      msg = "ERROR: Can not read from file %s\n" %f
      sys.stderr.write(msg)
  return msg

def write_barcode_count_stats(barcode_counts, tot, limit=None, is_stamp=True,
                              ofh=sys.stdout, labels={}):
  sorted_barcode_counts = dict_items_by_val(barcode_counts)
  num = 0
  ofh.write("\nNumber of reads: {}\n".format(tot))
  for (key,val) in sorted_barcode_counts:
    perc = val*100.0/tot if tot else 0
    l = "\t"+labels[key] if key in labels else ''
    barcode = pad_with_ns(key) if is_stamp else key
    ofh.write("{}\t{:8d}\t{:.2f}%{}\n".format(barcode,val, perc, l))
    num += 1
    if limit and num == limit: break
  ofh.flush()

def parse_sample2barcode(s2b_file):
    barcode2sample = {}
    if s2b_file:
        with open(s2b_file, 'r') as fh:
            barcode2sample = dict([ list(reversed(l.rstrip().split("\t"))) 
                             for l in fh.readlines() if '\t' in l ])
    return barcode2sample

def parse_samplesheetCSV(csv_file):
    barcode2sample = {}
    if csv_file:
        fields = None
        with open(csv_file, 'r') as fh:
            for line in fh:
                if 'Sample_Name' in line and 'index' in line:
                    fields = line.split(',')
                elif fields:
                    v = line.split(',')
                    d = dict(zip(fields, v))
                    barcode2sample[d['index']] = d['Sample_Name']
    return barcode2sample

# ------------- MAIN --------------
if __name__ == '__main__':
    is_stamp = True
    if len(sys.argv)==1:
        run_gui()
    elif sys.argv[1] == '-h':
        print "STAMP usage: %s I1.fastq.gz [sample2barcode.txt]" % sys.argv[0]
        print "Other usage: %s I1.fastq.gz SampleSheet.csv" % sys.argv[0]
    else:
        i1f = sys.argv[1]
        b2s = {}
        if len(sys.argv)>2:
            if sys.argv[2].endswith('.csv'):
                b2s = parse_samplesheetCSV(sys.argv[2])
                is_stamp = False
            else:
                b2s = parse_sample2barcode(sys.argv[2])

        # check file parameters
        if check_infile(i1f):
            sys.exit(2)

        # parse input fastq file
        i1p = fq.FastQParser(i1f)
        c = 0
        barcode_counts = {}
        for index in i1p:
            barcode = index.seq
            bc = get_sub_barcode(barcode) if is_stamp else barcode
            barcode_counts[bc] = barcode_counts.get(bc,0) + 1
            c = c + 1
            if c % 500000==0:
                write_barcode_count_stats(barcode_counts, c, limit=LIMIT, 
                                          ofh=sys.stderr, labels=b2s)
        write_barcode_count_stats(barcode_counts, c, labels=b2s)


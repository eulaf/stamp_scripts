#!/usr/bin/env python

"""
Record barcode counts for water control in STAMP runs.
"""

import os
import sys
import datetime
import openpyxl
import operator
import re
import sqlite3
import time
import wx
import wx.lib.agw.flatnotebook as fnb
import xlsxwriter
from collections import defaultdict
from argparse import ArgumentParser

VERSION="1.0"
BUILD="160909"

#----common.py----------------------------------------------------------------

LIMIT=0.005 # = 0.5%

def getScriptPath(addpath=None):
    """Return directory containing script or, if addpath is given, a location 
    relative to the script directory"""
    path = os.path.dirname(os.path.realpath(sys.argv[0]))
    if addpath:
        path = os.path.abspath(os.path.join(path, *addpath))
    return path

DEFAULT_DOCS_DIR = getScriptPath(["..", "docs"])
DEFAULT_DATA_DIR = getScriptPath(["..", "data"])
BARCODES = {'NNNNGTCA':'water',}
REFS = {'SCHEMAFILE': 'stamp_barcode_schema.sql',
        'SQLITEDB': 'stamp_water_barcode_counts.db',
        'SPREADSHEET': 'stamp_water_barcode_counts.xlsx', }


#----fileops.py---------------------------------------------------------------

def parse_barcode_file(infile, debug=False):
    data = {}
    with open(infile, 'r') as fh:
        for l in fh:
            row = l.rstrip().split()
            if len(row)==2:
                data[row[0]] = int(row[1])
    numlines = len(data)
    if debug:
        sys.stderr.write("  Parsed {} lines from {}\n".format(numlines, infile))
    return data

def get_file_run(infile, i=0):
    filepath = os.path.realpath(infile)
    filepath = filepath.replace('stamp','STAMP') # for naming consistency
    run = "NoName_{}".format(i) # default name
    m = re.search('([-\w]+)-analysis', filepath)
    m2 = re.search('(ST\w+\d-\d{3,3}.*)[\b\.]', filepath)
    m3 = re.search('[\b_](ST\w{3,3}\d{2,3}.*)[\b\.]', filepath)
    if m:
        run = m.group(1)
    elif m2:
        run = m2.group(1)
    elif m3:
        run = m3.group(1)
    sys.stderr.write("{}\t{}\n".format(run, infile))
    return run

def stamp_run_sortkey(runname):
    """Sort runs by run number.
       STAMP2-225 > STAMP223"""
    match = re.match('(STAMP\d)-(\d\d\d.*)', runname)
    stampversion = runname
    runnum = runname
    if match:
        stampversion = match.group(1)
        runnum = match.group(2)
    else:
        match = re.match('ST...(\d\d\d.*)', runname)
        if match:
            stampversion = 'STAMP1'
            runnum = match.group(1)
    return (stampversion, runnum, runname) 

def analyze_barcode_data(bcdata):
    results = {}
    totalreads = sum(bcdata.values())
    for barcode in sorted(BARCODES.keys()):
        if barcode in bcdata:
            results[barcode] = {'count':bcdata[barcode],
                                'percent': bcdata[barcode]*100.0/totalreads}
    return totalreads, results

#----dbops.py-----------------------------------------------------------------

def connect_db(dbfile):
    sys.stderr.write("  Connecting to db {}\n".format(dbfile))
    dbh = sqlite3.connect(dbfile)
    return dbh

def add_schema(dbh, schemafile):
    sys.stderr.write("  Reading schema {}\n".format(schemafile))
    with open(schemafile, 'r') as fh:
        schema = ' '.join(fh.readlines())
        dbh.executescript(schema)

def current_time():
    return datetime.datetime.now()

def results_as_dict(cursor):
    """Convert each row from db to dict keyed by column name.
    Returns a list of dicts"""
    columns = [ d[0] for d in cursor.description ]
    data = []
    for ans in cursor.fetchall():
        d = dict(zip(columns, ans))
        data.append(d)
    return data

def get_run(cursor, run):
    runs = get_runs(cursor, run)
    return runs[0] if runs else None
    
def get_runs(cursor, run=None, status=None):
    cmd = "SELECT * FROM run"
    where = []
    args = []
    if run:
        where.append("run_name=?")
        args.append(run)
    if status:
        where.append("run_status=?")
        args.append(status)
    if where:
        cmd += " WHERE " + " AND ".join(where)
    cursor.execute(cmd, args)
    results = results_as_dict(cursor)
    return results
    
def save_run(cursor, run_name, status=None, total_reads=None):
    fields = ['run_name', 'last_modified']
    args = [run_name, current_time()]
    if status:
        fields.append('run_status')
        args.append(status)
    if total_reads is not None:
        fields.append('total_reads')
        args.append(total_reads)
    cursor.execute("REPLACE INTO run ({})".format(", ".join(fields))+\
        " VALUES (?,?,?,?)", args)

def update_run(cursor, run_name, total_reads=None, status=None):
    setcmds=[]
    args=[]
    if total_reads is not None:
        setcmds.append("total_reads=?")
        args.append(total_reads)
    if status is not None:
        setcmds.append("run_status=?")
        args.append(status)
    setcmds.append("last_modified=?")
    args.append(current_time())
    cmd="UPDATE sample SET {} WHERE run_name=?".format(", ".join(setcmds))
    args.append(run_name)
    sys.stderr.write("  Updating run {}\n".format(run_name))
    cursor.execute(cmd, args)

def get_barcode_id(cursor, barcode):
    fields = ['barcode',]
    cursor.execute("SELECT id FROM barcode WHERE barcode=?", [barcode,])
    ans = cursor.fetchone()
    return ans[0] if ans else None

def save_barcode(cursor, barcode):
    fields = ['barcode',]
    cursor.execute("INSERT OR IGNORE INTO barcode ({})".format(", ".join(fields))+\
        " VALUES (?)", [barcode,])

def get_barcode_counts_for_run_id(cursor, run_id):
    cmd = "SELECT b.barcode, bc.bc_count FROM barcode_counts bc, barcode b" +\
          " WHERE b.id=bc.barcode_id AND bc.run_id=? "
    args = [run_id,]
    cursor.execute(cmd, [run_id,])
    bc_data = {}
    for ans in cursor.fetchall():
        (barcode, count) = ans
        bc_data[barcode] = {'count':count}
    return bc_data
    

def save_barcode_count(cursor, run_name, barcode, count):
    barcode_id = get_barcode_id(cursor, barcode)
    if not barcode_id:
        save_barcode(cursor, barcode)
        barcode_id = get_barcode_id(cursor, barcode)
    run = get_run(cursor, run_name)
    if not run:
        save_run(cursor, run_name)
        run = get_run(cursor, run_name)
    dbtable = 'barcode_counts'
    fields = ['run_id', 'barcode_id', 'bc_count', 'last_modified']
    columns = '({})'.format(','.join(fields))
    valq = ',?'*(len(fields)-1)
    ins_sql = 'REPLACE INTO {} {} VALUES (?{})'.format(dbtable, columns, valq)
    vals = [ run['id'], barcode_id, count ]
    vals.append(current_time())
    cursor.execute(ins_sql, vals)

def check_db(datadir, docsdir):
    """Look for existing DB or create new DB.  Return handle to DB"""
    dbfile = os.path.join(datadir, REFS['SQLITEDB'])
    schemafile = os.path.join(docsdir, REFS['SCHEMAFILE'])
    is_new_db = not os.path.exists(dbfile)
    dbh = connect_db(dbfile)
    cursor = dbh.cursor()
    msgs = []
    msgs.append('Water barcode: {}'.format(', '.join(BARCODES.keys())))
    if is_new_db:
        sys.stderr.write("Creating new db: {}\n".format(dbfile))
        add_schema(dbh, schemafile)
        msgs.append("    0 runs saved")
    else:
        runs = get_runs(cursor, status='PASS')
        numruns = len(runs)
        msgs.append("    {} runs saved".format(numruns))
    sys.stderr.flush()
    return (dbh, msgs)

def get_rundata_from_db(dbh, status='PASS'):
    cursor = dbh.cursor()
    runs = get_runs(cursor, status=status)
    rundata = {}
    for d in runs:
        d['bc_counts'] = get_barcode_counts_for_run_id(cursor, d['id'])
        rundata[d['run_name']] = d
    return rundata

def save_rundata_db(dbh, rundata, status='PASS'):
    cursor = dbh.cursor()
    runs = sorted(rundata.keys(), key=stamp_run_sortkey)
    numruns_saved = 0
    for run_name in runs:
        totreads = rundata[run_name]['total_reads']
        save_run(cursor, run_name, status=status, total_reads=totreads)
        numruns_saved += 1
        for barcode, d in sorted(rundata[run_name]['bc_counts'].items()):
            save_barcode_count(cursor, run_name, barcode, d['count'])
    dbh.commit()

#----spreadsheet.py-----------------------------------------------------------

def convert_to_excel_col(colnum):
    mod = colnum % 26
    let = chr(mod+65)
    if colnum >= 26:
        rep = colnum/26
        let1 = chr(rep+64)
        let = let1 + let
    return let

def add_formats_to_workbook(workbook):
    wbformat = {}
    wbformat['bold'] = workbook.add_format({'bold': True})
    wbformat['perc'] = workbook.add_format({'num_format': '#.####%'})
    wbformat['red'] = workbook.add_format({'bg_color': '#C58886', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltred'] = workbook.add_format({'bg_color': '#E9D4D3', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['orange'] = workbook.add_format({'bg_color': '#FCD5B4',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltblue'] = workbook.add_format({'bg_color': '#D7E1EB',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['blue'] = workbook.add_format({'bg_color': '#88A4C5',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltbluepatt'] = workbook.add_format({'fg_color': '#DCE6F0', 
                                       'pattern': 8,
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltgreen_bold'] = workbook.add_format({'bg_color': '#EBF1DE',
                                       'bold': True,
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltgreen'] = workbook.add_format({'bg_color': '#EBF1DE',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltgreen_perc'] = workbook.add_format({'num_format': '#.####%',
                                       'bg_color': '#EBF1DE',
                                       'border': 1, 'border_color':'#CDCDCD'})
    return wbformat

def add_barcode_sheet_excel(workbook, wbformat, rundata):
    goodruns = [r for r in rundata.keys() if rundata[r]['run_status']=='PASS']
    runs = sorted(goodruns, reverse=True, key=stamp_run_sortkey)
    barcodes = sorted(BARCODES.keys())
    worksheet = workbook.add_worksheet('Barcode counts')
    rownum = 0
    # comment lines
    for line in ('# This spreadsheet is automatically generated.' +\
                 ' Any edits will be lost in future versions.',
                 '# Num runs in spreadsheet: {}'.format(len(rundata))):
        worksheet.write(rownum, 0, line)
        rownum += 1
    rownum += 1 # skip row
    # print fields
    fields = ['Total reads', 'Read count', 'Read %']
    numfields = len(fields)
    worksheet.write(rownum+1, 0, 'Runs', wbformat['bold'])
    for i, barcode in enumerate(barcodes):
        collet_start = convert_to_excel_col(numfields*i+1)
        collet_end = convert_to_excel_col(numfields*(i+1))
        worksheet.merge_range('{0}{2}:{1}{2}'.format(collet_start, collet_end, 
                              rownum+1), barcode, wbformat['bold'])
        for j, field in enumerate(fields):
            worksheet.write(rownum+1, i*numfields+j+1, field, wbformat['bold'])
    rownum += 2
    worksheet.freeze_panes(rownum, 0)
    # print run data
    runrowxl_s = rownum+1 # excel row is 1-based
    for run in runs:
        worksheet.write(rownum, 0, run)
        for i, barcode in enumerate(barcodes):
            total = rundata[run]['total_reads']
            count = rundata[run]['bc_counts'][barcode]['count']
            collet_total = convert_to_excel_col(i*numfields+1)
            collet_count = convert_to_excel_col(i*numfields+2)
            frac = "={1}{2}/{0}{2}".format(collet_total, collet_count, rownum+1)
            worksheet.write_number(rownum, i*numfields+1, total)
            worksheet.write_number(rownum, i*numfields+2, count)
            worksheet.write(rownum, i*numfields+3, frac, wbformat['perc'])
        rownum += 1
    runrowxl_e = rownum
    # add median and average calculation
#    i_row_med = rownum
#    i_row_avg = rownum+1
#    worksheet.write(i_row_med, 0, 'Median', wbformat['ltgreen_bold'])
#    worksheet.write(i_row_avg, 0, 'Average', wbformat['ltgreen_bold'])
    for i, barcode in enumerate(barcodes):
#        for k, calc in enumerate(('Median', 'Average')):
        for j, calcformat in enumerate(('ltgreen', 'ltgreen', 
                                            'ltgreen_perc')):
                colnum = i*numfields+j+1
                collet = convert_to_excel_col(colnum)
                runrange = "{0}{1}:{0}{2}".format(collet, runrowxl_s, runrowxl_e)
#                worksheet.write(rownum+k, colnum, '={}({})'.format(calc, runrange),
#                                wbformat[calcformat])
        worksheet.conditional_format(runrange, {'type':'cell', 
            'criteria':'>', 'value':LIMIT, 'format':wbformat['red'], })
#        worksheet.conditional_format(runrange, {'type':'cell', 
#            'criteria':'between', 'minimum': 0.001, 'maximum': LIMIT,
#               'format':wbformat['ltred'], })
    for i in [0, ]:
        worksheet.set_row(i, None, None, {'hidden': True})

def create_excel_spreadsheet(rundata, outfile):
    sys.stderr.write("\nCreating barcode Excel file:\n{}\n".format(outfile))
    workbook = xlsxwriter.Workbook(outfile)
    wbformat = add_formats_to_workbook(workbook)
    nums = add_barcode_sheet_excel(workbook, wbformat, rundata)
    workbook.close()
    wb = openpyxl.load_workbook(outfile)
    wb.save(outfile)
    return nums

#----gui.py-------------------------------------------------------------------

class StampWaterBarcode_App(wx.App):
    def __init__(self, dbh, msg=None, spreadsheet=None, **kwargs):
        self.dbh = dbh
        self.msg = msg
        self.spreadsheet = spreadsheet
        wx.App.__init__(self, kwargs)

    def OnInit(self):
        self.frame = StampFrame(self.dbh, msg=self.msg,
                                spreadsheet=self.spreadsheet)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

class StampFrame(wx.Frame):
    def __init__(self, dbh, msg=None, spreadsheet=None):
        wx.Frame.__init__(self, None, title="STAMP Water Barcode v{}".format(VERSION), 
                          size=(550,425))
        self.dbh = dbh
        self.spreadsheet = spreadsheet
        panel = wx.Panel(self)
        label = wx.StaticText(panel, -1, 
            "Drop barcode_counts.txt file(s) here:")
        self.text = wx.TextCtrl(panel,-1, "",style=wx.TE_READONLY|
                                wx.TE_MULTILINE|wx.HSCROLL)
        button_save = wx.Button(panel, -1, "Update spreadsheets and DB")
        save_tooltip = "Update spreadsheet and database with data entered.\n"
        button_save.SetToolTip(wx.ToolTip(save_tooltip))
        self.Bind(wx.EVT_BUTTON, self.UpdateSpreadsheetAndDB, button_save)
        button_quit = wx.Button(panel, -1, "Quit", style=wx.BU_EXACTFIT)
        self.Bind(wx.EVT_BUTTON, self.OnCloseMe, button_quit)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)
        self.notebook = StampNotebook(panel, msg=msg)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(label, 0, wx.ALL, 5)
        sizer.Add(self.text, 1, wx.EXPAND|wx.ALL, 5)
        sizer.Add(self.notebook, 0, wx.EXPAND|wx.ALL, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.Add(button_save, 0, wx.ALIGN_CENTER_VERTICAL)
        button_sizer.AddStretchSpacer()
        button_sizer.Add(button_quit, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(button_sizer, 0, wx.ALL|wx.EXPAND, 5)
        panel.SetSizer(sizer)

        dt = FileDrop(self.text, self.notebook)
        self.text.SetDropTarget(dt)

    def UpdateSpreadsheetAndDB(self, event):
        self.text.AppendText("\nUpdating data:\n")
        if not self.notebook.results:
            self.text.AppendText("  No data to save to db.\n")
        else:
            for i, info in enumerate(self.notebook.results):
                if not info: continue
                entries = self.notebook.entries[i]
                run = entries['run'].GetValue()
                filenum = info['num']
                if not run:
                    msg = "    {}: Not saved.".format(filenum)
                    msg += "  Need run name\n"
                    self.text.AppendText(msg)
                    continue
                statusnum = entries['run_status'].GetSelection()
                status = entries['run_status'].GetString(statusnum)
                save_rundata_db(self.dbh, {run: info}, status=status)
                if get_run(self.dbh.cursor(), run):
                    self.text.AppendText(
                        "    {}: Saved {} data to db.\n".format(filenum, run))
                else:
                    self.text.AppendText(
                        "    {}: {} not saved to db.\n".format(filenum, run))
            allrundata = get_rundata_from_db(dbh)
            numruns = len(allrundata)
            msg = "    {} runs saved".format(numruns)
            self.notebook.tabOne.ChangeMessage(msg)
            if numruns:
                self.text.AppendText("  Updating spreadsheet.\n")
                create_excel_spreadsheet(allrundata, self.spreadsheet)
                self.text.AppendText(
                    "      Spreadsheet now contains {} runs\n".format(numruns))
        self.text.AppendText("\n")

    def OnCloseMe(self, event):
        self.Close(True)

    def OnCloseWindow(self, event):
        self.Destroy()
        
class FileDrop(wx.FileDropTarget):
    def __init__(self, window, notebook):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.notebook = notebook
        self.num_runs = 0

    def OnDropFiles(self, x, y, filenames):
        oldfiles = self.notebook.ReportFiles(include_num=True)
        oldruns2files = dict([ [get_file_run(f[0],f[1]), f[0]] for f in oldfiles ])
        runs2files = dict([ [get_file_run(f, self.num_runs+1+i), f] \
                            for (i,f) in enumerate(filenames) ])
        for run in sorted(runs2files.keys(), key=stamp_run_sortkey):
            if run in oldruns2files and oldruns2files[run] == runs2files[run]:
                continue # no need to update
            else:
                self.notebook.DeletePageRun(run)
            # add new entry
            try:
                bcdata = parse_barcode_file(runs2files[run])
                total, results = analyze_barcode_data(bcdata)
                self.num_runs += 1
                run_name = get_file_run(runs2files[run], self.num_runs)
                title = "{}: {}".format(self.num_runs, run_name)
                info = {'num': self.num_runs, 'file': runs2files[run],
                    'run': run_name, 'run_status': args.status,
                    'total_reads': total, 'bc_counts': results, }
                # add msg to drop window showing file was processed
                self.window.AppendText("Barcode counts file {}:    {}\n".format(
                                   self.num_runs, info['file']))
                self.notebook.AddResultsTab(info, title=title)
            except ValueError:
                self.window.AppendText(
                    "ERROR Could not parse file: {}\n".format(runs2files[run]))


class StampNotebook(fnb.FlatNotebook):
    def __init__(self, parent, msg=None):
        fnb.FlatNotebook.__init__(self, parent, id=wx.ID_ANY, size=(500, 170),
            agwStyle=fnb.FNB_VC8|fnb.FNB_X_ON_TAB|fnb.FNB_NO_X_BUTTON|
            fnb.FNB_NAV_BUTTONS_WHEN_NEEDED)

        self.tabOne = TabPanel_Text(self, msg=msg)
        self.AddPage(self.tabOne, "DB content")
        self.results = ['',]
        self.entries = ['',]
        self.Bind(fnb.EVT_FLATNOTEBOOK_PAGE_CLOSING, self.OnTabClosing)
        self.Bind(fnb.EVT_FLATNOTEBOOK_PAGE_DROPPED, self.OnTabDrop)

    def AddResultsTab(self, info, title=None):
        if not title:
            num = info['num'] if 'num' in info else ''
            title = "Run {}".format(num)
        newTab = TabPanel_Results(self, info)
        self.AddPage(newTab, title)
        numpages = self.GetPageCount()
        self.SetSelection(numpages-1)
        self.results.append(info)

    def OnTabClosing(self, event):
        selected = self.GetSelection()
        res = self.results.pop(selected)
        ent = self.entries.pop(selected)
        txt = self.GetPageText(selected)
        sys.stderr.flush()

    def OnTabDrop(self, event):
        selected = self.GetSelection()
        oldselected = event.GetOldSelection()
        res = self.results.pop(oldselected)
        ent = self.entries.pop(oldselected)
        self.results.insert(selected, res)
        self.entries.insert(selected, ent)

    def ReportFiles(self, include_num=False):
        """Return list of files in notebook"""
        reports = []
        for info in self.results:
            if not info: continue
            if 'file' in info:
                if include_num:
                    entry = (info['file'], info['num'])
                else:
                    entry = info['file']
                reports.append(entry)
        return reports

    def DeletePageRun(self, run):
        """Delete pages where run is given run"""
        numpages = 0
        for i, info in enumerate(self.results):
            if not info: continue
            if 'run' in info and info['run']==run:
                self.SetSelection(i)
                self.DeletePage(i)
                self.SendSizeEvent()
                numpages += 1
        return numpages

class TabPanel_Text(wx.Panel):
    def __init__(self, parent, msg="\n\n\n\n"):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)
        self.textWidget = wx.StaticText(self, -1, '\n'+msg, pos=(15,10))
#        font = wx.Font(8, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL,
#                       wx.FONTWEIGHT_NORMAL)
#        self.textWidget.SetFont(font)

    def ChangeMessage(self, msg):
        self.textWidget.Destroy()
        self.textWidget = wx.StaticText(self, -1, '\n'+msg)

class TabPanel_Results(wx.Panel):
    def __init__(self, parent, info):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)

        runLabel = wx.StaticText(self, -1, "Run:")
        runname = '' if info['run'].startswith('NoName') else info['run']
        runEntry = wx.TextCtrl(self, -1, runname)
        statusLabel = wx.StaticText(self, -1, "Status:")
        statusEntry = wx.Choice(self, -1, choices=['PASS', 'FAIL'])
        statusEntry.SetSelection(0 if info['run_status']=='PASS' else 1)
        parent.entries.append({'run': runEntry, 'run_status': statusEntry})
        infostr = 'Total reads: {:9d}\n'.format(info['total_reads'])
        for barcode in info['bc_counts']:
            count = info['bc_counts'][barcode]['count']
            perc = count*100.0/info['total_reads']
            infostr += '{}: {:8d} reads ({:6.4f}%)\n'.format(
                        barcode, count, perc)
        infoText = wx.StaticText(self, -1, infostr)

        panelSizer = wx.BoxSizer(wx.VERTICAL)
        infoSizer = wx.BoxSizer(wx.HORIZONTAL)
#        infoSizer.Add(infoText, 1, wx.ALL, 8)
        infoSizer.Add(infoText, 1, wx.EXPAND|wx.ALL, 8)
        entrySizer = wx.FlexGridSizer(cols=2, hgap=5, vgap=5)
        entrySizer.AddGrowableCol(1)
        entrySizer.Add(runLabel, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        entrySizer.Add(runEntry, 0, wx.EXPAND)
        entrySizer.Add(statusLabel, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        entrySizer.Add(statusEntry, 0)
        panelSizer.Add(entrySizer, 0, wx.EXPAND|wx.ALL, 10)
        panelSizer.Add(infoSizer, 0, wx.ALIGN_LEFT)
        self.SetSizer(panelSizer)

def run_gui(dbh, msgs, spreadsheet):
    msg = '\n'.join(msgs)
    app = StampWaterBarcode_App(dbh, msg=msg, spreadsheet=spreadsheet)
    app.MainLoop()

#-----------------------------------------------------------------------------

if __name__=='__main__':
    descr = "Saves barcode counts for STAMP water barcode: {}.".format(
            ', '.join(BARCODES.keys()))
    parser = ArgumentParser(description=descr)
    parser.add_argument("bc_files", nargs="*",
                        help="Barcode counts files(s)")
    parser.add_argument("-s", "--save", default=False, action='store_true',
                        help="Save data to database.")
    parser.add_argument("-x", "--excel", default=False, action='store_true',
                        help="Print Excel spreadsheet summarizing all data.")
    parser.add_argument("-d", "--debug", default=False, action='store_true',
                        help="Print extra messages")
    parser.add_argument("--datadir", default=DEFAULT_DATA_DIR,
                        help="Directory to find/save databases and "+\
                             "spreadsheets")
    parser.add_argument("--docsdir", default=DEFAULT_DOCS_DIR,
                        help="Directory to find db schema")
    parser.add_argument("--status", default='PASS',
                        help="Status to use for all reports (default: PASS)")

    args = parser.parse_args()
    WATERBC = BARCODES.keys()[0]
    spreadsheet = os.path.join(args.datadir, REFS['SPREADSHEET'])
    dbh, msgs = check_db(args.datadir, args.docsdir)
    sys.stderr.write('\n'.join(msgs)+'\n\n')
    if len(args.bc_files)==0:
        run_gui(dbh, msgs, spreadsheet)
    else:
        outfile = {}
        rundata = {} 
        for i, infile in enumerate(sorted(args.bc_files, 
                                   key=lambda v: v.upper())):
            bcdata = parse_barcode_file(infile, args.debug)
            if not bcdata:
                sys.stderr.write("Bad file: {}. Skipping\n".format(infile))
                continue
            run = get_file_run(infile, i+1)
            if args.debug:
                sys.stderr.write("{}) Run {}\tFile {}\n".format(i+1, run, infile))
            total, results = analyze_barcode_data(bcdata)
            flag = '!!!' if results[WATERBC]['percent']>1 else ''
            rundata[run] = { 'total_reads': total, 'bc_counts': results,
                             'run_status': args.status }
#            sys.stderr.write("Run: {}\ttotal_reads: {}\tbc_counts: {} {}\n".\
#                             format(run, total, results, flag))
        if args.save:
            save_rundata_db(dbh, rundata, status=args.status)
        if args.excel:
            allrundata = get_rundata_from_db(dbh)
            allrundata.update(rundata)
            create_excel_spreadsheet(allrundata, spreadsheet)
    dbh.close()



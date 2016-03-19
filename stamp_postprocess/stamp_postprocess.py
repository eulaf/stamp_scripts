#!/usr/bin/env python

"""
Depth report: 
Sort depth files by min depth, highlight rows < 200 and
save as Excel file.

Variant report:
Separate ACCEPTED and NOT_REPORTED by yellow hightlighted
row and save as Excel file
"""

import os
import sys
import sqlite3
import xlsxwriter
import openpyxl
import datetime
import wx
import wx.lib.agw.flatnotebook as fnb
from collections import defaultdict
from argparse import ArgumentParser

def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

#----common.py----------------------------------------------------------------

def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

DATA_DIR = os.path.abspath(os.path.join(getScriptPath(), 
                           "..", "data", "stampqc"))
REFS = {
    'TRUTHFILE': os.path.join(DATA_DIR, "truq3_truths.txt"),
    'SPREADSHEET': os.path.join(DATA_DIR, "stampQC_TruQ3.xlsx"),
    'DBFILE': os.path.join(DATA_DIR, "stampQC_TruQ3.db"),
    'SCHEMAFILE': os.path.join(DATA_DIR, "stampQC_schema.sql")
}

def is_float(v):
    try:
        float(v)
    except ValueError:
        return False
    return True

def outfile_name(report, outdir=None):
    outfile = report.replace('.txt', '') + '.xlsx'
    if outdir:
        outfile = os.path.join(outdir, os.path.basename(outfile))
    return outfile

#----fileops.py---------------------------------------------------------------

def create_dkey(d):
    """Unique key for mutations"""
    return "{}:{}:{}:{}".format(d['gene'], d['position'], d['ref'], d['var'])

def massage_data(data):
    """Unify differing formats and convert data types to appropriate types"""
    for d in data:
        d['position'] = int(d['position'])
        if not 'HGVS' in d and 'CDS_Change' in d:
            d['HGVS'] = 'c.' + d['CDS_Change']
        elif not 'CDS_Change' in d and 'HGVS' in d:
            d['CDS_Change'] = d['HGVS'].lstrip('c').lstrip('.')
        if not 'protein' in d and 'AA_Change' in d:
            d['protein'] = d['AA_Change']
        elif not 'AA_Change' in d and 'protein' in d:
            d['AA_Change'] = d['protein']
        if 'expectedVAF' in d:
            if d['expectedVAF']:
                d['expectedVAF'] = float(d['expectedVAF'].rstrip('%'))
            else:
                d['expectedVAF'] = None
        for f in ('dbSNP', 'COSMIC'):
            if f in d and d[f]=='NA':
                d[f] = None

def count_uppercase(s):
    return sum(1 for c in s if c.isupper())

def field2dbfield(f):
    """Convert field from variant report style to match db column name"""
    newf = f.replace(' ','_')
    if f.startswith('dbSNP'):
        newf = 'dbSNP'
    elif f.startswith('COSMIC'):
        newf = 'COSMIC'
    elif count_uppercase(newf)<3:
        newf = newf.lower() 
    return newf

def field2reportfield(f):
    """Convert field to variant report style"""
    newf = f.replace('_',' ')
    if count_uppercase(newf)==0:
        newf = newf.title() 
    return newf

def parse_tab_file(tabfile, keyfunc=None, fieldfunc=None):
    """Parse a tab-delimited file with column headers.  Returns a dict with
       key values:
       'fields': a list of fields in the order they appear in the column header
       'data': a list of dicts containing the values of each row in the order
               they appear in the file.  The dicts are keyed by field name with 
               values the row values; also, an additional key 'line' is added
               with value the original line from the file.  If the optional arg
               keyfunc is supplied, then and additional key 'dkey' is added
               with value the result from the keyfunc function on that row.
       'datadict': (optional) a dict keyed by the result of keyfunc operated on 
               the row dict with value the row dict.  This value only appears
               if keyfunc is present.

       Keyword arguments:
       keyfunc -- function to create a unique key for each row of data.  The
               supplied function should take a dict as its only argument and
               return a unique key for each data row.
       fieldfunc -- function to reformat field names from header."""
    data = []
    datadict = {}
    with open(tabfile, 'r') as fh:
        fields = [ fieldfunc(f) if fieldfunc else f for f in \
                   fh.readline().rstrip('\n\r').split("\t") ]
        for line in fh.readlines():
            d = dict(zip(fields, line.rstrip('\n\r').split("\t")))
            d['line'] = line
            data.append(d)
            if keyfunc:
                dkey = keyfunc(d)
                d['dkey'] = dkey
                datadict[dkey] = d
    tabfileinfo = {'fields':fields, 'data':data}
    if keyfunc:
        tabfileinfo['datadict'] = datadict
    return tabfileinfo

def parse_truth(truthfile):
    sys.stderr.write("\nReading truths: {}\n".format(truthfile))
    truthinfo = parse_tab_file(truthfile, keyfunc=create_dkey,
                               fieldfunc=field2dbfield)
    massage_data(truthinfo['data'])
    sys.stderr.write("  {} mutations\n".format(len(truthinfo['data'])))
    return truthinfo

#-----------------------------------------------------------------------------

def compare_variants(truth, vinfo, counts={}, defaultstatus='PASS'):
    truths_seen = dict([ (dkey, False) for dkey in truth.keys() ])
    summary = { 'Total':0, 'Expected':0, 'Unexpected':0, 
                'Status': defaultstatus }
    if not counts:
        counts.update({ 'Total':0, 'Expected':defaultdict(int),
               'Unexpected':defaultdict(int), })
    counts['Total'] += 1
    for d in vinfo['data']:
        d['Expected?'] = 'Not expected'
#        if d['HGVS']=='-': continue
        dkey = create_dkey(d)
        summary['Total'] += 1
        if dkey in truth:
            counts['Expected'][dkey] += 1
            summary['Expected'] += 1
            d['Expected?'] = 'Expected'
            truths_seen[dkey] = True
        else:
            counts['Unexpected'][dkey] += 1
            summary['Unexpected'] += 1
    notseen = [ dkey for dkey in truths_seen.keys() if not truths_seen[dkey] ]
    summary['Not found'] = len(notseen)
    summary['notseen'] = notseen
    if notseen:
        sys.stderr.write("    Not found: "+", ".join(notseen)+"\n")
        if len(notseen) > summary['Expected']: 
            summary['Status'] = 'FAIL'
            sys.stderr.write("Status: FAIL {} > {}\n".format(len(notseen),
                             summary['Expected']))
    else:
        sys.stderr.write("All truths found\n")
    return summary

#----dbops.py-----------------------------------------------------------------

def current_time():
    return datetime.datetime.now()

def results_as_dict(cursor):
    columns = [ d[0] for d in cursor.description ]
    data = []
    for ans in cursor.fetchall():
        d = dict(zip(columns, ans))
        data.append(d)
    return data

def connect_db(dbfile):
    sys.stderr.write("\nConnecting to db {}\n".format(dbfile))
    dbh = sqlite3.connect(dbfile)
    return dbh

def add_schema(dbh, schemafile):
    sys.stderr.write("  Reading schema {}\n".format(schemafile))
    with open(schemafile, 'r') as fh:
        schema = ' '.join(fh.readlines())
        dbh.executescript(schema)

def save_mutations(cursor, data, is_expected=0):
    fields = ['id', 'gene', 'chr', 'position', 'strand', 'ref_transcript', 
              'ref', 'var', 'dbSNP', 'COSMIC', 'HGVS', 'protein', 
              'whitelist', 'expectedVAF']
    ins_sql = 'INSERT INTO mutation VALUES (?,?' + ',?'*len(fields) + ')'
    for row in data:
        mut = [ row[f] if f in row and len(str(row[f]))>0 else None \
                for f in fields ]
        exp_val = row['is_expected'] if 'is_expected' in row else is_expected
        mut.append(exp_val)
        mut.append(current_time())
        cursor.execute(ins_sql, mut)

def save_run(cursor, run_name, sample_name, status):
    cursor.execute("INSERT INTO run (run_name, sample_name, run_status"+\
                   ", last_modified) VALUES (?,?,?,?)", 
                   (run_name, sample_name, status, current_time()))

def update_run(cursor, run_name, sample_name, status):
    sys.stderr.write("Updating run {}, {}, status={}\n".format(run_name,
                     sample_name, status))
    cursor.execute("UPDATE run SET run_status=?, last_modified=?"+\
                   " WHERE run_name=? AND sample_name=?",
                   (status, current_time(), run_name, sample_name))

def save_vaf(cursor, run_id, d, debug=False):
    mut = get_mutation(cursor, d['gene'], d['position'], d['ref'], d['var'])
    if debug: print "\nd{}\nmut {}".format(d, mut)
    if not mut: # mutation not in db, so save
        save_mutations(cursor, [d,])
        mut = get_mutation(cursor, d['gene'], d['position'], d['ref'], 
                           d['var'], debug=debug)
    ins_sql = 'INSERT INTO vaf VALUES (?,?,?,?,?)'
    vals = [ d[f] if f in d else None for f in ('VAF%', 'status') ]
    vals.append(current_time())
    cursor.execute(ins_sql, ([run_id, mut['id'],]+vals))

def get_mutation(cursor, gene, pos, ref, var, debug=False):
    muts = get_mutations(cursor, gene, pos, ref, var, debug)
    return muts[0] if muts else None
    
def get_mutations(cursor, gene=None, pos=None, ref=None, var=None,
                  debug=False):
    cmd = "SELECT * FROM mutation"
    where = []
    args = []
    if gene:
        where.append("gene=?")
        args.append(gene)
    if pos:
        where.append("position=?")
        args.append(pos)
    if ref:
        where.append("ref=?")
        args.append(ref)
    if var:
        where.append("var=?")
        args.append(var)
    if where:
        cmd += " WHERE " + " AND ".join(where)
    if debug: print "cmd {} ({})".format(cmd, args)
    cursor.execute(cmd, args)
    results = results_as_dict(cursor)
    return results

def get_run(cursor, run, sample):
    runs = get_runs(cursor, run, sample)
    return runs[0] if runs else None
    
def get_runs(cursor, run=None, sample=None):
    cmd = "SELECT * FROM run"
    where = []
    args = []
    if run:
        where.append("run_name=?")
        args.append(run)
    if sample:
        where.append("sample_name=?")
        args.append(sample)
    if where:
        cmd += " WHERE " + " AND ".join(where)
    cursor.execute(cmd, args)
    results = results_as_dict(cursor)
    return results
    
def get_num_expected_mutations(cursor):
    cmd = "SELECT count(*) FROM mutation WHERE is_expected=1"
    cursor.execute(cmd)
    ans = cursor.fetchone()
    return ans[0] if ans else None

def update_run_counts(cursor, run_id):
#    tot_expected = get_num_expected_mutations(cursor)
#    limit = int(tot_expected/2)
    cmd = "UPDATE run SET " +\
          " num_mutations=(SELECT COUNT(*) FROM vaf WHERE run_id=?)," +\
          " num_expected=(SELECT COUNT(*) FROM vaf v JOIN mutation m" +\
          " ON m.id=v.mutation_id WHERE m.is_expected=? AND v.run_id=?)" +\
          " WHERE run.id=?"
    cursor.execute(cmd, [run_id, 1, run_id, run_id])
#    cmd = "UPDATE run SET run_status = ? WHERE run.id=?" +\
#          " AND run.num_mutations < {}".format(limit)
#    cursor.execute(cmd, ['FAIL', run_id])

def delete_vafs_for_run(cursor, run_id):
    cursor.execute("DELETE FROM vaf WHERE run_id=?", (run_id,))

def get_vafs_for_run(cursor, run_id, run_status=None):
    cmd = "SELECT * FROM vaf v, mutation m, run r ON v.mutation_id=m.id" +\
          " AND v.run_id=r.id WHERE v.run_id=?" 
    args = [run_id,]
    if run_status:
        cmd += " AND r.run_status=?"
        args.append(run_status)
    cursor.execute(cmd, args)
    results = results_as_dict(cursor)
    return results

def get_vafs_for_all_runs(cursor):
    cmd = "SELECT * FROM vaf v, mutation m, run r ON v.mutation_id=m.id" +\
          " AND v.run_id=r.id" 
    cursor.execute(cmd)
    results = results_as_dict(cursor)
    return results

#-----------------------------------------------------------------------------

def truths_from_db(cursor):
    cmd = "SELECT * FROM mutation WHERE is_expected=?"
    cursor.execute(cmd, [1,])
    columns = [ d[0] for d in cursor.description ]
    data = []
    datadict = {}
    for ans in cursor.fetchall():
        d = dict(zip(columns, ans))
        dkey = create_dkey(d)
        d['dkey'] = dkey
        data.append(d)
        datadict[dkey] = d
    columns.remove('id')
    columns.remove('is_expected')
    columns.remove('last_modified')
    return { 'data': data, 'datadict': datadict, 'fields': columns }

def db_summary(cursor):
    msgs = []
    cursor.execute('SELECT COUNT(*)' +\
        ', (SELECT COUNT(*) FROM run WHERE run_status=?)'*2 +\
        ' FROM run', ['PASS', 'FAIL'])
    ans = cursor.fetchone()
    msgs.append("    {} runs stored ({} good; {} failed)\n".format(ans[0],
                ans[1], ans[2]))
    cursor.execute('SELECT COUNT(*) FROM mutation WHERE is_expected=1')
    ans = cursor.fetchone()
    msgs.append("    {} expected mutations\n".format(ans[0]))
    cursor.execute('SELECT COUNT(*) FROM mutation WHERE is_expected=0')
    ans = cursor.fetchone()
    msgs.append("    {} unexpected mutations\n".format(ans[0]))
    cursor.execute('SELECT COUNT(*) FROM vaf')
    ans = cursor.fetchone()
    msgs.append("    {} vafs\n".format(ans[0]))
    return msgs

def check_db(dbfile, schemafile, truthfile):
    is_new_db = not os.path.exists(dbfile)
    dbh = connect_db(dbfile)
    cursor = dbh.cursor()
    msgs = []
    if is_new_db:
        tinfo = parse_truth(truthfile)
        add_schema(dbh, schemafile)
        msgs.append("  Saving truths\n")
        save_mutations(cursor, tinfo['data'], is_expected=1)
        dbh.commit()
        cursor.execute('SELECT COUNT(*) FROM mutation')
        msgs.append("    {} rows inserted\n".format(cursor.fetchone()[0]))
    else:
        tinfo = truths_from_db(cursor)
        msgs = db_summary(cursor)
    sys.stderr.write(''.join(msgs))
    return (dbh, tinfo, msgs)

def save2db(dbh, runname, sample, status, vinfo, force=False):
    if not runname or not sample:
        if sample:
            sys.stderr.write("  Need run name for sample {}\n".format(sample))
        elif runname:
            sys.stderr.write("  Need sample for run {}\n".format(runname))
        else:
            sys.stderr.write("  Need sample and run name\n")
        sys.stderr.flush()
        return 0
    cursor = dbh.cursor()
    run = get_run(cursor, runname, sample)
    if force or not run:
        if run and force:
            sys.stderr.write('  Deleting old data for {}:{} in db.\n'.format(
                             runname, sample))
            update_run(cursor, runname, sample, status)
            delete_vafs_for_run(cursor, run['id'])
        elif not run:
            sys.stderr.write('  Saving run {}:{} in db.\n'.format(runname, 
                             sample))
            save_run(cursor, runname, sample, status)
            run = get_run(cursor, runname, sample)
        for d in vinfo['data']:
            save_vaf(cursor, run['id'], d)
        update_run_counts(cursor, run['id'])
        dbh.commit()
    vafs = get_vafs_for_run(cursor, run['id'])
    sys.stderr.write('  Have {} mutations for run {}:{} in db.\n'.format(
                     len(vafs), runname, sample))
    sys.stderr.flush()
    return 1


#----spreadsheet.py-----------------------------------------------------------

def convert_to_excel_col(colnum):
    mod = colnum % 26
    let = chr(mod+65)
    if colnum > 26:
        rep = colnum/26
        let1 = chr(rep+64)
        let = let1 + let
    return let

def add_formats_to_workbook(workbook):
    wbformat = {}
    wbformat['yellow'] = workbook.add_format({'bg_color': '#FFFF00',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['gold'] = workbook.add_format({'bg_color': '#FFC000', })
    wbformat['red'] = workbook.add_format({'bg_color': '#FF0000', })
    wbformat['green'] = workbook.add_format({'bg_color': '#92D050', })
    return wbformat

def print_spreadsheet_excel(header, data, outfile, sheetname=None):
#    sys.stderr.write("  Writing {}\n".format(outfile))
    if sheetname and len(sheetname)>30:
        sheetname = sheetname[:30]
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet(sheetname)
    wbformat = add_formats_to_workbook(workbook)
    numlines = 0
    for i, rowdat in enumerate(header+data):
        fmt = wbformat[rowdat.highlight] if rowdat.highlight else None
        numlines += 1
        for j, r in enumerate(rowdat.data):
            if r.isdigit():
                r = int(r)
            elif is_float(r):
                r = float(r)
            if not rowdat.cell or j==rowdat.cell:
                worksheet.write(i, j, r, fmt)
            else:
                worksheet.write(i, j, r)
#    worksheet.freeze_panes(len(header), 0)
    workbook.close()
    wb = openpyxl.load_workbook(outfile)
    wb.save(outfile)
    return numlines

#----gui.py-------------------------------------------------------------------

class StampQC_App(wx.App):
    def __init__(self, dbh, tinfo, db, spreadsheet, msg=None, **kwargs):
        self.dbh = dbh
        self.tinfo = tinfo
        self.db = db
        self.spreadsheet = spreadsheet
        self.msg = msg
        wx.App.__init__(self, kwargs)

    def OnInit(self):
        self.frame = StampFrame(self.dbh, self.tinfo, self.db,
                                self.spreadsheet, msg=self.msg)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

class StampFrame(wx.Frame):
    def __init__(self, dbh, tinfo, db, spreadsheet, msg=None):
        wx.Frame.__init__(self, None, title="STAMP TruQ3 QC", size=(550,500))
        self.dbh = dbh
        self.tinfo = tinfo
        self.db = db
        self.spreadsheet = spreadsheet

        panel = wx.Panel(self)
        label = wx.StaticText(panel, -1, "Drop TruQ3 variant reports here:")
        self.text = wx.TextCtrl(panel,-1, "",style=wx.TE_READONLY|
                                wx.TE_MULTILINE|wx.HSCROLL)
        button_print = wx.Button(panel, -1, "Print reports")
        print_tooltip = "Creates new variant reports with variants "+\
            "labelled expected, not expected or not found. New reports "+\
            "are named <Sample>.variant_report.checked.txt and saved "+\
            "in same folder as original report."
        button_print.SetToolTip(wx.ToolTip(print_tooltip))
        self.Bind(wx.EVT_BUTTON, self.PrintReports, button_print)
        button_save = wx.Button(panel, -1, "Update spreadsheet and DB")
        save_tooltip = "Update spreadsheet and database with data "+\
            "entered.\nSpreadsheet:  {}\n".format(spreadsheet)+\
            "Database:  {}\n".format(db)
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
        button_sizer.Add(button_print, 0, wx.ALIGN_CENTER_VERTICAL)
        button_sizer.Add(button_save, 0, wx.ALIGN_CENTER_VERTICAL)
        button_sizer.AddStretchSpacer()
        button_sizer.Add(button_quit, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(button_sizer, 0, wx.ALL|wx.EXPAND, 5)
        panel.SetSizer(sizer)

        dt = VariantReportDrop(self.text, self.notebook, tinfo)
        self.text.SetDropTarget(dt)

    def PrintReports(self, event):
        self.text.AppendText("\nPrinting reports:\n")
        if not self.notebook.results:
            self.text.AppendText("  No reports to process.\n\n")
            return
        for i, info in enumerate(self.notebook.results):
            if not info: continue
            entries = self.notebook.entries[i]
            sample = entries['sample'].GetValue()
            outfile = info['file'].replace('.txt','') + ".checked.txt"
            if not sample==info['sample']:
                outfile = outfile.replace(info['sample'], sample)
            self.text.AppendText("  "+outfile+"\n")
        self.text.AppendText("\n")

    def UpdateSpreadsheetAndDB(self, event):
        self.text.AppendText("\nUpdating data:\n")
        if not self.notebook.results:
            self.text.AppendText("  No data to save to db.\n")
        else:
            for i, info in enumerate(self.notebook.results):
                if not info: continue
                entries = self.notebook.entries[i]
                sample = entries['sample'].GetValue()
                run = entries['run'].GetValue()
                if not run or not sample:
                    msg = "    {}: Not saved.".format(i)
                    if not run and not sample:
                        msg += "  Need run and sample\n"
                    elif not run:
                        msg += "  Need run name\n"
                    else:
                        msg += "  Need sample name\n"
                    self.text.AppendText(msg)
                    continue
                statusnum = entries['status'].GetSelection()
                status = entries['status'].GetString(statusnum)
                saved = save2db(self.dbh, run, sample, status, info['vinfo'], True)
                if saved:
                    self.text.AppendText("        Saved {} data to db.\n".format(sample))
                else:
                    self.text.AppendText("        {} not saved to db.\n".format(sample))
            summ = db_summary(self.dbh.cursor())
            self.notebook.tabOne.ChangeMessage(''.join(summ))
        try:
            self.text.AppendText("  Updating spreadsheet.\n")
            res = generate_excel_spreadsheet(self.dbh, 
                  self.tinfo['fields'], self.spreadsheet)
            self.text.AppendText("      Spreadsheet now contains "+\
                "{} runs and {} unique variants\n".format(res['num_runs'],
                res['num_variants']))
            if res['failedruns']:
                self.text.AppendText(
                     "      Failed runs not included: {}\n".format(
                     ", ".join(res['failedruns'])))
        except Exception, e:
            self.text.AppendText("    ERROR: {}{}\n\n".format(
                                       type(e).__name__, e))
            raise
        self.text.AppendText("\n")

    def OnCloseMe(self, event):
        self.Close(True)

    def OnCloseWindow(self, event):
        self.Destroy()
        
class VariantReportDrop(wx.FileDropTarget):
    def __init__(self, window, notebook, tinfo):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.notebook = notebook
        self.tinfo = tinfo
        self.num_files = 0

    def OnDropFiles(self, x, y, filenames):
        counts = {}
        for variant_file in filenames:
            self.num_files += 1
            self.window.AppendText("File {}:    {}\n".format(self.num_files, 
                                   variant_file))
            try:
                (vinfo, run, sample) = parse_variant_file(variant_file)
                summary = compare_variants(self.tinfo['datadict'], vinfo, 
                                           counts)
                info = ({'num': self.num_files, 'file':variant_file,
                         'vinfo': vinfo, 'summary': summary, 
                         'status': summary['Status'],
                         'run': run, 'sample': sample })
                title = "{}: {}".format(self.num_files, run)
                self.notebook.AddResultsTab(info, title=title)
            except KeyError, e:
                self.window.AppendText("    ERROR:  Bad file format.  " +\
                            "This does not look like a variant report.\n")
            except Exception, e:
                self.window.AppendText("    ERROR: {} {}\n\n".format(
                                       type(e).__name__, e))
                raise

class StampNotebook(fnb.FlatNotebook):
    def __init__(self, parent, msg=None):
        fnb.FlatNotebook.__init__(self, parent, id=wx.ID_ANY, size=(500, 200),
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
            title = "File {}".format(num)
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

    def OnTabDrop(self, event):
        selected = self.GetSelection()
        oldselected = event.GetOldSelection()
        res = self.results.pop(oldselected)
        ent = self.entries.pop(oldselected)
        self.results.insert(selected, res)
        self.entries.insert(selected, ent)

class TabPanel_Text(wx.Panel):
    def __init__(self, parent, msg="\n\n\n\n"):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)
        self.textWidget = wx.StaticText(self, -1, '\n'+msg)

    def ChangeMessage(self, msg):
        self.textWidget.Destroy()
        self.textWidget = wx.StaticText(self, -1, '\n'+msg)

class TabPanel_Results(wx.Panel):
    def __init__(self, parent, info):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)

        runLabel = wx.StaticText(self, -1, "Run:")
        runEntry = wx.TextCtrl(self, -1, info['run'])
        sampleLabel = wx.StaticText(self, -1, "Sample:")
        sampleEntry = wx.TextCtrl(self, -1, info['sample'])
        statusLabel = wx.StaticText(self, -1, "Status:")
#        statusEntry = wx.TextCtrl(self, -1, info['status'])
        statusEntry = wx.Choice(self, -1, choices=['PASS', 'FAIL'])
        statusEntry.SetSelection(0 if info['status']=='PASS' else 1)
        parent.entries.append({'run': runEntry, 'sample': sampleEntry,
                               'status': statusEntry})
        msg = 'All expected mutations found.'
        if len(info['summary']['notseen'])>0:
            msg = "Expected mutations not found:{:5d}".format(
                   len(info['summary']['notseen']))
        infostr = "Total mutations:{:9d}".format(info['summary']['Total'])+\
            "              {}\n".format(msg)+\
            "  Expected found:{:8d}\n".format(info['summary']['Expected'])+\
            "  Unexpected found:{:4d}\n".format(info['summary']['Unexpected'])

        infoText = wx.StaticText(self, -1, infostr)

        panelSizer = wx.BoxSizer(wx.VERTICAL)
        entrySizer = wx.FlexGridSizer(cols=2, hgap=5, vgap=5)
        entrySizer.AddGrowableCol(1)
        entrySizer.Add(runLabel, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        entrySizer.Add(runEntry, 0, wx.EXPAND)
        entrySizer.Add(sampleLabel, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
        entrySizer.Add(sampleEntry, 0, wx.EXPAND)
        entrySizer.Add(statusLabel, 0, wx.ALIGN_RIGHT|wx.ALIGN_CENTER_VERTICAL)
#        entrySizer.Add(statusEntry, 0, wx.EXPAND)
        entrySizer.Add(statusEntry, 0)
        panelSizer.Add(entrySizer, 0, wx.EXPAND|wx.ALL, 10)
        panelSizer.Add(infoText, 0, wx.ALIGN_LEFT)
        self.SetSizer(panelSizer)

def run_gui(dbh, tinfo, db, spreadsheet, msg):
    app = StampQC_App(dbh, tinfo, db, spreadsheet, msg=msg)
    app.MainLoop()

#-----------------------------------------------------------------------------

class TabData:
    def __init__(self, data=None, fields=None, header=None, numlines=None):
        self.data = data
        self.fields = fields
        self.header = header
        self.numlines = numlines

def parse_tab_file(tabfile, commentstart='#'):
    header = []
    fields = []
    data = []
    numlines = 0
    with open(tabfile, 'r') as fh:
        for line in fh:
            numlines += 1
            if line.startswith(commentstart):
                header.append(line.rstrip())
            elif fields: # have fields so data lines follow
                data.append(line.rstrip().split("\t"))
            else:
                fields = line.rstrip().split("\t")
    return TabData(data, fields, header, numlines)

class RowData:
    def __init__(self, data=None, highlight=None, cell=None):
        self.data = data
        self.highlight = highlight
        self.cell = cell

def create_depth_report_xlsx(report, args):
    outfile = outfile_name(report, args.outdir)
    if args.debug:
        sys.stderr.write("    Writing {}\n".format(outfile))
    sheetname = os.path.basename(outfile).replace('.xlsx','')
    header = []
    fields = None
    data = defaultdict(list)
    tabdata = parse_tab_file(report)
    header = [ RowData([l,]) for l in tabdata.header ]
    header.append(RowData(tabdata.fields))
    i_mindepth = None
    if 'Min Depth' in tabdata.fields:
        i_mindepth = tabdata.fields.index('Min Depth')
    elif 'Min_Depth' in tabdata.fields:
        i_mindepth = tabdata.fields.index('Min_Depth')
    else:
        sys.exit("{} Bad format.  ".format(report) +\
                 "Min Depth column not found.")
    for row in tabdata.data:
        data[int(row[i_mindepth])].append(row)
    rows = []
    for mindepth, row in sorted(data.items()):
        hi = 'yellow' if mindepth < 200 else None
        for r in row:
            rows.append(RowData(r, hi))
    numxlines = print_spreadsheet_excel(header, rows, outfile, sheetname)
    if tabdata.numlines != numxlines:
        sys.stderr.write("    {} lines in report\n".format(tabdata.numlines))
        sys.stderr.write("    {} lines in spreadsheet\n".format(numxlines))
        sys.exit("  ERROR: Num lines don't match\n")
    return tabdata


def create_variant_report_xlsx(report, args):
    outfile = outfile_name(report, args.outdir)
    if args.debug:
        sys.stderr.write("    Writing {}\n".format(outfile))
    sheetname = os.path.basename(outfile).replace('.xlsx','')
    fields = None
    highlight_row = RowData(['']*26, 'gold')
    tabdata = parse_tab_file(report)
    header = [ RowData([l,]) for l in tabdata.header ]
    header.append(RowData(tabdata.fields))
    data = []
    i_status = None
    if 'Status' in tabdata.fields:
        i_status = tabdata.fields.index('Status')
    else:
        sys.exit("{} Bad format.  ".format(report) +\
                 "Status column not found.")
    for row in tabdata.data:
        if row[i_status]=='NOT_REPORTED':
            if highlight_row:
                data.append(highlight_row)
                highlight_row = None
            data.append(RowData(row))#, 'red', i_status))
        elif row[i_status]=='ACCEPT':
            data.append(RowData(row))#, 'green', i_status))
        else:
            data.append(RowData(row))
    numxlines = print_spreadsheet_excel(header, data, outfile, sheetname)
    if tabdata.numlines+1 != numxlines:
        sys.stderr.write("    {} lines in report\n".format(tabdata.numlines))
        sys.stderr.write("    {} lines in spreadsheet\n".format(numxlines))
        sys.exit("  ERROR: Unexpected num lines\n")
    return tabdata

def split_vcf(vcffile, vinfo, args):
    label = vcffile.replace('.vcf', '')
    if args.outdir:
        label = os.path.join(args.outdir, os.path.basename(label))
    acceptfile = label + '_accepted.vcf'
    rejectfile = label + '_rejected.vcf'
    i_status = vinfo.fields.index('Status')
    i_chrom = vinfo.fields.index('Chr')
    i_pos = vinfo.fields.index('Position')
    variantdata = defaultdict(dict)
    for row in vinfo.data:
        chrom = row[i_chrom].replace('chr', '')
        pos = int(row[i_pos])
        variantdata[chrom][pos] = row[i_status]
    vcfhead = []
    vcfaccept = []
    vcfreject = []
    with open(vcffile, 'r') as fh:
        for line in fh:
            if line.startswith('#'):
                vcfhead.append(line)
            else:
                row = line.split("\t", 3)
                chrom = row[0]
                pos = int(row[1])
                if variantdata[chrom][pos]=='NOT_REPORTED':
                    vcfreject.append(line)
                else:
                    vcfaccept.append(line)
    with open(acceptfile, 'w') as ofh:
        ofh.write(''.join(vcfhead))
        ofh.write(''.join(vcfaccept))
    with open(rejectfile, 'w') as ofh:
        ofh.write(''.join(vcfhead))
        ofh.write(''.join(vcfreject))


def group_files_by_sample(inputfiles, args):
    extensions = {
        '.vcf': 'vcf',
        '.depth_report_indels.txt': 'dp_indels', 
        '.depth_report_snvs.txt': 'dp_snvs',
        '.variant_report.txt': 'v_report', }
    samples = defaultdict(dict)
    infiles = []
    for in_arg in inputfiles:
        if os.path.isfile(in_arg):
            infiles.append(in_arg)
        elif os.path.isdir(in_arg):
            infiles.extend([ os.path.join(in_arg, f) for f in \
                             os.listdir(in_arg) ])
    for infile in infiles:
        if infile.endswith('_accepted.vcf') or \
           infile.endswith('_rejected.vcf'):
            continue
        for ext in extensions:
            if infile.endswith(ext):
                sample = os.path.basename(infile).replace(ext,'')
                samples[sample][extensions[ext]] = infile
    return samples

#-----------------------------------------------------------------------------

if __name__=='__main__':
    descr = "This script post-processes STAMP report files."
    descr += " Depth reports will be sorted by Min depth with values less"
    descr += " than 200 highlighted and saved as Excel."
    descr += " Variant reports will have a highlighted row inserted between"
    descr += " ACCEPTED and NOT_REPORTED variants and saved as Excel."
    descr += " If VCF and variant reports are available, the VCF will be"
    descr += " split into accepted and rejected VCF files."
    parser = ArgumentParser(description=descr)
    parser.add_argument("reports", nargs="*",
                        help="STAMP depth and/or variant report(s)")
    parser.add_argument("-o", "--outdir", 
                        help="Directory to save output file(s)")
    parser.add_argument("--debug", default=False, action='store_true',
                        help="Write debugging messages")

    args = parser.parse_args()
    if len(args.reports)==0:
        print("GUI not implemented yet")
#        run_gui(dbh, tinfo, args.db, args.spreadsheet, msg=''.join(msgs))
    else:
        samples = group_files_by_sample(args.reports, args)
        for sample, d in sorted(samples.items()):
            sys.stderr.write("\nSample {}\n".format(sample))
            sys.stderr.write("- Sorting indel depth report: ")
            dpindelinfo = None
            dpsnvinfo = None
            vinfo = None
            if 'dp_indels' in d:
                sys.stderr.write(" YES\n")
                dpindelinfo = create_depth_report_xlsx(d['dp_indels'], args)
            else:
                sys.stderr.write(" NO\n")
            sys.stderr.write("- Sorting snv depth report: ")
            if 'dp_snvs' in d:
                sys.stderr.write(" YES\n")
                dpsnvinfo = create_depth_report_xlsx(d['dp_snvs'], args)
            else:
                sys.stderr.write(" NO\n")
            sys.stderr.write("- Formatting variant report: ")
            if 'v_report' in d:
                sys.stderr.write(" YES\n")
                vinfo = create_variant_report_xlsx(d['v_report'], args)
            else:
                sys.stderr.write(" NO\n")
            sys.stderr.write("- Splitting vcf: ")
            if 'vcf' in d and vinfo:
                sys.stderr.write(" YES\n")
                split_vcf(d['vcf'], vinfo, args)
            else:
                sys.stderr.write(" NO\n")
#            sys.stderr.write("- Generating low coverage comment: ")
#            if 'dp_indels' in d and 'dp_snvs' in d:
#                sys.stderr.write(" YES")
#            else:
#                sys.stderr.write(" NO")





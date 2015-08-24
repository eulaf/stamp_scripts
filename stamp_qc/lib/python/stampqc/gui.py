import wx
import dbops
import fileops
import spreadsheet

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
        label = wx.StaticText(panel, -1, "Drop variant reports here:")
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
            entries = self.notebook.entries[i]
            sample = entries['sample'].GetValue()
            outfile = info['file'].replace('.txt','') + ".checked.txt"
            if not sample==info['sample']:
                outfile = outfile.replace(info['sample'], sample)
            self.text.AppendText("  "+outfile+"\n")
            fileops.print_checked_file(info, self.tinfo, outfile)
        self.text.AppendText("\n")

    def UpdateSpreadsheetAndDB(self, event):
        self.text.AppendText("\nUpdating data:\n")
        if not self.notebook.results:
            self.text.AppendText("  No data to save to db.\n")
        else:
            for i, info in enumerate(self.notebook.results):
                entries = self.notebook.entries[i]
                stamprun = entries['run'].GetValue()
                sample = entries['sample'].GetValue()
                dbops.save2db(self.dbh, stamprun, sample, info['vinfo'])
                self.text.AppendText("  Saved {} data to db.\n".format(sample))
            summ = dbops.db_summary(self.dbh.cursor())
            self.notebook.tabOne.ChangeMessage(''.join(summ))
        try:
            self.text.AppendText("  Updating spreadsheet.\n")
            res = spreadsheet.generate_excel_spreadsheet(self.dbh, 
                  self.tinfo['fields'], self.spreadsheet)
            self.text.AppendText("    Spreadsheet now contains "+\
                "{} runs and {} unique variants\n".format(res['num_runs'],
                res['num_variants']))
            if res['failedruns']:
                self.text.AppendText(
                     "    Failed runs not included: {}\n".format(
                     ", ".join(res['failedruns'])))
        except Exception, e:
            self.text.AppendText("    ERROR: {}{}\n\n".format(
                                       type(e).__name__, e))
            raise
        self.text.AppendText("\n")
        pass


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
                (vinfo, stamprun, sample) = fileops.parse_variant_file(
                                                 variant_file)
                summary = fileops.compare_variants(self.tinfo['datadict'], 
                                                    vinfo, counts)
                info = ({'num': self.num_files, 'file':variant_file,
                         'vinfo': vinfo, 'summary': summary, 
                         'stamprun': stamprun, 'sample': sample })
                title = "{}: {}".format(self.num_files, stamprun)
                self.notebook.AddResultsTab(info, title=title)
                self.notebook.results.append(info)
            except KeyError, e:
                self.window.AppendText("    ERROR:  Bad file format.  " +\
                            "This does not look like a variant report.\n")
            except Exception, e:
                self.window.AppendText("    ERROR: {} {}\n\n".format(
                                       type(e).__name__, e))
                raise


class StampNotebook(wx.Notebook):
    def __init__(self, parent, msg=None):
        wx.Notebook.__init__(self, parent, id=wx.ID_ANY, size=(500, 175))

        self.tabOne = TabPanelText(self, msg=msg)
        self.AddPage(self.tabOne, "DB content")
        self.results = []
        self.entries = []

    def AddResultsTab(self, info, title=None):
        if not title:
            num = info['num'] if 'num' in info else ''
            title = "File {}".format(num)
        newTab = TabPanelResults(self, info)
        self.AddPage(newTab, title)
        numpages = self.GetPageCount()
        self.SetSelection(numpages-1)

class TabPanelText(wx.Panel):
    def __init__(self, parent, msg="\n\n\n\n"):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)
        self.textWidget = wx.StaticText(self, -1, '\n'+msg)

    def ChangeMessage(self, msg):
        self.textWidget.Destroy()
        self.textWidget = wx.StaticText(self, -1, '\n'+msg)

class TabPanelResults(wx.Panel):
    def __init__(self, parent, info):
        wx.Panel.__init__(self, parent=parent, id=wx.ID_ANY)

        runLabel = wx.StaticText(self, -1, "Run:")
        runEntry = wx.TextCtrl(self, -1, info['stamprun'])
        sampleLabel = wx.StaticText(self, -1, "Sample:")
        sampleEntry = wx.TextCtrl(self, -1, info['sample'])
        parent.entries.append({'run': runEntry, 'sample': sampleEntry})
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
        panelSizer.Add(entrySizer, 0, wx.EXPAND|wx.ALL, 10)
        panelSizer.Add(infoText, 0, wx.ALIGN_LEFT)
        self.SetSizer(panelSizer)


if __name__ == '__main__':
    app = StampQC_App()
    app.MainLoop()


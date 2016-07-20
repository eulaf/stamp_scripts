#!/usr/bin/env python

"""
Variant report:
Separate ACCEPTED, CHECK_COMPOUND, CHECK_1-5PCT from NOT_REPORTED 
by yellow hightlighted row and save as Excel file.  Add comment
snippet

Variant report and VCF file:
Create accepted and rejected VCFs where NOT_REPORTED variants
are in the rejected VCF and all others in the accepted VCF

Depth report: 
Sort depth files by min depth, highlight rows < 200 and
save as Excel file.

SNV and indel depth reports:
Create text file with low coverage comment.

Fusion reports with transcripts:
Add transcripts to fusions.filtered.txt files

"""

import os
import sys
import xlsxwriter
import openpyxl
import re
import wx
import wx.richtext 
from collections import defaultdict
from argparse import ArgumentParser

VERSION="1.2"
BUILD="160719"

# version 1.2 changes 160719
#   add transcripts to fusion files
#   add comment to variant report xlsx

# version 1.2 changes 160629
#   low coverage comment - get gene name by split on '-'
#   group samples - ignore files ending in 'unfiltered.vcf'
#   split vcf - sort by chr, pos

#----common.py----------------------------------------------------------------

def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))

MINCOVERAGE = 200 # min depth coverage for STAMP
MALE_MINCOV = 60 # chrY min coverage to determine if sample likely male
LOWCOV_TEXT = "Portions of the following gene(s) failed to meet the"+\
    " minimum coverage of {}x: GENELIST.".format(MINCOVERAGE) +\
    " Low coverage may adversely affect the sensitivity of the assay."+\
    " If clinically indicated, repeat testing on a new specimen can"+\
    " be considered."

FUSION_TRANSCRIPT_FILE = os.path.join(getScriptPath(), os.pardir, "docs", 
                         "stamp2_fusion_gene_transcripts.txt")
FUSION_TRANSCRIPTS = {}

AAPATT = re.compile('([A-Z])')
AA_convert = {
    'A': ['Ala', 'Alanine',], 
    'R': ['Arg', 'Arginine',], 
    'N': ['Asn', 'Asparagine',], 
    'D': ['Asp', 'Aspartic acid',], 
    'C': ['Cys', 'Cysteine',], 
    'Q': ['Gln', 'Glutamine',], 
    'E': ['Glu', 'Glutamic acid',], 
    'G': ['Gly', 'Glycine',], 
    'H': ['His', 'Histidine',], 
    'I': ['Ile', 'Isoleucine',], 
    'L': ['Leu', 'Leucine',], 
    'K': ['Lys', 'Lysine',], 
    'M': ['Met', 'Methionine',], 
    'F': ['Phe', 'Phenylalanine',], 
    'P': ['Pro', 'Proline',], 
    'O': ['Pyl', 'Pyrrolysine',], 
    'S': ['Ser', 'Serine',], 
    'U': ['Sec', 'Selenocysteine',], 
    'T': ['Thr', 'Threonine',], 
    'W': ['Trp', 'Tryptophan',], 
    'Y': ['Tyr', 'Tyrosine',], 
    'V': ['Val', 'Valine',], 
    'B': ['Asx', 'Aspartic acid or Asparagine',], 
    'Z': ['Glx', 'Glutamic acid or Glutamine',], 
    'J': ['Xle', 'Leucine of Isoleucine',], 
    'X': ['Ter', 'Termination codon',], 
}

def aa_expand(matchobj, num=0):
    m = matchobj.group(0)
    expansion = m
    if m in AA_convert:
        expansion = AA_convert[m][num]
    return expansion

def is_float(v):
    try:
        float(v)
    except ValueError:
        return False
    return True

#----spreadsheet.py-----------------------------------------------------------

def convert_to_excel_col(colnum):
    mod = colnum % 26
    let = chr(mod+65)
    if colnum > 26:
        rep = colnum/26
        let1 = chr(rep+64)
        let = let1 + let
    return let

class ExcelRowData:
    def __init__(self, data=None, highlight=None, cell=None):
        self.data = data
        self.highlight = highlight
        self.cell = cell

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


#----fileops.py---------------------------------------------------------------

class TabData:
    def __init__(self, tabfile=None, data=None, fields=None, header=None, 
                 numlines=None, outfile=None):
        self.tabfile = tabfile
        self.data = data
        self.fields = fields
        self.header = header
        self.numlines = numlines
        self.outfile = outfile

def parse_tab_file(tabfile, outfile=None, commentstart='#'):
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
    return TabData(tabfile, data, fields, header, numlines, outfile)

def outfile_name(report, outdir=None, outext='', inext='.txt'):
    outfile = report.replace(inext, '') 
    if outext:
        outfile += outext
    if outdir:
        outfile = os.path.join(outdir, os.path.basename(outfile))
    return outfile

#-----------------------------------------------------------------------------

def create_depth_report_xlsx(report, args):
    outfile = outfile_name(report, args.outdir, '.xlsx')
    if args.debug:
        sys.stderr.write("    Writing {}\n".format(outfile))
    sheetname = os.path.basename(outfile).replace('.xlsx','')
    header = []
    fields = None
    data = defaultdict(list)
    tabdata = parse_tab_file(report, outfile=outfile)
    header = [ ExcelRowData([l,]) for l in tabdata.header ]
    header.append(ExcelRowData(tabdata.fields))
    i_mindepth = None
    if 'Min Depth' in tabdata.fields:
        i_mindepth = tabdata.fields.index('Min Depth')
    elif 'Min_Depth' in tabdata.fields:
        i_mindepth = tabdata.fields.index('Min_Depth')
    else:
        sys.exit("{} Bad format.  ".format(report) +\
                 "Min Depth column not found.")
    for row in tabdata.data: # create dict keyed by min depth
        data[int(row[i_mindepth])].append(row)
    rows = []
    for mindepth, row in sorted(data.items()):
        hi = 'yellow' if mindepth < MINCOVERAGE else None
        for r in row:
            rows.append(ExcelRowData(r, hi))
    numxlines = print_spreadsheet_excel(header, rows, outfile, sheetname)
    if tabdata.numlines != numxlines:
        sys.stderr.write("    {} lines in report\n".format(tabdata.numlines))
        sys.stderr.write("    {} lines in spreadsheet\n".format(numxlines))
        sys.exit("  ERROR: Num lines don't match\n")
    return tabdata

def generate_low_coverage_comment(outlabel, dpindelinfo, dpsnvinfo):
    low_cov_genes = {}
    is_female = True
    for tabdata in (dpindelinfo, dpsnvinfo):
        i_chr = None
        i_mindepth = None
        i_description = None
        if 'Min Depth' in tabdata.fields:
            i_mindepth = tabdata.fields.index('Min Depth')
        elif 'Min_Depth' in tabdata.fields:
            i_mindepth = tabdata.fields.index('Min_Depth')
        else:
            sys.exit("{} Bad format.  ".format(tabdata.tabfile) +\
                     "Min Depth column not found.")
        if 'Description' in tabdata.fields:
            i_description = tabdata.fields.index('Description')
        else:
            sys.exit("{} Bad format.  ".format(tabdata.tabfile) +\
                     "Description column not found.")
        if 'Chr' in tabdata.fields:
            i_chr = tabdata.fields.index('Chr')
        else:
            sys.exit("{} Bad format.  ".format(tabdata.tabfile) +\
                     "Chr column not found.")
        for row in tabdata.data: 
            mindepth = int(row[i_mindepth])
            if mindepth < MINCOVERAGE:
                # '_' for STAMPv1 regions; '-' for STAMPv2 regions
                gene = row[i_description].split('_')[0].split('-')[0]
                low_cov_genes[gene] = row[i_chr]
            if row[i_chr]=='chrY' and mindepth >= MALE_MINCOV:
                is_female = False
    genestr = ', '.join(sorted(low_cov_genes))
    genestr = ', and '.join(genestr.rsplit(', ', 1))
    male_lcc = LOWCOV_TEXT.replace('GENELIST', genestr)
    if is_female:
        noY_genestr = ', '.join(sorted([g for g in low_cov_genes if \
                      low_cov_genes[g] != 'chrY']))
        noY_genestr = ', and '.join(noY_genestr.rsplit(', ', 1))
        female_lcc = LOWCOV_TEXT.replace('GENELIST', noY_genestr)
        lcc = 'All chrY regions have coverage < {}.\n'.format(MALE_MINCOV)
        lcc += "FEMALE (no chrY genes):\n{}\n\nMALE:\n{}\n".format(
               female_lcc, male_lcc)
    else:
        lcc = male_lcc
    outfile = outlabel + '.low_coverage_comment.txt'
    with open(outfile, 'w') as ofh:
        ofh.write(lcc + '\n')
    sys.stderr.write(lcc + '\n')
    sys.stderr.flush()
    return outfile, is_female

def is_substitution_bases(cdot):
    """Checks if mutation is simple substitution"""
    flag = False
    if '>' in cdot:
        bases = cdot.lstrip('1234567890').split('>')
        if len(bases)==2 and len(bases[0])==1 and len(bases[1])==1:
            flag = True
    return flag 

def aa_change_names_and_codon(pdot):
    """Returns full amino acid names and codon for amino acids in AA change"""
    match = re.match(r"([A-Z])(\d+)([A-Z])$", pdot)
    result = None
    if match:
        (aa1, codon, aa2) = match.groups()
        aa1_name = AA_convert.get(aa1, ['',''])[1].lower()
        aa2_name = AA_convert.get(aa2, ['',''])[1].lower()
        if aa1_name and aa2_name:
            result = [aa1_name, codon, aa2_name]
    return result

def add_comment_snippet(tabdata):
    """Variant comments crafted by the fellows generally start with a set
    format that can be auto-generated to save them time.  For examples:  
    1. frameshifts
    The p.Val166fs (c.491dup, p.V166fs) mutation in the PTEN gene results in a
    frame shift
    2. termination codon
    The p.Glu219Ter (c.655G>T, p.E219X) mutation in the KEAP1 gene results in a
    premature termination codon at amino acid position 219
    3. simple substitutions: 
    The p.Thr41Ala (c.121A>G, p.T41A) mutation in the CTNNB1 gene results in a
    substitution of alanine for threonine at codon 41
    4. If pdot and AA Change are identical
    The p.747_753del (c.2240_2257del) mutation in the EGFR gene
    """
    tabdata.fields.append('Comment snippet')
    i_aa = tabdata.fields.index('AA Change')
    i_cds = tabdata.fields.index('CDS Change')
    i_gene = tabdata.fields.index('Gene')
    for row in tabdata.data:
        if row[i_aa] != '.' and row[i_cds] != '.':
            pdot = 'p.'+AAPATT.sub(aa_expand, row[i_aa])
            cdot = 'c.'+row[i_cds]
            gene = row[i_gene]
            mutation = '{} ({}, p.{})'.format(pdot, cdot, row[i_aa]) \
                if pdot != 'p.'+row[i_aa] else '{} ({})'.format(pdot, cdot)
            comment = 'The {} mutation in the {} gene'.format(mutation, gene)
            if pdot.endswith('fs'):
                comment += ' results in a frameshift'
            elif pdot.endswith('Ter'):
                comment += ' results in a premature termination codon'
                p = aa_change_names_and_codon(row[i_aa])
                if p:
                    comment += ' at amino acid position {}'.format(p[1])
            elif is_substitution_bases(row[i_cds]):
                p = aa_change_names_and_codon(row[i_aa])
#                sys.stderr.write("substitution p {} -> {}\n".format(row[i_aa], p))
                if p:
                    n = 'n' if p[2][0] in ('a','i') else ''
                    comment += ' results in a substitution of '+\
                      'a{} {} for the wild-type {} at codon {}'.format(n,
                      p[2], p[0], p[1])
#            sys.stderr.write(comment+'\n')
            row.append(comment)
    return tabdata

def create_variant_report_xlsx(report, args):
    outfile = outfile_name(report, args.outdir, '.xlsx')
    if args.debug:
        sys.stderr.write("    Writing {}\n".format(outfile))
    sheetname = os.path.basename(outfile).replace('.xlsx','')
    fields = None
    highlight_row = ExcelRowData(['']*26, 'gold')
    tabdata = parse_tab_file(report, outfile=outfile)
    tabdata = add_comment_snippet(tabdata)
    header = [ ExcelRowData([l,]) for l in tabdata.header ]
    header.append(ExcelRowData(tabdata.fields))
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
            data.append(ExcelRowData(row))#, 'red', i_status))
        elif row[i_status]=='ACCEPT':
            data.append(ExcelRowData(row))#, 'green', i_status))
        else:
            data.append(ExcelRowData(row))
    numxlines = print_spreadsheet_excel(header, data, outfile, sheetname)
    if highlight_row:
        num_expect = tabdata.numlines
        sys.stderr.write("    No NOT_REPORTED variants\n")
    else:
        num_expect = tabdata.numlines+1
    if num_expect != numxlines:
        sys.stderr.write("    {} lines in report\n".format(tabdata.numlines))
        sys.stderr.write("    {} lines in spreadsheet\n".format(numxlines))
        sys.exit("  ERROR: Unexpected num lines\n")
    return tabdata

def pos_sortkey(chrom, pos):
    chrom = "%02d" % int(chrom) if chrom.isdigit() else "%-2s" % chrom
    return "%s.%011d" % (chrom, pos)

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
    vcfaccept = defaultdict(list)
    vcfreject = defaultdict(list)
    with open(vcffile, 'r') as fh:
        for line in fh:
            if line.startswith('#'):
                vcfhead.append(line)
            else:
                row = line.split("\t", 3)
                chrom = row[0]
                pos = int(row[1])
                sortkey = pos_sortkey(chrom, pos)
                if variantdata[chrom][pos]=='NOT_REPORTED':
                    vcfreject[sortkey].append(line)
                else:
                    vcfaccept[sortkey].append(line)
    with open(acceptfile, 'w') as ofh:
        ofh.write(''.join(vcfhead))
        ofh.write(''.join([ ''.join(vcfaccept[k]) for k in \
                  sorted(vcfaccept)]))
    sys.stderr.write("    Num accepted:{:4d}\n".format(len(vcfaccept)))
    with open(rejectfile, 'w') as ofh:
        ofh.write(''.join(vcfhead))
        ofh.write(''.join([ ''.join(vcfreject[k]) for k in \
                  sorted(vcfreject)]))
    sys.stderr.write("    Num rejected:{:4d}\n".format(len(vcfreject)))
    return acceptfile, rejectfile

def read_transcript_file(transcriptfile):
    """Input should be tab-delimited file with two columns: gene and
    transcript.  Return dict keyed by gene with value transcript."""
    sys.stderr.write("Reading {}\n".format(transcriptfile))
    if not os.path.isfile(transcriptfile):
        sys.exit("  Fusion transcript file {} not found".format(
                 transcriptfile))
    transcripts = {}
    with open(transcriptfile, 'r') as fh:
        for l in fh:
            if not l.startswith('#'):
                data = l.rstrip().split("\t")
                if len(data)>1:
                    transcripts[data[0]] = data[1]
#        transcripts = dict(zip([l.rstrip().split("\t")[:2] for l in fh \
#                      if "\t" in l.rstrip() and not l.startswith('#')]))
    return transcripts

def add_transcripts_to_fusion_report(fusionfile, args):
    fields = []
    lines = []
    with open(fusionfile, 'r') as fh:
        fields = fh.readline().split("\t")
        lines = [ l for l in fh.readlines() if "\t" in l.rstrip() ]
    if not lines: 
        sys.stderr.write("    No fusions\n")
        return 0
    else:
        sys.stderr.write("    {} fusions\n".format(len(lines)))
    newfile = fusionfile.replace("filtered.txt",'')+'with_transcripts.txt'
    if args.outdir:
        newfile = os.path.join(args.outdir, os.path.basename(newfile))
    i_region1 = None
    i_region2 = None
    try:
        i_region1 = fields.index('Region1')
        i_region2 = fields.index('Region2')
    except:
        sys.exit("Bad format file {}".format(fusionfile))
    i = i_region2 + 1 # insert transcripts after Region2
    fields[i:0] = ['Transcript1', 'Transcript2']
    with open(newfile, 'w') as ofh:
        ofh.write("\t".join(fields))
        for line in lines:
            data = line.split("\t")
            t1 = FUSION_TRANSCRIPTS.get(data[i_region1], '')
            t2 = FUSION_TRANSCRIPTS.get(data[i_region2], '')
            data[i:0] = [t1, t2]
            ofh.write("\t".join(data))
    return newfile

def group_files_by_sample(inputfiles):
    extensions = {
        '.depth_report_indels.txt': 'dp_indels', 
        '.depth_report_snvs.txt': 'dp_snvs',
        '.fusions.filtered.txt': 'fusions',
        '.variant_report.txt': 'v_report', 
        '.vcf': 'vcf', }
    samples = defaultdict(dict)
    infiles = []
    for in_arg in inputfiles: # input can be files or folders
        if os.path.isfile(in_arg):
            infiles.append(in_arg)
        elif os.path.isdir(in_arg):
            infiles.extend([ os.path.join(in_arg, f) for f in \
                             os.listdir(in_arg) ])
    badfiles = []
    for infile in infiles:
        if infile.endswith('_accepted.vcf') or \
           infile.endswith('_rejected.vcf') or \
           infile.endswith('.unfiltered.vcf'):
            badfiles.append(infile)
            continue
        notfound = True
        for ext in extensions:
            if infile.endswith(ext):
                sample = os.path.basename(infile).replace(ext,'')
                samples[sample][extensions[ext]] = infile
                notfound = False
                break
        if notfound:
            badfiles.append(infile)
    return samples, badfiles

#----gui.py-------------------------------------------------------------------

class StampPostProcess_App(wx.App):
    def __init__(self, args, **kwargs):
        self.args = args
        wx.App.__init__(self, kwargs)

    def OnInit(self):
        self.frame = StampFrame(self.args)
        self.frame.Show()
        self.SetTopWindow(self.frame)
        return True

class StampRTC(wx.richtext.RichTextCtrl):
    def __init__(self, parent):
        wx.richtext.RichTextCtrl.__init__(self, parent, -1, "",
                        style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
        self.Bind(wx.EVT_MOUSE_EVENTS, self.DoNothing)

    def DoNothing(self, event):
        pass

    def AddIntroBlurb(self):
        intro_blurb="This script creates a variety of files depending on "+\
                    "the combination of STAMP output files received:"
        intro_items = [     
            ['Excel variant reports', 
             'requires sample.variant_report.txt'],
            ['Accepted and rejected VCF files', 
             'requires sample.variant_report.txt and sample.vcf'],
            ['Sorted Excel depth reports', 
             'requires sample.depth_report_indels.txt and/or ' +\
             ' sample.depth_report_snvs.txt'],
            ['Low coverage comment', 
             'requires sample.depth_report_indels.txt and '+\
             'sample.depth_report_snvs.txt'],
            ['Fusion file with transcripts', 'requires fusions.filtered.txt'],
        ]
        self.BeginFontSize(10)
        self.Newline()
        self.Newline()
        self.WriteText(intro_blurb)
        self.Newline()
        self.BeginSymbolBullet('*', 25, 30)
        for [label, descr] in intro_items:
            self.BeginBold()
            self.WriteText(label)
            self.EndBold()
            self.WriteText(' -- '+descr)
            self.Newline()
        self.EndSymbolBullet()
        self.EndFontSize()
        self.Newline()
        self.Newline()

class StampFrame(wx.Frame):
    def __init__(self, args):
        self.args = args
        wx.Frame.__init__(self, None, size=(550,500),
                          title="STAMP Post-Processing v"+VERSION, )

        panel = wx.Panel(self)
        label = wx.StaticText(panel, -1, "Drop STAMP depth reports, variant "+\
            " reports, and VCF files here:")
#        self.rtc = wx.richtext.RichTextCtrl(panel,-1, "",
#                        style=wx.TE_READONLY|wx.TE_MULTILINE|wx.HSCROLL)
        self.rtc = StampRTC(panel)
        self.rtc.AddIntroBlurb()
        button_quit = wx.Button(panel, -1, "Quit", style=wx.BU_EXACTFIT)
        button_quit.SetToolTip(wx.ToolTip("Quit application"))
        self.Bind(wx.EVT_BUTTON, self.OnCloseMe, button_quit)
        self.Bind(wx.EVT_CLOSE, self.OnCloseWindow)

        sizer = wx.BoxSizer(wx.VERTICAL)
        sizer.Add(label, 0, wx.ALL, 5)
        sizer.Add(self.rtc, 1, wx.EXPAND|wx.ALL, 5)

        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        button_sizer.AddStretchSpacer()
        button_sizer.Add(button_quit, 0, wx.ALIGN_CENTER_VERTICAL)
        sizer.Add(button_sizer, 0, wx.ALL|wx.EXPAND, 5)
        panel.SetSizer(sizer)

        filedrop = FileDropProcessing(self.rtc, self.args)
        self.rtc.SetDropTarget(filedrop)

    def OnCloseMe(self, event):
        self.Close(True)

    def OnCloseWindow(self, event):
        self.Destroy()
        
class FileDropProcessing(wx.FileDropTarget):
    def __init__(self, window, args):
        wx.FileDropTarget.__init__(self)
        self.window = window
        self.args = args
        self.num_samples = 0
        self.oldsamples = {}
        self.current_pos = 0

    def ScrollWindow(self):
        pos = self.window.GetScrollRange(wx.VERTICAL)
        self.window.Scroll(0, pos)

    def WriteFormattedText(self, boldtext='', normaltext='', bullet=False,
                           newline=True):
        self.window.MoveEnd()
        if self.current_pos:
            self.window.SetCaretPosition(self.current_pos)
        if bullet:
            self.window.BeginSymbolBullet('*', 25, 30)
        if boldtext:
            self.window.BeginBold()
            self.window.WriteText(boldtext)
            self.window.EndBold()
        if normaltext:
            self.window.WriteText(normaltext)
        if newline: 
            self.window.Newline()
        if bullet:
            self.window.EndSymbolBullet()
        self.ScrollWindow()
        self.current_pos = self.window.GetCaretPosition()

    def OnDropFiles(self, x, y, filenames):
        counts = {}
        samples, badfiles = group_files_by_sample(filenames)
        if badfiles:
            self.window.MoveEnd()
            for badfile in badfiles:
                self.WriteFormattedText(newline=False,
                    normaltext="Not a recognized input file: {}\n".\
                    format(os.path.basename(badfile)))
                self.ScrollWindow()
        self.WriteFormattedText(newline=True)
        for sample, d in sorted(samples.items()):
            if sample in self.oldsamples:
                old_d = self.oldsamples[sample]
                d['sample_num'] = old_d['sample_num']
                for filetype in ('v_report', 'vcf', 'dp_indels', 'dp_snvs'):
                    if filetype not in d and filetype in old_d:
                        d[filetype] = old_d[filetype]
            else:
                self.num_samples += 1
                d['sample_num'] = self.num_samples
            self.oldsamples[sample] = d
            sys.stderr.write('Sample {}:  {}\n'.format(d['sample_num'], 
                             sample))
            self.WriteFormattedText('Sample {}:  '.format(d['sample_num']),
                                    sample)
            try:
                dpindelinfo = None
                dpsnvinfo = None
                vinfo = None
                if 'v_report' in d:
                    self.WriteFormattedText("Variant report:  ", 
                        os.path.basename(d['v_report']), True)
                    sys.stderr.write("- Formatting variant report\n")
                    vinfo = create_variant_report_xlsx(d['v_report'], self.args)
                    if os.path.isfile(vinfo.outfile):
                        self.WriteFormattedText("","      --Wrote {}".format(
                            os.path.basename(vinfo.outfile)))
                    sys.stderr.flush()
                if 'vcf' in d:
                    self.WriteFormattedText("VCF file:  ", 
                        os.path.basename(d['vcf']), True)
                    if vinfo:
                        sys.stderr.write("- Splitting vcf\n")
                        outfiles = split_vcf(d['vcf'], vinfo, self.args)
                        for outfile in outfiles:
                            if os.path.isfile(outfile):
                                self.WriteFormattedText("",
                                    "      --Wrote {}".format(
                                    os.path.basename(outfile)))
                    sys.stderr.flush()
                if 'dp_indels' in d:
                    self.WriteFormattedText("Indel depth file:  ", 
                        os.path.basename(d['dp_indels']), True)
                    sys.stderr.write("- Sorting indel depth report\n")
                    dpindelinfo = create_depth_report_xlsx(d['dp_indels'], 
                                                           self.args)
                    if os.path.isfile(dpindelinfo.outfile):
                        self.WriteFormattedText("","      --Wrote {}".format(
                            os.path.basename(dpindelinfo.outfile)))
                    sys.stderr.flush()
                if 'dp_snvs' in d:
                    self.WriteFormattedText("SNV depth file:  ", 
                        os.path.basename(d['dp_snvs']), True)
                    sys.stderr.write("- Sorting snv depth report\n")
                    dpsnvinfo = create_depth_report_xlsx(d['dp_snvs'], self.args)
                    if os.path.isfile(dpsnvinfo.outfile):
                        self.WriteFormattedText("","      --Wrote {}".format(
                            os.path.basename(dpsnvinfo.outfile)))
                    sys.stderr.flush()
                if dpindelinfo and dpsnvinfo:
                    sys.stderr.write("- Generating low coverage comment\n")
                    outlabel = outfile_name(dpsnvinfo.tabfile, args.outdir, 
                                            inext='.depth_report_snvs.txt')
                    outfile, is_female = generate_low_coverage_comment(
                                   outlabel, dpindelinfo, dpsnvinfo)
                    if os.path.isfile(outfile):
                        gender = '(F)' if is_female else '(M)'
                        self.WriteFormattedText("",
                            "      --Wrote {} {}".format(
                            os.path.basename(outfile), gender))
                    sys.stderr.flush()
                if 'fusions' in d:
                    self.WriteFormattedText("Fusions:  ", 
                        os.path.basename(d['fusions']), True)
                    sys.stderr.write("- Adding transcripts to fusion file\n")
                    newfusionfile = add_transcripts_to_fusion_report(
                                         d['fusions'], self.args)
                    if newfusionfile==0:
                        self.WriteFormattedText("","      --No fusions")
                    elif newfusionfile and os.path.isfile(newfusionfile):
                        self.WriteFormattedText("","      --Wrote {}".format(
                            os.path.basename(newfusionfile)))
                    sys.stderr.flush()
            except Exception, e:
                self.window.WriteText("    ERROR: {} {}\n\n".format(
                                       type(e).__name__, e))
                raise
            self.WriteFormattedText(newline=True)

def run_gui(args):
    app = StampPostProcess_App(args)
    app.MainLoop()

#-----------------------------------------------------------------------------
if __name__=='__main__':
    descr = "This script post-processes STAMP report files."
    descr += " Depth reports will be sorted by Min depth with values less"
    descr += " than 200 highlighted and saved as Excel."
    descr += " Variant reports will have a highlighted row inserted between"
    descr += " ACCEPTED and NOT_REPORTED variants and saved as Excel."
    descr += " If VCF and variant reports are available, the VCF will be"
    descr += " split into accepted and rejected VCF files."
    descr += " A file with low coverage comment is generated if both"
    descr += " indel and SNV depth reports are input."
    descr += " Transcripts are added any fusions.filtered.txt files."
    parser = ArgumentParser(description=descr)
    parser.add_argument("reports", nargs="*",
                        help="STAMP depth and/or variant report(s)")
    parser.add_argument("-o", "--outdir", 
                        help="Directory to save output file(s)")
    parser.add_argument("-t", "--transcripts", default=FUSION_TRANSCRIPT_FILE,
                        help="Fusion transcript file")
    parser.add_argument("--debug", default=False, action='store_true',
                        help="Write debugging messages")

    args = parser.parse_args()
    FUSION_TRANSCRIPTS = read_transcript_file(args.transcripts)
    if len(args.reports)==0:
        run_gui(args)
    else:
        samples, badfiles = group_files_by_sample(args.reports)
        for sample, d in sorted(samples.items()):
            sys.stderr.write("\nSample {}\n".format(sample))
            dpindelinfo = None
            dpsnvinfo = None
            vinfo = None
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
            sys.stderr.write("- Sorting indel depth report: ")
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
            sys.stderr.write("- Generating low coverage comment: ")
            if dpindelinfo and dpsnvinfo:
                sys.stderr.write(" YES\n")
                outlabel = outfile_name(dpsnvinfo.tabfile, args.outdir, 
                                        inext='.depth_report_snvs.txt')
                lcc = generate_low_coverage_comment(outlabel, dpindelinfo, 
                                                    dpsnvinfo)
            else:
                sys.stderr.write(" NO\n")
            sys.stderr.write("- Adding transcripts to fusion file: ")
            if 'fusions' in d:
                sys.stderr.write(" YES\n")
                numfusions = add_transcripts_to_fusion_report(d['fusions'], args)
            else:
                sys.stderr.write(" NO\n")





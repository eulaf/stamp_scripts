#!/usr/bin/env python

import os
import sys
from collections import defaultdict

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
        for f in ('dbSNP138_ID', 'COSMIC70_ID'):
            if f in d and d[f]=='NA':
                d[f] = None

def count_uppercase(s):
    return sum(1 for c in s if c.isupper())

def field2dbfield(f):
    """Convert field from variant report style to match db column name"""
    newf = f.replace(' ','_')
    if count_uppercase(newf)<3:
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

def parse_variant_file(vfile):
    sys.stderr.write("\nReading variant report: {}\n".format(vfile))
    sample = os.path.basename(vfile).replace('.variant_report.txt', '')
    runnum = sample.lstrip('TruQ3_')
    stamprun = "STAMP{}".format(runnum) if runnum else ''
    vinfo = parse_tab_file(vfile, keyfunc=create_dkey, 
                           fieldfunc=field2dbfield)
    massage_data(vinfo['data'])
    sys.stderr.write("  {} mutations\n".format(len(vinfo['data'])))
    return (vinfo, stamprun, sample)

#-----------------------------------------------------------------------------

def compare_variants(truth, vinfo, counts={}):
    truths_seen = dict([ (dkey, False) for dkey in truth.keys() ])
    summary = { 'Total':0, 'Expected':0, 'Unexpected':0, }
    if not counts:
        counts.update({ 'Total':0, 'Expected':defaultdict(int),
                   'Unexpected':defaultdict(int), })
    counts['Total'] += 1
    for d in vinfo['data']:
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
            d['Expected?'] = 'Not expected'
    notseen = [ dkey for dkey in truths_seen.keys() if not truths_seen[dkey] ]
    summary['Not found'] = len(notseen)
    summary['notseen'] = notseen
    if notseen:
        sys.stderr.write("Not found: "+", ".join(notseen)+"\n")
    else:
        sys.stderr.write("All truths found\n")
    return summary

#-----------------------------------------------------------------------------

def print_checked_file(runinfo, tinfo, outfile):
    tdat = tinfo['datadict']
    vinfo = runinfo['vinfo']
    summary = runinfo['summary']
    fields = vinfo['fields'][:] + ['Expected?',]
    sys.stderr.write("Writing {}\n".format(outfile))
    content = "# Num expected found: {}\n".format(summary['Expected']) +\
              "# Num not expected: {}\n".format(summary['Unexpected']) +\
              "# Num not found: {}\n".format(summary['Not found'])
    sys.stderr.write(content)
    content += "\t".join([field2reportfield(f) for f in fields])+"\n"
    with open(outfile, 'w') as ofh:
        for d in vinfo['data']:
            if d['Expected?']=='Expected':
                dkey = d['dkey']
                if tdat[dkey]['expectedVAF']:
                    d['Expected?'] +=' ({}%)'.format(tdat[dkey]['expectedVAF'])
            row = [ d[f] for f in fields ]
            content += "\t".join([ str(r) for r in row ])+"\n"
        ofh.write(content)
        for dkey in sorted(summary['notseen']):
            d = tdat[dkey].copy()
            d['CDS Change'] = d['HGVS'].replace('c.','')
            d['AA Change'] = d['protein']
            d['Expected?'] = 'Not found'
            if tdat[dkey]['expectedVAF']:
                d['Expected?'] += ' ({})'.format(tdat[dkey]['expectedVAF'])
            row = [ d[f] if f in d else '' for f in fields ]
            ofh.write("\t".join([ str(r) for r in row ]) + "\n")



#!/usr/bin/env python

"""
Check STAMP control for expected variants.
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

def have_file(filename, force=False):
    if os.path.isfile(filename) and force:
        sys.stderr.write("  Removing {}\n".format(filename))
        os.remove(filename)
    return True if os.path.isfile(filename) else False

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

def parse_variant_file(vfile):
    sys.stderr.write("\nReading variant report: {}\n".format(vfile))
    sample = os.path.basename(vfile).replace('.variant_report.txt', '')
    runnum = sample.lstrip('TRUQtruq3_').lstrip('STAMP')
    run = "STAMP{}".format(runnum) if runnum else ''
    vinfo = parse_tab_file(vfile, keyfunc=create_dkey, 
                           fieldfunc=field2dbfield)
    massage_data(vinfo['data'])
    sys.stderr.write("  Run\t{}\n  Sample\t{}\n".format(run, sample))
    sys.stderr.write("  {} mutations\n".format(len(vinfo['data'])))
    return (vinfo, run, sample)

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
    wbformat['bold'] = workbook.add_format({'bold': True})
    wbformat['perc'] = workbook.add_format({'num_format': '#.##%'})
    wbformat['red'] = workbook.add_format({'bg_color': '#C58886', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltred'] = workbook.add_format({'bg_color': '#E9D4D3', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['orange'] = workbook.add_format({'bg_color': '#FCD5B4',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltgreen'] = workbook.add_format({'bg_color': '#EBF1DE',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltblue'] = workbook.add_format({'bg_color': '#D7E1EB',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['blue'] = workbook.add_format({'bg_color': '#88A4C5',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['gray'] = workbook.add_format({'bg_color': '#F0F0F0',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['dkgray'] = workbook.add_format({'bg_color': '#BFBFBF',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['ltbluepatt'] = workbook.add_format({'fg_color': '#DCE6F0', 
                                       'pattern': 8,
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['gray_perc'] = workbook.add_format({'num_format': '#.##%',
                                       'bg_color': '#F0F0F0',
                                       'border': 1, 'border_color':'#CDCDCD'})
    wbformat['dkgray_perc'] = workbook.add_format({'num_format': '#.##%',
                                       'bg_color': '#BFBFBF',
                                       'border': 1, 'border_color':'#CDCDCD'})
    return wbformat

def print_spreadsheet_excel(data, outfile, hiderows=[], fieldfunc=None):
    sys.stderr.write("Writing {}\n".format(outfile))
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()
    wbformat = add_formats_to_workbook(workbook)
    rownum = 0
    # comment lines
    for line in data['header']:
        worksheet.write(rownum, 0, line)
        rownum += 1
    # print column names
    for colnum, f in enumerate(data['fields']):
        colname = fieldfunc(f) if fieldfunc else f
        worksheet.write(rownum, colnum, colname, wbformat['bold'])
    calc_fields = ['AverageVAF', 'StddevVAF', '%Detection']
    i_col_avg = colnum+1
    i_col_std = colnum+2
    for f in calc_fields:
        colnum += 1
        worksheet.write(rownum, colnum, f, wbformat['bold'])
    i_col_run_s = colnum + 1
    for f in data['runs']:
        colnum += 1
        worksheet.write(rownum, colnum, f, wbformat['bold'])
    i_col_run_e = colnum
    runcolxl_s = convert_to_excel_col(i_col_run_s)
    runcolxl_e = convert_to_excel_col(i_col_run_e)
    avgcolxl = convert_to_excel_col(i_col_avg)
    stdcolxl = convert_to_excel_col(i_col_std)
    i_position = data['fields'].index('position')
    i_expectVAF = data['fields'].index('expectedVAF')
    # print data
    numvariants = 0
    for label in ('expected', 'not_expected'):
      percformat = wbformat['perc'] if label=='expected' else \
                   wbformat['gray_perc']
      for dkey in sorted(data[label].keys()):
        numvariants += 1
        rownum += 1
        ddat = [ data[label][dkey][run] if run in data[label][dkey] \
                   else None for run in data['runs'] ] 
        mutdat = [ d for d in ddat if d ]
        for colnum, f in enumerate(data['fields']):
            v = mutdat[0][f] if f in mutdat[0] else ''
            if colnum == i_position: # format as number
                worksheet.write_number(rownum, colnum, v)
            elif colnum == i_expectVAF: # format as number/percent
                if v: 
                    worksheet.write(rownum, colnum, v/100, percformat)
            else:
                worksheet.write(rownum, colnum, v)
        for i, d in enumerate(ddat):
            if d: worksheet.write_number(rownum, colnum+i+4, float(d['vaf']))
        runrange = "{1}{0}:{2}{0}".format(rownum+1, runcolxl_s, runcolxl_e)
        colnum += 1
#        worksheet.write(rownum, colnum, '=AVERAGE({})'.format(runrange))
        worksheet.write_array_formula(rownum, colnum, rownum, colnum,
                        '{'+'=AVERAGE(IF(ISBLANK({0}),0,{0}))'.format(
                        runrange)+'}')
        colnum += 1
        if len(mutdat)>1:
#            worksheet.write(rownum, colnum, '=STDEV({})'.format(runrange))
            worksheet.write_array_formula(rownum, colnum, rownum, colnum,
                        '{'+'=STDEV(IF(ISBLANK({0}),0,{0}))'.format(
                        runrange)+'}')
        colnum += 1
        worksheet.write(rownum, colnum, '=COUNT({})/{}'.format(runrange, 
                        len(data['runs'])), percformat)
        if label=='expected': 
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'between', 
               'minimum':'2*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                        stdcolxl),
               'maximum':'3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                        stdcolxl),
               'format':wbformat['ltred'], })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'>', 
               'value':'3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                      stdcolxl),
               'format':wbformat['red'], })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'between', 
               'minimum':'-2*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                         stdcolxl),
               'maximum':'-3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                         stdcolxl),
               'format':wbformat['ltblue'], })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'<', 
               'value':'-3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                       stdcolxl),
               'format':wbformat['blue'], })
            worksheet.conditional_format(runrange, {'type':'blanks', 
                                         'format':wbformat['ltbluepatt'], })
        else: 
#            wholerow = "A{0}:{1}{0}".format(rownum+1, runcolxl_e)
            worksheet.set_row(rownum, None, wbformat['gray'])
    worksheet.set_column(i_col_run_s-1, i_col_run_e-1, 10) #set col width
    worksheet.set_column(i_position, i_position, 9)
    for i in hiderows:
        worksheet.set_row(i, None, None, {'hidden': True})
        worksheet.set_row(i, None, None, {'hidden': True})
    worksheet.freeze_panes(5, 0)#, 0, 0)
    workbook.close()
    wb = openpyxl.load_workbook(outfile)
    numruns = len(data['runs'])
    wb.save(outfile)
    return {'num_runs':numruns, 'num_variants':numvariants}

def generate_excel_spreadsheet(dbh, tfields, outfile):
    allvafs = get_vafs_for_all_runs(dbh.cursor())
    failed_runs = {}
    good_runs = {}
    data = { 'expected':defaultdict(dict), 'not_expected':defaultdict(dict) }
    for d in allvafs:
        if d['run_status']=='FAIL':
            failed_runs[d['run_name']] = 1
        else:
            good_runs[d['run_name']] = 1
            if d['is_expected']:
                vdict = data['expected']
                dkey = int(d['mutation_id'])
            else:
                vdict = data['not_expected']
                dkey = create_dkey(d)
            vdict[dkey][d['run_name']] = d
    header = [ "# This spreadsheet is automatically generated." +\
               " Any edits will be lost in future versions.",
               "# Failed runs (not in spreadsheet): {}".format(
               ", ".join(sorted(failed_runs.keys())) if failed_runs else 0),
               "# Num runs in spreadsheet: {}".format(len(good_runs.keys())), 
               "# Num expected variants: {}".format(len(data['expected'])), ]
    fields = tfields[:]
    fields.remove('dbSNP')
    fields.remove('COSMIC')
    fields.remove('ref')
    fields.remove('var')
    if 'is_expected' in fields: fields.remove('is_expected')
    data['header'] = header
    data['fields'] = fields
    data['runs'] = sorted(good_runs.keys(), reverse=True)
    nums = print_spreadsheet_excel(data, outfile, hiderows=[0,1],
                                        fieldfunc=field2reportfield) 
    nums['failedruns'] = failed_runs
    return nums

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
            print_checked_file(info, self.tinfo, outfile)
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

if __name__=='__main__':
    descr = "Checks STAMP TruQ3 variant report(s) for expected variants."
    descr += " Creates new annotated variant report(s) in the same"
    descr += " directory unless otherwise specified."
    parser = ArgumentParser(description=descr)
    parser.add_argument("variant_file", nargs="*",
                        help="STAMP TruQ3 variant report(s)")
    parser.add_argument("-o", "--outdir", 
                        help="Directory to save output file(s)")
    parser.add_argument("-s", "--status", default='PASS',
                        help="Status to use for all reports (default: PASS)")
    parser.add_argument("-t", "--text", default=False, action='store_true',
                        help="Print checked variant reports.")
    parser.add_argument("-x", "--excel", default=False, action='store_true',
                        help="Print Excel spreadsheet summarizing all data.")
    parser.add_argument("-d", "--debug", default=False, action='store_true',
                        help="Print extra messages")
    parser.add_argument("--db", default=REFS['DBFILE'],
                        help="SQLite db file.")
    parser.add_argument("--ref", default=REFS['TRUTHFILE'],
                        help="Expected truths file.")
    parser.add_argument("--safe", default=True, action='store_false',
                        dest="force",
                        help="Do not overwrite existing data in db.")
    parser.add_argument("--schema", default=REFS['SCHEMAFILE'],
                        help="SQLite db schema file.")
    parser.add_argument("--spreadsheet", default=REFS['SPREADSHEET'],
                        help="Full path name for Excel spreadsheet.")

    args = parser.parse_args()
    (dbh, tinfo, msgs) = check_db(args.db, args.schema, args.ref)
    if len(args.variant_file)==0:
        msg = ''.join(msgs)
        run_gui(dbh, tinfo, args.db, args.spreadsheet, msg=''.join(msgs))
    else:
        counts = { 'Total':0, 'Expected':defaultdict(int),
                   'Unexpected':defaultdict(int), }
        runs = {}
        for variant_file in args.variant_file:
            if variant_file=='none': continue
            (vinfo, run, sample) = parse_variant_file(variant_file)
            summary = compare_variants(tinfo['datadict'], vinfo, counts,
                                       args.status)
            save2db(dbh, run, sample, summary['Status'], vinfo, args.force)
            runs[run] = {'vinfo':vinfo, 'summary':summary }
            if args.text:
                outfile = variant_file.replace('.txt','') + ".checked.txt"
                if args.outdir:
                    outfile = os.path.join(args.outdir, 
                                           os.path.basename(outfile))
                if have_file(outfile, force=True):
                    sys.stderr.write("  Already have {}.\n".format(outfile))
                else:
                    print_checked_file(runs[run], tinfo, outfile)
        if args.excel:
            generate_excel_spreadsheet(dbh, tinfo['fields'], args.spreadsheet)
        dbh.close()
        if args.debug:
            sys.stderr.write("Total reports:\t{}\n".format(counts['Total']))
            for status in ('Expected', 'Unexpected'):
                sys.stderr.write("\n"+ status + " variants:\n")
                for dkey in sorted(counts[status]):
                    sys.stderr.write("  {:30s}:\t{:3d}\n".format(dkey, 
                                     counts[status][dkey]))



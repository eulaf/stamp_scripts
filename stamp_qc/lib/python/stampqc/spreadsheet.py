#!/usr/bin/env python

import os
import sys
import xlsxwriter
import openpyxl
from collections import defaultdict

import dbops
import fileops

def convert_to_excel_col(colnum):
    mod = colnum % 26
    let = chr(mod+65)
    if colnum > 26:
        rep = colnum/26
        let1 = chr(rep+64)
        let = let1 + let
    return let

def print_spreadsheet_excel(data, outfile, hiderows=[], fieldfunc=None):
    sys.stderr.write("Writing {}\n".format(outfile))
    workbook = xlsxwriter.Workbook(outfile)
    worksheet = workbook.add_worksheet()
    bold_format = workbook.add_format({'bold': True})
    perc_format = workbook.add_format({'num_format': '#.##%'})
    red_format = workbook.add_format({'bg_color': '#C58886', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    ltred_format = workbook.add_format({'bg_color': '#E9D4D3', 
                                       'border': 1, 'border_color':'#CDCDCD'})
    orange_format = workbook.add_format({'bg_color': '#FCD5B4',
                                       'border': 1, 'border_color':'#CDCDCD'})
    ltblue_format = workbook.add_format({'bg_color': '#D7E1EB',
                                       'border': 1, 'border_color':'#CDCDCD'})
    blue_format = workbook.add_format({'bg_color': '#88A4C5',
                                       'border': 1, 'border_color':'#CDCDCD'})
    gray_format = workbook.add_format({'bg_color': '#F0F0F0',
                                       'border': 1, 'border_color':'#CDCDCD'})
    ltbluepatt_format = workbook.add_format({'fg_color': '#DCE6F0', 
                                       'pattern': 8,
                                       'border': 1, 'border_color':'#CDCDCD'})
    gray_perc_format = workbook.add_format({'num_format': '#.##%',
                                       'bg_color': '#F0F0F0',
                                       'border': 1, 'border_color':'#CDCDCD'})
    rownum = 0
    # comment lines
    for line in data['header']:
        worksheet.write(rownum, 0, line)
        rownum += 1
    # print column names
    for colnum, f in enumerate(data['fields']):
        colname = fieldfunc(f) if fieldfunc else f
        worksheet.write(rownum, colnum, colname, bold_format)
    calc_fields = ['AverageVAF', 'StddevVAF', '%Detection']
    i_col_avg = colnum+1
    i_col_std = colnum+2
    for f in calc_fields:
        colnum += 1
        worksheet.write(rownum, colnum, f, bold_format)
    i_col_run_s = colnum + 1
    for f in data['runs']:
        colnum += 1
        worksheet.write(rownum, colnum, f, bold_format)
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
      percformat = perc_format if label=='expected' else gray_perc_format
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
        worksheet.write(rownum, colnum, '=AVERAGE({})'.format(runrange))
        colnum += 1
        if len(mutdat)>1:
            worksheet.write(rownum, colnum, '=STDEV({})'.format(runrange))
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
               'format':ltred_format, })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'>', 
               'value':'3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                      stdcolxl),
               'format':red_format, })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'between', 
               'minimum':'-2*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                         stdcolxl),
               'maximum':'-3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                         stdcolxl),
               'format':ltblue_format, })
            worksheet.conditional_format(runrange, 
              {'type':'cell', 'criteria':'<', 
               'value':'-3*${2}${0} + ${1}${0}'.format(rownum+1, avgcolxl, 
                                                       stdcolxl),
               'format':blue_format, })
            worksheet.conditional_format(runrange, {'type':'blanks', 
                                         'format':ltbluepatt_format, })
        else: 
#            wholerow = "A{0}:{1}{0}".format(rownum+1, runcolxl_e)
            worksheet.set_row(rownum, None, gray_format)
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
    allvafs = dbops.get_vafs_for_all_runs(dbh.cursor())
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
                dkey = fileops.create_dkey(d)
            vdict[dkey][d['run_name']] = d
    header = [ "# This spreadsheet is automatically generated." +\
               " Any edits will be lost in future versions.",
               "# Failed runs (not in spreadsheet): {}".format(
               ", ".join(sorted(failed_runs.keys()))),
               "# Num runs in spreadsheet: {}".format(len(good_runs.keys())), 
               "# Num expected variants: {}".format(len(data['expected'])), ]
    fields = tfields[:]
    fields.remove('dbSNP138_ID')
    fields.remove('COSMIC70_ID')
    fields.remove('ref')
    fields.remove('var')
    if 'is_expected' in fields: fields.remove('is_expected')
    data['header'] = header
    data['fields'] = fields
    data['runs'] = sorted(good_runs.keys(), reverse=True)
    nums = print_spreadsheet_excel(data, outfile, hiderows=[0,1],
                                        fieldfunc=fileops.field2reportfield) 
    nums['failedruns'] = failed_runs
    return nums


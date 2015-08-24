#!/usr/bin/env python

"""
Check STAMP control for expected variants.
"""

import os
import sys
import sqlite3
from collections import defaultdict
from argparse import ArgumentParser

def getScriptPath():
    return os.path.dirname(os.path.realpath(sys.argv[0]))
lib_dir = os.path.join(getScriptPath(), '..', 'lib', 'python')
sys.path.insert(0, lib_dir)

import stampqc.fileops as fileops 
import stampqc.common as common
import stampqc.dbops as dbops
import stampqc.spreadsheet as spreadsheet
import stampqc.gui as gui

def run_gui(dbh, tinfo, db, spreadsheet, msg):
    app = gui.StampQC_App(dbh, tinfo, db, spreadsheet, msg=msg)
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
    parser.add_argument("-t", "--text", default=False, action='store_true',
                        help="Print checked variant reports.")
    parser.add_argument("-x", "--excel", default=False, action='store_true',
                        help="Print Excel spreadsheet summarizing all data.")
    parser.add_argument("-d", "--debug", default=False, action='store_true',
                        help="Print extra messages")
    parser.add_argument("-f", "--force", default=False, action='store_true',
                        help="Overwrite existing data in db.")
    parser.add_argument("--db", default=common.REFS['DBFILE'],
                        help="SQLite db file.")
    parser.add_argument("--ref", default=common.REFS['TRUTHFILE'],
                        help="Expected truths file.")
    parser.add_argument("--schema", default=common.REFS['SCHEMAFILE'],
                        help="SQLite db schema file.")
    parser.add_argument("--spreadsheet", default=common.REFS['SPREADSHEET'],
                        help="Full path name for Excel spreadsheet.")

    args = parser.parse_args()
    (dbh, tinfo, msgs) = dbops.check_db(args.db, args.schema, args.ref)
    if len(args.variant_file)==0:
        msg = ''.join(msgs)
        run_gui(dbh, tinfo, args.db, args.spreadsheet, msg=''.join(msgs))
    else:
        counts = { 'Total':0, 'Expected':defaultdict(int),
                   'Unexpected':defaultdict(int), }
        runs = {}
        for variant_file in args.variant_file:
            if variant_file=='none': continue
            (vinfo, stamprun, sample) = fileops.parse_variant_file(
                                        variant_file)
            summary = fileops.compare_variants(tinfo['datadict'], vinfo, 
                                              counts)
            dbops.save2db(dbh, stamprun, sample, vinfo, args.force)
            runs[stamprun] = {'vinfo':vinfo, 'summary':summary }
            if args.text:
                outfile = variant_file.replace('.txt','') + ".checked.txt"
                if args.outdir:
                    outfile = os.path.join(args.outdir, 
                                           os.path.basename(outfile))
                if common.have_file(outfile, force=True):
                    sys.stderr.write("  Already have {}.\n".format(outfile))
                else:
                    fileops.print_checked_file(runs[stamprun], tinfo, outfile)
        if args.excel:
            spreadsheet.generate_excel_spreadsheet(dbh, tinfo['fields'], 
                                                   args.spreadsheet)
        dbh.close()
        if args.debug:
            sys.stderr.write("Total reports:\t{}\n".format(counts['Total']))
            for status in ('Expected', 'Unexpected'):
                sys.stderr.write("\n"+ status + " variants:\n")
                for dkey in sorted(counts[status]):
                    sys.stderr.write("  {:30s}:\t{:3d}\n".format(dkey, 
                                     counts[status][dkey]))



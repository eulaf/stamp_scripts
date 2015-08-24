#!/usr/bin/env python

"""
Download a table from SQLite3 database.
"""

import os
import sys
import sqlite3
from argparse import ArgumentParser

def download_data(dbfile, dbtable, delim):
    sys.stderr.write("\nConnecting to db {}\n".format(dbfile))
    dbh = sqlite3.connect(dbfile)
    cursor = dbh.cursor()
    cursor.execute('SELECT * FROM {}'.format(dbtable))
    print delim.join([ d[0] for d in cursor.description ])
    for row in cursor.fetchall():
        print delim.join([str(r) if r is not None else '' for r in row ])


if __name__=='__main__':
    descr = "Download a table from a SQLite3 database"
    descr += " and print to STDOUT in tab-delimited format."
    parser = ArgumentParser(description=descr)
    parser.add_argument("db", help="SQLite3 database file")
    parser.add_argument("table", help="Name of database table to download")
    parser.add_argument("-d", "--debug", default=False, action='store_true',
                        help="Print extra messages")
    parser.add_argument("-f", "--force", default=False, action='store_true',
                        help="Overwrite existing data.")
    if len(sys.argv)<2:
        parser.print_help()
    else:
        args = parser.parse_args()
        download_data(args.db, args.table, "\t")



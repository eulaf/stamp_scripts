#!/usr/bin/env python

import os
import sys

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


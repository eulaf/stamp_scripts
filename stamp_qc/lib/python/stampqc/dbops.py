#!/usr/bin/env python

import os
import sys
import datetime
import sqlite3

from fileops import create_dkey, parse_truth

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
              'ref', 'var', 'dbSNP138_ID', 'COSMIC70_ID', 'HGVS', 'protein', 
              'whitelist', 'expectedVAF']
    ins_sql = 'INSERT INTO mutation VALUES (?,?' + ',?'*len(fields) + ')'
    for row in data:
        mut = [ row[f] if f in row and len(str(row[f]))>0 else None \
                for f in fields ]
        exp_val = row['is_expected'] if 'is_expected' in row else is_expected
        mut.append(exp_val)
        mut.append(current_time())
        cursor.execute(ins_sql, mut)

def save_run(cursor, run_name, sample_name):
#    m = re.search("\d+$", stamprun)
#    num = m.group(0)
    cursor.execute("INSERT INTO run (run_name, sample_name, last_modified)"+\
                   " VALUES (?,?,?)", (run_name, sample_name, current_time()))

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
    tot_expected = get_num_expected_mutations(cursor)
    limit = int(tot_expected/2)
    cmd = "UPDATE run SET run_status=?," +\
          " num_mutations=(SELECT COUNT(*) FROM vaf WHERE run_id=?)," +\
          " num_expected=(SELECT COUNT(*) FROM vaf v JOIN mutation m" +\
          " ON m.id=v.mutation_id WHERE m.is_expected=? AND v.run_id=?)" +\
          " WHERE run.id=?"
    cursor.execute(cmd, ['PASS', run_id, 1, run_id, run_id])
    cmd = "UPDATE run SET run_status = ? WHERE run.id=?" +\
          " AND run.num_mutations < {}".format(limit)
    cursor.execute(cmd, ['FAIL', run_id])

def delete_vafs_for_run(cursor, run_id):
    cursor.execute("DELETE FROM vaf WHERE run_id=?", (run_id,))

def get_vafs_for_run(cursor, run_id):
    cmd = "SELECT * FROM vaf AS v JOIN mutation AS m ON v.mutation_id=m.id" +\
          " WHERE v.run_id = ?"
    cursor.execute(cmd, [run_id,])
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

def save2db(dbh, stamprun, sample, vinfo, force=False):
    cursor = dbh.cursor()
    run = get_run(cursor, stamprun, sample)
    if not run or force:
        if run and force:
            sys.stderr.write('  Deleting old data for {}:{} in db.\n'.format(
                             stamprun, sample))
            delete_vafs_for_run(cursor, run['id'])
        elif not run:
            sys.stderr.write('  Saving run {}:{} in db.\n'.format(stamprun, 
                             sample))
            save_run(cursor, stamprun, sample)
            run = get_run(cursor, stamprun, sample)
        for d in vinfo['data']:
            save_vaf(cursor, run['id'], d)
        update_run_counts(cursor, run['id'])
        dbh.commit()
    vafs = get_vafs_for_run(cursor, run['id'])
    sys.stderr.write('  Have {} mutations for run {}:{} in db.\n'.format(
                     len(vafs), stamprun, sample))




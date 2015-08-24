create table mutation(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    gene    TEXT,
    chr    TEXT,
    position    INTEGER,
    strand    TEXT,
    ref_transcript    TEXT,
    ref    TEXT,
    var    TEXT,
	dbSNP138_ID    TEXT,
	COSMIC70_ID    TEXT,
    HGVS    TEXT,
    protein    TEXT,
    whitelist    TEXT,
    expectedVAF    REAL,
    is_expected    INTEGER,
	last_modified    TIMESTAMP
);

create unique index mutation_index on mutation 
    (gene, position, ref, var);

create table run(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    run_name    TEXT,
    sample_name    TEXT,
    num_mutations    INTEGER,
    num_expected    INTEGER,
    run_status    TEXT,
	last_modified    TIMESTAMP
);

create unique index run_index on run (run_name, sample_name);

create table vaf(
    run_id    INTEGER,
    mutation_id    INTEGER,
    vaf    REAL,
    vaf_status    TEXT,
	last_modified    TIMESTAMP,
    PRIMARY KEY(run_id, mutation_id)
);

create view total_count as select
    (select count(*) from run) as tot_runs,
    (select count(*) from run where run_status!='FAIL') as tot_good_runs,
    (select count(*) from mutation) as tot_mutations,
    (select count(*) from mutation where is_expected=1) as tot_expected
	from mutation where id=1;

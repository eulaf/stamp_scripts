create table mutation(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    gene    TEXT,
    chr    TEXT,
    position    INTEGER,
    ref_transcript    TEXT,
    ref    TEXT,
    var    TEXT,
    HGVS    TEXT,
    protein    TEXT,
    whitelist    TEXT,
    HorizonVAF    REAL,
    is_expected    INTEGER,
	last_modified    TIMESTAMP
);

create unique index mutation_index on mutation 
    (gene, position, ref, var);

create table fusion (
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT, 
	region1    TEXT, 
	region2    TEXT, 
	break1    TEXT, 
	break2    TEXT, 
	HorizonVAF    REAL, 
	is_expected    INTEGER, 
	last_modified    TIMESTAMP
);

create unique index fusion_index on fusion 
    (region1, region2, break1, break2);

create table sample(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    sample_name    TEXT,
    run_name    TEXT,
    num_mutations    INTEGER,
    num_mutations_missing    INTEGER,
    num_mutations_unexpected    INTEGER,
    num_fusions    INTEGER,
    num_fusions_missing    INTEGER,
    num_fusions_unexpected    INTEGER,
    sample_status    TEXT,
	last_modified    TIMESTAMP
);

create unique index sample_index on sample (run_name, sample_name);

create table sample_mutation(
    sample_id    INTEGER,
    mutation_id    INTEGER,
    vaf    REAL,
    vaf_status    TEXT,
	last_modified    TIMESTAMP,
    PRIMARY KEY(sample_id, mutation_id)
);

create table sample_fusion(
    sample_id    INTEGER,
    fusion_id    INTEGER,
	last_modified    TIMESTAMP,
    PRIMARY KEY(sample_id, fusion_id)
);


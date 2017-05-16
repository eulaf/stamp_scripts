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

create table cnv (
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT, 
	gene    TEXT, 
	locus    TEXT, 
	HorizonCopies    REAL, 
	is_expected    INTEGER, 
	last_modified    TIMESTAMP
);

create unique index cnv_index on cnv (gene);

create table sample(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    sample_name    TEXT,
    run_name    TEXT,
    num_mutations    INTEGER,
    num_mutations_missing    INTEGER,
    num_mutations_other    INTEGER,
    num_fusions    INTEGER,
    num_fusions_missing    INTEGER,
    num_fusions_other    INTEGER,
    num_cnvs    INTEGER,
    num_cnvs_missing    INTEGER,
    num_cnvs_other    INTEGER,
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

create table sample_cnv(
    sample_id    INTEGER,
    cnv_id    INTEGER,
    mean_z    REAL,
    mcopies    REAL,
    status    TEXT,
	last_modified    TIMESTAMP,
    PRIMARY KEY(sample_id, cnv_id)
);


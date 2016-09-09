create table barcode(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    barcode    TEXT,
	last_modified    TIMESTAMP
);

create unique index barcode_index on barcode(barcode);

create table run(
    id    INTEGER    PRIMARY KEY    AUTOINCREMENT,
    run_name    TEXT,
    total_reads    INTEGER,
    run_status    TEXT,
	last_modified    TIMESTAMP
);

create unique index run_index on run (run_name);

create table barcode_counts(
    run_id    INTEGER,
    barcode_id    INTEGER,
    bc_count    INTEGER,
	last_modified    TIMESTAMP,
    PRIMARY KEY(run_id, barcode_id)
);


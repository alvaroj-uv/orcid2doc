create table if not exists master_doi
(
    doi         TEXT,
    autores     TEXT,
    autor_prin  TEXT,
    autor_sec   TEXT,
    anno        INTEGER,
    titulo      TEXT,
    revista     TEXT,
    ref_revista TEXT,
    isbn        TEXT,
    factor      TEXT,
    json        TEXT
);


create table WOS
(
    Journal_Name     TEXT,
    ISSN             TEXT,
    EISSN            TEXT,
    Category_Q       TEXT,
    Citations        INTEGER,
    IF_2022          REAL,
    JCI              REAL,
    percentageOAGold REAL,
    JIF_Quartile     TEXT
);

create index idx_eissn
    on WOS (EISSN);

create index idx_issn
    on WOS (ISSN);
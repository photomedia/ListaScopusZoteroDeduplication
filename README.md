# Deduplication scripts

Python scripts used to deduplicate records exported from
**Scopus**, **Zotero** and **LISTA** (Library, Information Science & Technology
Abstracts, via EBSCO) for a review of open source research software support in
libraries. They have been only minimally edited for sharing (for example,
reviewer names replaced with placeholders) and contain hard-coded file names
and column names. This repository contains **only the Python scripts**. File names
mentioned below (e.g. `lista-345.xlsx`) are the names the scripts expect or
produce.

## Contents

| Script | Purpose | Stage |
|---|---|---|
| `duplicates.py` | Find and remove duplicates across Scopus topic exports (earlier version) | 1 (superseded) |
| `dedupe-lista.py` | Find and remove duplicates across topic exports; adapted for LISTA/EBSCO as well as Scopus | 1 |
| `combine345.py` | Combine exports and keep records rated 3–5 by at least one reviewer | 2 |
| `duplicates_2.py` | Keep only LISTA records that are not already in the Scopus/Zotero set | 3 |

## Workflow

```
Topic exports, one file per topic, per database
   │
   ▼
1. duplicates.py / dedupe-lista.py
      → dedup/<file>_dedup.(csv|xlsx)   (duplicated articles removed)
      → duplicates.csv                  (one row per duplicated article)
   │
   ▼
   Reviewers rate each record in Excel (columns Reviewer1–Reviewer4)
   │
   ▼
2. combine345.py
      → filtered_rows.xlsx   (saved as e.g. scopus-zotero-345.xlsx, lista-345.xlsx)
   │
   ▼
3. duplicates_2.py
      → new_lista_entries_only.xlsx   (LISTA records not already in Scopus/Zotero)
```

Steps 1–2 were run twice: once for the Scopus topic exports together with a
Zotero CSV export, giving `scopus-zotero-345.xlsx`, and once for the LISTA topic
exports, giving `lista-345.xlsx`. Step 3 compares the two sets.

## Requirements

- Python 3.8 or later
- `pandas` (`combine345.py`, `duplicates_2.py`)
- `openpyxl` (all scripts)
- `xlrd` (only if `combine345.py` is given `.xls` files)

```
pip install pandas openpyxl
```

## Expected input formats

The scripts read the first sheet of each Excel file (or the CSV), with column
names in the first row. These are the columns they use; other columns are carried
through unchanged.

| Source | Columns used | Notes |
|---|---|---|
| Scopus CSV/Excel export | `DOI`, `EID`, `Title`, `Year` | Standard Scopus export column names |
| Zotero CSV export | `DOI`, `Title` | Zotero's standard CSV export already uses these column names; its year column is `Publication Year` |
| LISTA (EBSCO) export | `doi`, `title`, `publicationDate` (e.g. `20210301`), `coverDate` | No EID |
| Rated files (`combine345.py`) | `Reviewer1`, `Reviewer2`, `Reviewer3`, `Reviewer4` | One column per reviewer; rating 1–5, optionally followed by a comment |

**File naming:** topic labels in the duplicates summary are taken from the
file names:
- `duplicates.py` uses the part before `_scopus`, e.g. `copyright_scopus_export….csv` → `copyright`
- `dedupe-lista.py` uses the part before `EBSCO`, e.g. `copyright_EBSCO-Metadata….xlsx` → `copyright_`

---

## `duplicates.py` (earlier version, superseded by `dedupe-lista.py`)

**Purpose:** Find articles that appear in more than one export in a folder
(Scopus topic exports, plus a Zotero CSV export), and write copies of each file
with those articles removed.

**Usage:**
```
python duplicates.py [folder]
```
If no folder is given, the current folder is scanned. Outputs are always written
to the current folder.

**Input:** all `.csv` and `.xlsx` files in the folder (top level only; first sheet
of each workbook). Expects Scopus column names exactly: `DOI`, `EID`, `Title`,
`Year`.

**How duplicates are identified:** each row gets one key, in this order:
1. DOI (trimmed, lowercased)
2. EID (if no DOI)
3. Title + Year (if neither; title lowercased and whitespace collapsed)

Any key that occurs more than once, in any file or within the same file, is a
duplicate.

**Outputs:**
- `duplicates.csv`: one row per duplicated article, with:
  - `key_type`, `key_value`: the key used
  - `OriginalFile`: topic labels, taken from the part of each file name before `_scopus`
  - `AllFilesForKey`: full names of every file the article appeared in
  - the metadata of the first copy found, using the first file's columns
- `dedup/<name>_dedup.<ext>`: each input file with **every** copy of a duplicated
  article removed

Duplicated articles therefore appear only in `duplicates.csv`. To get the full set
of unique records, combine the `dedup/` files with `duplicates.csv`.

## `dedupe-lista.py` (Stage 1)

**Purpose:** Same as `duplicates.py`, adapted to work on LISTA/EBSCO exports
as well as Scopus exports. This is the version used for the LISTA topic exports.

**Usage:**
```
python dedupe-lista.py [folder]
```

**Changes compared with `duplicates.py`:**
- Column names `DOI`/`doi` and `EID`/`eid` are matched regardless of case.
- Title is taken from a column named `title`, `ti`, `document_title` or
  `itemtitle`. If none of these exist, the first column whose name contains
  "title" is used.
- Year is taken from `Year`, then LISTA's `publicationDate`, then `coverDate`,
  then the first 4-digit year (1800–2199) found in any field.
- Topic labels in `OriginalFile` are the part of the file name before `EBSCO`.
  The function is still named `extract_prefix_before_scopus`.

Key order, outputs and the removal of all copies are the same as in
`duplicates.py`.

## `combine345.py` (Stage 2)

**Purpose:** Combine all rated Excel files in a folder into one file, keeping
only records that at least one reviewer rated 3, 4 or 5.

**Usage:** run in the folder containing the rated files
```
python combine345.py
```

**Input:** all `.xlsx` and `.xls` files in the current folder (first sheet only),
except the output file. Each file must have the columns `Reviewer1`,
`Reviewer2`, `Reviewer3` and `Reviewer4` (set in `PEOPLE_COLUMNS`); files
missing any of them are skipped with a warning.

**Selection rule:** a row is kept if any reviewer's cell contains a 3, 4 or 5
anywhere in its text (`3`, `4.0`, `"3, too old"`, `"5 (good overview)"`, …).
Ratings in this review were often written as free text, so the match is
deliberately not limited to cells containing only a number.

**Output:** `filtered_rows.xlsx` (renamed afterwards, e.g. `scopus-zotero-345.xlsx`,
`lista-345.xlsx`), with two added columns:
- `source_file`: the file each row came from (first column)
- `matched_people`: reviewers whose cell matched

Columns from different files are combined: the first file sets the order and new
columns are added at the end.

## `duplicates_2.py` (Stage 3)

**Purpose:** Find the records in the LISTA set that are not already in the
Scopus/Zotero set. The comparison is one-directional: the Scopus/Zotero file
is never modified.

**Usage:** put both files in the current folder and run
```
python duplicates_2.py
```

**Inputs (hard-coded names):**
- `scopus-zotero-345.xlsx`: master set (columns `Title`/`title`, `DOI`/`doi`)
- `lista-345.xlsx`: set to check (columns `title`/`Title`, `doi`/`DOI`)

**Matching rule:** a LISTA record is a duplicate if **either**
- its DOI matches any master DOI (trimmed, lowercased, `http(s)://doi.org/`
  removed), **or**
- its title matches any master title (lowercased, all characters other than
  a–z and 0–9 removed).

**Output:** `new_lista_entries_only.xlsx`, containing the LISTA rows with no match,
plus a printed summary.

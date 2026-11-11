#!/usr/bin/env python3
"""
dedupe-lista.py

Detect duplicates across CSV/XLSX files using:
1. DOI  (case-insensitive)
2. EID  (case-insensitive)
3. Title+Year fallback (robust + LISTA-compatible)

Outputs:
- duplicates.csv (1 row per duplicate)
- dedup/<originalname>_dedup.(csv/xlsx)

Fully supports Scopus + LISTA exports.
"""

import csv
import re
import sys
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook, Workbook

# ---------------------- CONFIG ----------------------

DUPLICATES_OUTPUT = "duplicates.csv"
DEDUP_FOLDER_NAME = "dedup"

# ---------------------- HELPERS ----------------------

def normalize_fieldname(name: str) -> str:
    if name is None:
        return ""
    return name.replace("\ufeff", "").strip()


def normalize_title(title: str) -> str:
    if not title:
        return ""
    return " ".join(title.lower().split())


def extract_year_generic(s: str):
    """Extract the first 4-digit year (>=1800)."""
    if not s:
        return ""
    match = re.search(r"\b(18[0-9]{2}|19[0-9]{2}|20[0-9]{2}|21[0-9]{2})\b", s)
    return match.group(1) if match else ""


# ---------------------- KEY SELECTION ----------------------

def choose_key(row: dict):
    """
    Determine the unique key for a row using:
    1. DOI (any case)
    2. EID (any case)
    3. Title+Year fallback (robust, LISTA-compatible)
    """

    # -------- CASE-INSENSITIVE DOI --------
    doi = ""
    eid = ""
    title = ""
    year = ""

    # Normalize keys once
    lower_map = {k.lower(): k for k in row.keys()}

    # DOI
    if "doi" in lower_map:
        val = row.get(lower_map["doi"], "")
        doi = val.strip().lower() if val else ""

    # EID
    if "eid" in lower_map:
        val = row.get(lower_map["eid"], "")
        eid = val.strip() if val else ""

    # If DOI found, use it
    if doi:
        return ("doi", doi)

    # If EID found, use it
    if eid:
        return ("eid", eid)

    # -------- TITLE FIELD DETECTION --------
    POSSIBLE_TITLE_FIELDS = ["title", "ti", "document_title", "itemtitle"]
    for key in lower_map:
        if key in POSSIBLE_TITLE_FIELDS:
            title = normalize_title(row.get(lower_map[key], "") or "")
            break

    if not title:
        # Try fallback: any field containing 'title'
        for key in lower_map:
            if "title" in key:
                title = normalize_title(row.get(lower_map[key], "") or "")
                break

    if not title:
        return (None, None)

    # -------- YEAR DETECTION --------
    # 1) Scopus: Year
    if "year" in lower_map:
        year = extract_year_generic(row.get(lower_map["year"], ""))

    # 2) LISTA: publicationDate
    if not year and "publicationdate" in lower_map:
        year = extract_year_generic(row.get(lower_map["publicationdate"], ""))

    # 3) LISTA: coverDate
    if not year and "coverdate" in lower_map:
        year = extract_year_generic(row.get(lower_map["coverdate"], ""))

    # 4) Search all fields for 4-digit year
    if not year:
        for key in row:
            y = extract_year_generic(str(row[key]))
            if y:
                year = y
                break

    if title and year:
        return ("title_year", f"{title}||{year}")

    return (None, None)


def extract_prefix_before_scopus(filename: str) -> str:
    """Take 'open_science_scopus_export_...' → 'open_science'."""
    if "EBSCO" in filename:
        return filename.split("EBSCO")[0]
    return filename.rsplit(".", 1)[0]


# ---------------------- FILE I/O ----------------------

def read_csv(path: Path):
    rows = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames:
            return [], []
        header = [normalize_fieldname(x) for x in reader.fieldnames]
        reader.fieldnames = header
        for row in reader:
            rows.append(dict(row))
    return header, rows


def write_csv(path: Path, header, rows):
    with path.open("w", encoding="utf-8", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=header)
        writer.writeheader()
        for row in rows:
            writer.writerow({col: row.get(col, "") for col in header})


def read_xlsx(path: Path):
    wb = load_workbook(path, data_only=True)
    ws = wb.active

    rows_iter = ws.iter_rows(values_only=True)
    try:
        header_row = next(rows_iter)
    except StopIteration:
        return [], []

    header = [normalize_fieldname(str(h) if h else "") for h in header_row]
    rows = []

    for row_cells in rows_iter:
        if row_cells is None:
            continue
        if all(c is None or str(c).strip() == "" for c in row_cells):
            continue
        row = {}
        for idx, col in enumerate(header):
            val = row_cells[idx] if idx < len(row_cells) else ""
            row[col] = "" if val is None else str(val)
        rows.append(row)

    return header, rows


def write_xlsx(path: Path, header, rows):
    wb = Workbook()
    ws = wb.active
    ws.append(header)
    for row in rows:
        ws.append([row.get(col, "") for col in header])
    wb.save(path)


# ---------------------- SCANNING ----------------------

def scan_folder(folder: Path):
    per_file = {}
    occurrences = defaultdict(list)
    global_header = None
    file_id_counter = 0

    files = sorted(list(folder.glob("*.csv")) + list(folder.glob("*.xlsx")))
    if not files:
        print(f"No CSV/XLSX files found in {folder}")
        sys.exit(1)

    for path in files:
        print(f"Reading {path.name} ...")

        if path.suffix.lower() == ".csv":
            header, rows = read_csv(path)
        else:
            header, rows = read_xlsx(path)

        if not header:
            print(f"Skipping empty file: {path.name}")
            continue

        if global_header is None:
            global_header = list(header)

        file_id = file_id_counter
        file_id_counter += 1

        per_file[file_id] = {
            "path": path,
            "ext": path.suffix.lower(),
            "header": header,
            "rows": rows,
        }

        for idx, row in enumerate(rows):
            key_type, key_value = choose_key(row)
            if key_type:
                occurrences[(key_type, key_value)].append((file_id, idx))

    return per_file, occurrences, global_header


# ---------------------- OUTPUT ----------------------

def build_outputs(per_file, occurrences, global_header):
    duplicate_keys = [k for k, occ in occurrences.items() if len(occ) > 1]
    if not duplicate_keys:
        print("No duplicates found.")
        return

    print(f"Found {len(duplicate_keys)} duplicate articles.")

    dedup_folder = Path.cwd() / DEDUP_FOLDER_NAME
    dedup_folder.mkdir(exist_ok=True)

    # Build duplicates.csv
    dup_header = ["key_type", "key_value", "OriginalFile", "AllFilesForKey"] + list(global_header)
    dup_path = Path.cwd() / DUPLICATES_OUTPUT
    dup_rows = []

    for key in duplicate_keys:
        occs = occurrences[key]

        filenames = [per_file[fid]["path"].name for fid, _ in occs]

        prefixes = sorted({extract_prefix_before_scopus(name) for name in filenames})
        prefixes_str = ", ".join(prefixes)
        files_str = ", ".join(sorted(filenames))

        rep_fid, rep_idx = occs[0]
        rep_row = per_file[rep_fid]["rows"][rep_idx]

        key_type, key_value = key

        out = {
            "key_type": key_type,
            "key_value": key_value,
            "OriginalFile": prefixes_str,
            "AllFilesForKey": files_str,
        }

        for col in global_header:
            out[col] = rep_row.get(col, "")

        dup_rows.append(out)

    write_csv(dup_path, dup_header, dup_rows)
    print(f"Wrote duplicates summary: {dup_path}")

    # Dedup outputs
    for fid, info in per_file.items():
        path = info["path"]
        header = info["header"]
        rows = info["rows"]
        ext = info["ext"]

        dedup_rows = []
        for row in rows:
            key_type, key_value = choose_key(row)
            if key_type and (key_type, key_value) in duplicate_keys:
                continue
            dedup_rows.append(row)

        out_name = f"{path.stem}_dedup{path.suffix}"
        out_path = dedup_folder / out_name

        print(f"Writing {out_path.name} ({len(dedup_rows)} kept / {len(rows)} original)")

        if ext == ".csv":
            write_csv(out_path, header, dedup_rows)
        else:
            write_xlsx(out_path, header, dedup_rows)

    print("Done! Deduped files in ./dedup/ and duplicates.csv generated.")


# ---------------------- MAIN ----------------------

def main():
    if len(sys.argv) >= 2:
        folder = Path(sys.argv[1]).expanduser().resolve()
    else:
        folder = Path.cwd()

    print(f"Scanning folder: {folder}")

    per_file, occurrences, global_header = scan_folder(folder)
    build_outputs(per_file, occurrences, global_header)


if __name__ == "__main__":
    main()

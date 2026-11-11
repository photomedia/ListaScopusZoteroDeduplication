#!/usr/bin/env python3
"""
duplicates.py

Detect duplicates across CSV/XLSX files using DOI → EID → Title+Year.
Writes:
- duplicates.csv (1 row per duplicate article)
- dedup/<filename>_dedup.(csv/xlsx) with duplicates removed

Duplicates.csv includes:
- key_type, key_value
- OriginalFile = prefixes before _scopus (sorted, unique)
- AllFilesForKey = full filenames where article appeared
- All original metadata columns
"""

import csv
import sys
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook, Workbook

# ---------------------- CONFIG ----------------------

COL_TITLE = "Title"
COL_YEAR = "Year"
COL_DOI = "DOI"
COL_EID = "EID"

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


def choose_key(row: dict):
    """Choose DOI → EID → Title+Year."""
    doi = (row.get(COL_DOI) or "").strip().lower()
    eid = (row.get(COL_EID) or "").strip()
    title = normalize_title(row.get(COL_TITLE) or "")
    year = (row.get(COL_YEAR) or "").strip()

    if doi:
        return ("doi", doi)
    if eid:
        return ("eid", eid)
    if title and year:
        return ("title_year", f"{title}||{year}")
    return (None, None)


def extract_prefix_before_scopus(filename: str) -> str:
    if "_scopus" in filename:
        return filename.split("_scopus")[0]
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
        print(f"No files found in {folder}")
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
    duplicate_keys = [k for k, occs in occurrences.items() if len(occs) > 1]
    if not duplicate_keys:
        print("No duplicates found.")
        return

    print(f"Found {len(duplicate_keys)} duplicate articles.")

    # Create dedup folder (Option B: subfolder in CWD)
    dedup_folder = Path.cwd() / DEDUP_FOLDER_NAME
    dedup_folder.mkdir(exist_ok=True)

    # Prepare duplicates.csv
    out_header = ["key_type", "key_value", "OriginalFile", "AllFilesForKey"] + list(global_header)
    duplicates_path = Path.cwd() / DUPLICATES_OUTPUT

    duplicates_rows = []

    for key in duplicate_keys:
        occs = occurrences[key]

        # All filenames
        filenames = [per_file[file_id]["path"].name for file_id, _ in occs]

        # Extract prefixes before "_scopus"
        prefixes = sorted({extract_prefix_before_scopus(name) for name in filenames})
        prefixes_str = ", ".join(prefixes)

        all_files_str = ", ".join(sorted(filenames))

        # Representative row (first occurrence)
        rep_file_id, rep_row_idx = occs[0]
        rep_row = per_file[rep_file_id]["rows"][rep_row_idx]

        key_type, key_value = key

        outrow = {
            "key_type": key_type,
            "key_value": key_value,
            "OriginalFile": prefixes_str,
            "AllFilesForKey": all_files_str,
        }

        for col in global_header:
            outrow[col] = rep_row.get(col, "")

        duplicates_rows.append(outrow)

    # Write duplicates.csv
    write_csv(duplicates_path, out_header, duplicates_rows)
    print(f"Wrote duplicates summary: {duplicates_path}")

    # Build deduped per-file outputs
    for file_id, info in per_file.items():
        path = info["path"]
        rows = info["rows"]
        header = info["header"]
        ext = info["ext"]

        dedup_rows = []
        for row in rows:
            key_type, key_value = choose_key(row)
            if key_type and (key_type, key_value) in duplicate_keys:
                continue
            dedup_rows.append(row)

        dedup_name = f"{path.stem}_dedup{path.suffix}"
        dedup_path = dedup_folder / dedup_name

        print(f"Writing: {dedup_path.name}  "
              f"({len(dedup_rows)} kept / {len(rows)} original)")

        if ext == ".csv":
            write_csv(dedup_path, header, dedup_rows)
        else:
            write_xlsx(dedup_path, header, dedup_rows)

    print("Done. Duplicates removed. Dedup files in ./dedup/. Summary in duplicates.csv.")


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

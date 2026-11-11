import pandas as pd
import re
from pathlib import Path

print("Starting extraction script...")

# -------------------------------------------------
# Configuration
# -------------------------------------------------
WORKDIR = Path.cwd()  # current working directory
OUTPUT_FILE = WORKDIR / "filtered_rows.xlsx"

PEOPLE_COLUMNS = ["Reviewer1", "Reviewer2", "Reviewer3", "Reviewer4"]

# Match 3, 4, 5, optionally .0, anywhere in text
RATING_PATTERN = re.compile(r"[345](?:\.0+)?")

# -------------------------------------------------
# Helper function
# -------------------------------------------------
def find_matching_people(row):
    """
    Returns list of people whose column contains
    3, 4, or 5 (optionally as 3.0, 4.0, etc.)
    """
    matched = []

    for person in PEOPLE_COLUMNS:
        value = row.get(person)

        if pd.isna(value):
            continue

        if RATING_PATTERN.search(str(value)):
            matched.append(person)

    return matched

# -------------------------------------------------
# Main processing
# -------------------------------------------------
all_matches = []
canonical_columns = None  # preserves original column order

# Collect .xlsx and .xls files
excel_files = list(WORKDIR.glob("*.xlsx")) + list(WORKDIR.glob("*.xls"))

# Exclude output file if re-running
excel_files = [
    f for f in excel_files
    if f.resolve() != OUTPUT_FILE.resolve()
]

if not excel_files:
    print(f"No Excel files found in {WORKDIR}")
else:
    for excel_file in sorted(excel_files):
        print(f"Processing {excel_file.name}")

        try:
            if excel_file.suffix.lower() == ".xls":
                df = pd.read_excel(excel_file, engine="xlrd")
            else:
                df = pd.read_excel(excel_file, engine="openpyxl")
        except Exception as e:
            print(f"Failed to read {excel_file.name}: {e}")
            continue

        # Validate expected people columns
        missing = [c for c in PEOPLE_COLUMNS if c not in df.columns]
        if missing:
            print(f"Skipping {excel_file.name} (missing columns: {missing})")
            continue

        # Apply rating match
        df["matched_people"] = df.apply(find_matching_people, axis=1)
        filtered = df[df["matched_people"].astype(bool)].copy()

        if filtered.empty:
            print("   ↳ No matches found")
            continue

        filtered["matched_people"] = filtered["matched_people"].apply(", ".join)

        # Always add source_file
        filtered["source_file"] = excel_file.name

        # -------------------------------------------------
        # Preserve column order across heterogeneous files
        # -------------------------------------------------
        if canonical_columns is None:
            # First valid file defines baseline column order
            canonical_columns = list(filtered.columns)

            # Force source_file to be first
            canonical_columns.remove("source_file")
            canonical_columns.insert(0, "source_file")
        else:
            for col in filtered.columns:
                if col not in canonical_columns:
                    canonical_columns.append(col)

        # Reindex to canonical order (fills missing cols with NaN)
        filtered = filtered.reindex(columns=canonical_columns)

        all_matches.append(filtered)

# -------------------------------------------------
# Output
# -------------------------------------------------
if all_matches:
    result_df = pd.concat(all_matches, ignore_index=True, sort=False)
    result_df.to_excel(OUTPUT_FILE, index=False)
    print(f"Saved {len(result_df)} rows to {OUTPUT_FILE.name}")
else:
    print("No matching rows found across the Excel files.")

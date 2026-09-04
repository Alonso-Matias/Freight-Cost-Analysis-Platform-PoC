"""
Code 1: Ingest Worker — Read downloads, map, create posting file.
================================================================
This is the config-driven version of process_incoming_posting.py.
Instead of hardcoded values, it reads everything from YAML configs.

What it does:
  1. Reads all .xlsx files from the download folder
  2. Combines them (stacks one below the other)
  3. Adds "Type of cost" from charge_type_mapping.yaml (matched by Charge Type)
  4. Adds "Carrier Name" from carrier_mapping.yaml + carrier_map.parquet fallback
  5. Adds "Posting Month" (derived from Posting Date)
  6. Saves as Posting_YYYY-MM.xlsx in the postings folder

To adapt for another department: edit configs/, not this file.
"""
import pandas as pd
import numpy as np
from pathlib import Path
import sys
import shutil

# Add parent to path so we can import config_loader
sys.path.insert(0, str(Path(__file__).parent))
from config_loader import (
    get_target_columns,
    get_charge_type_map,
    get_carrier_map,
    clear_cache,
)

# Default paths (can be overridden by master.py)
DEFAULT_BASE = Path(r"W:\1. SELS WEDC & GPCE\Accrual Accuracy")
DEFAULT_DOWNLOAD_DIR = DEFAULT_BASE / "Download"
DEFAULT_POSTING_DIR = DEFAULT_BASE / "Postings"


def load_carrier_name_mapping():
    """Load carrier mapping from carrier_mapping.yaml (no parquet cache needed).
    
    All 1386 carrier mappings live in YAML. Edit there, not in code.
    """
    mapping = get_carrier_map()
    print(f"  Loaded {len(mapping)} carrier mappings from carrier_mapping.yaml")
    return mapping



def process_downloads(download_dir=None, posting_dir=None):
    """
    Main entry point: process BMS downloads into posting files.

    Args:
        download_dir: Path to folder with T4BA.xlsx, T4BN.xlsx (default: W: drive)
        posting_dir: Path to output folder for Posting_YYYY-MM.xlsx
    """
    # Reload configs in case they were edited
    clear_cache()

    download_dir = Path(download_dir) if download_dir else DEFAULT_DOWNLOAD_DIR
    posting_dir = Path(posting_dir) if posting_dir else DEFAULT_POSTING_DIR

    print("=" * 70)
    print("Code 1: Ingest Worker — Processing BMS Downloads")
    print("=" * 70)

    # 1. Find all .xlsx files in Download folder
    download_files = sorted(download_dir.glob("*.xlsx"))
    download_files = [f for f in download_files if not f.name.startswith("~$")]
    if not download_files:
        print(f"\nNo .xlsx files found in: {download_dir}")
        print("Drop your T4BA.xlsx and T4BN.xlsx files there and try again.")
        return None

    print(f"\nFound {len(download_files)} file(s) in Download folder:")
    for f in download_files:
        print(f"  - {f.name}")

    # 2. Load and combine all files
    print(f"\n--- Loading and combining files ---")
    all_dfs = []
    for f in download_files:
        print(f"  Reading {f.name}...", end=" ", flush=True)
        df = pd.read_excel(f, sheet_name="Data")
        df["_source_file"] = f.name
        all_dfs.append(df)
        print(f"{len(df)} rows")
    combined = pd.concat(all_dfs, ignore_index=True)
    print(f"  Combined total: {len(combined)} rows")

    # 3. Strip column names
    combined.columns = [c.strip() if isinstance(c, str) else c for c in combined.columns]

    # 4. Add "Type of cost" from config
    print(f"\n--- Adding 'Type of cost' from charge_type_mapping.yaml ---")
    toc_map = get_charge_type_map()
    if "Charge Type" in combined.columns:
        ct_numeric = pd.to_numeric(combined["Charge Type"], errors="coerce")
        ct_str = ct_numeric.dropna().astype(int).astype(str)
        combined["_ct_key"] = ct_numeric.where(ct_numeric.isna(), ct_str)
        combined["Type of cost"] = combined["_ct_key"].map(toc_map)
        combined = combined.drop(columns=["_ct_key"])

        missing_toc = combined["Type of cost"].isna().sum()
        total = len(combined)
        print(f"  Mapped: {total - missing_toc}/{total} rows ({(total-missing_toc)/total*100:.1f}%)")
        if missing_toc > 0:
            print(f"  WARNING: {missing_toc} rows have no Type of cost mapping")
            unmapped = combined[combined["Type of cost"].isna()]["Charge Type"].unique()
            if len(unmapped) <= 20:
                print(f"  Unmapped Charge Types: {unmapped.tolist()}")

    # 5. Add "Carrier Name" from config + cache
    print(f"\n--- Adding 'Carrier Name' from carrier_mapping.yaml ---")
    carrier_map = load_carrier_name_mapping()
    if "Carrier" in combined.columns:
        combined["Carrier Name"] = combined["Carrier"].astype(str).str.strip().map(carrier_map)
        missing_cn = combined["Carrier Name"].isna().sum()
        total = len(combined)
        print(f"  Mapped: {total - missing_cn}/{total} rows ({(total-missing_cn)/total*100:.1f}%)")
        if missing_cn > 0:
            unmapped = combined[combined["Carrier Name"].isna()]["Carrier"].unique()
            if len(unmapped) <= 20:
                print(f"  Unmapped Carrier IDs: {unmapped.tolist()}")

    # 6. Add "Posting Month" from Posting Date
    print(f"\n--- Adding 'Posting Month' from Posting Date ---")
    if "Posting Date" in combined.columns:
        combined["Posting Date"] = pd.to_datetime(combined["Posting Date"], errors="coerce")
        combined["Posting Month"] = combined["Posting Date"].dt.to_period("M").astype(str)
        months = combined["Posting Month"].dropna().unique()
        print(f"  Posting months found: {sorted(months.tolist())}")
    else:
        print("  WARNING: No 'Posting Date' column found!")

    # 7. Save posting files per month
    months_in_data = sorted(combined["Posting Month"].dropna().unique().tolist())
    print(f"\n--- Creating posting file(s) for: {months_in_data} ---")

    target_cols = get_target_columns()
    output_files = []

    for month_label in months_in_data:
        month_df = combined[combined["Posting Month"] == month_label].copy()

        # Add any missing target columns as empty
        for col in target_cols:
            if col not in month_df.columns:
                month_df[col] = np.nan

        # Keep only target columns, in the right order
        output_df = month_df[target_cols].copy()

        # Clean up: remove rows where Plant is NaN
        plant_before = len(output_df)
        if "Plant" in output_df.columns:
            output_df = output_df[output_df["Plant"].notna()].copy()
        plant_after = len(output_df)
        if plant_before != plant_after:
            print(f"  Dropped {plant_before - plant_after} rows with no Plant")

        plants = output_df["Plant"].unique().tolist()
        print(f"\n  Month {month_label}: {len(output_df)} rows, Plants: {plants}")

        # Check if file already exists → backup
        output_file = posting_dir / f"Posting_{month_label}.xlsx"
        if output_file.exists():
            print(f"  WARNING: {output_file.name} already exists!")
            print(f"  Overwriting with new data...")
            backup = posting_dir / f"Posting_{month_label}_backup.xlsx"
            if not backup.exists():
                shutil.copy2(output_file, backup)
                print(f"  Backed up old file to: {backup.name}")

        # Save
        print(f"  Saving to: {output_file.name}...", end=" ", flush=True)
        try:
            with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
                output_df.to_excel(writer, sheet_name="Data", index=False)
            print(f"DONE ({len(output_df)} rows)")
            output_files.append(output_file)
        except PermissionError:
            print(f"FAILED - file is locked (open in Excel?)")
            print(f"  Saving to temp file instead: {output_file.stem}_NEW.xlsx")
            temp_file = posting_dir / f"{output_file.stem}_NEW.xlsx"
            with pd.ExcelWriter(temp_file, engine="openpyxl") as writer:
                output_df.to_excel(writer, sheet_name="Data", index=False)
            print(f"  Saved! Close the original in Excel, then rename the _NEW file.")
            output_files.append(temp_file)

    # 8. Summary
    print(f"\n{'=' * 70}")
    print("CODE 1 COMPLETE")
    print(f"{'=' * 70}")
    print(f"  Files processed: {len(download_files)}")
    print(f"  Total rows: {len(combined)}")
    print(f"  Months created: {months_in_data}")
    print(f"  Output location: {posting_dir}")
    print(f"{'=' * 70}")

    return output_files


if __name__ == "__main__":
    process_downloads()

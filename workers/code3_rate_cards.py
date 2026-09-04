"""
Code 3: Rate Card Worker — Compare 2025 vs 2026 contract rates.
===============================================================
Extracts and compares rate card data from B2B contract files.
Matches rates by carrier + destination country, isolating rate increases.

This worker is optional — it runs only if rate card files exist.
"""
import pandas as pd
import numpy as np
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).parent))
from config_loader import load_formula_config, clear_cache

# Default rate card file location
DEFAULT_RATES_FILE = Path(r"W:\1. SELS WEDC & GPCE\Accrual Accuracy\GPCE 2026 - Rates.xlsx")


def load_rate_cards(rates_file=None):
    """
    Load rate card data from the contract file.
    Returns (old_rates_df, new_rates_df) for 2025 and 2026.

    Expected columns: Carrier, Country, Truck Type, Rate
    """
    rates_file = Path(rates_file) if rates_file else DEFAULT_RATES_FILE
    if not rates_file.exists():
        print(f"  Rate card file not found: {rates_file}")
        return None, None

    print(f"  Loading rate cards from: {rates_file.name}")
    xls = pd.ExcelFile(rates_file)
    print(f"  Sheets: {xls.sheet_names}")

    # Try to find 2025 and 2026 rate sheets
    old_df = None
    new_df = None

    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet)
        df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]

        if "2025" in sheet or "old" in sheet.lower():
            old_df = df
            print(f"    2025 rates: sheet '{sheet}' ({len(df)} rows)")
        elif "2026" in sheet or "new" in sheet.lower():
            new_df = df
            print(f"    2026 rates: sheet '{sheet}' ({len(df)} rows)")

    # If only one sheet, treat it as the new rates
    if new_df is None and len(xls.sheet_names) == 1:
        new_df = pd.read_excel(xls, sheet_name=0)
        new_df.columns = [c.strip() if isinstance(c, str) else c for c in new_df.columns]
        print(f"    Using single sheet as 2026 rates ({len(new_df)} rows)")

    return old_df, new_df


def compare_rates(old_df, new_df, carrier_col="Carrier", country_col="Country",
                   rate_col="Rate", truck_col="Truck Type"):
    """
    Compare old vs new rate cards by carrier + country + truck type.

    Returns a DataFrame with columns:
      carrier, country, truck_type, old_rate, new_rate, delta, delta_pct
    """
    if old_df is None or new_df is None:
        print("  Cannot compare: missing one or both rate card DataFrames")
        return pd.DataFrame()

    # Normalize column names
    for df in [old_df, new_df]:
        if carrier_col not in df.columns:
            # Try to find a carrier-like column
            for c in df.columns:
                if "carrier" in str(c).lower() or "lsp" in str(c).lower():
                    df.rename(columns={c: carrier_col}, inplace=True)
                    break

    # Merge on carrier + country + truck type
    merge_keys = [k for k in [carrier_col, country_col, truck_col] if k in old_df.columns and k in new_df.columns]
    if not merge_keys:
        print(f"  No common merge keys found. Columns: old={list(old_df.columns)}, new={list(new_df.columns)}")
        return pd.DataFrame()

    merged = old_df[merge_keys + [rate_col]].merge(
        new_df[merge_keys + [rate_col]],
        on=merge_keys,
        suffixes=("_old", "_new"),
        how="outer",
    )

    # Calculate deltas
    merged["delta"] = merged[f"{rate_col}_new"] - merged[f"{rate_col}_old"]
    merged["delta_pct"] = (merged["delta"] / merged[f"{rate_col}_old"] * 100).where(merged[f"{rate_col}_old"] != 0)

    # Sort by biggest absolute increase
    merged = merged.sort_values("delta", ascending=False)

    return merged


def run_rate_card_analysis(rates_file=None):
    """
    Main entry point: load and compare rate cards.

    Returns a summary dict or None if rate cards not available.
    """
    clear_cache()
    print("=" * 70)
    print("Code 3: Rate Card Worker — Comparing 2025 vs 2026 Contract Rates")
    print("=" * 70)

    old_df, new_df = load_rate_cards(rates_file)
    if old_df is None and new_df is None:
        print("\n  No rate card file found. Skipping rate card analysis.")
        print("  (This worker is optional — the dashboard works without it.)")
        return None

    if old_df is not None and new_df is not None:
        result = compare_rates(old_df, new_df)
        if not result.empty:
            print(f"\n  Compared {len(result)} rate lanes")
            print(f"  Average rate change: {result['delta_pct'].mean():.1f}%")
            print(f"  Lanes with increases: {(result['delta'] > 0).sum()}")
            print(f"  Lanes with decreases: {(result['delta'] < 0).sum()}")
            return result
    else:
        print("\n  Could not find both 2025 and 2026 rate sheets.")
        print("  Rate card comparison requires both periods.")

    return None


if __name__ == "__main__":
    result = run_rate_card_analysis()
    if result is not None:
        print(f"\nTop 10 rate increases:")
        print(result.head(10).to_string(index=False))

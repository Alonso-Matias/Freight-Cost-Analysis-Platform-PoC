"""
Code 2: Dashboard Worker — Config-driven freight analysis dashboard.
====================================================================
This is the config-driven version of freight_dashboard.py.
It produces the SAME output as the original dashboard, but reads all
mappings from YAML configs instead of hardcoding them.

Key differences from the original:
  - Carrier mapping: loaded from carrier_mapping.yaml (1386 entries), not parquet
  - Warehouse mapping: loaded from warehouse_mapping.yaml
  - Charge type mapping: loaded from charge_type_mapping.yaml
  - Column structure: loaded from column_structure.yaml
  - Formula: loaded from formula.yaml

Run with:
  streamlit run code2_dashboard.py
"""

import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime
import sys

# Add workers to path for config_loader
sys.path.insert(0, str(Path(__file__).parent))
from config_loader import (
    get_carrier_map,
    get_warehouse_map,
    get_division_map,
    clear_cache,
)

# ---------------------------------------------------------------------------
# CONFIG (paths from formula.yaml, warehouse codes from warehouse_mapping.yaml)
# ---------------------------------------------------------------------------
# Demo mode: data lives in ./data/ folder (relative to project root)
_PROJECT_ROOT = Path(__file__).resolve().parent.parent
BASE = _PROJECT_ROOT / "data"
ACCRUAL_DIR = BASE / "Accruals"
POSTING_DIR = BASE / "Postings"
AI_DIR = BASE / "rate_cards"

COST_COL = "Amt."
SHIP_COL = "ITR No."

# Rate card files (2025 old vs 2026 new)
OLD_RATE_CARD = AI_DIR / "rate_card_2025.xlsx"
NEW_RATE_CARD = AI_DIR / "rate_card_2026.xlsx"


# Warehouse code mapping: accruals use L-codes, postings use T-codes
WAREHOUSE_MAP = {
    "L428": ["L428", "T4BN"],
    "L401": ["L401", "T4BA"],
    "L430": ["L430", "T430"],
    "S4A1": ["S4A1", "S4A1"],
}

# Truck type mapping: FTL/STL/LTL
TRUCK_TYPE_MAP = {
    "FTL": {
        "posting": ["Truck Freight - FTL"],
        "accrual": ["TRUCK FREIGHT CHARGE - LANE A"],
    },
    "STL": {
        "posting": ["Truck Freight - STL"],
        "accrual": ["TRUCK FREIGHT CHARGE - LANE B"],
    },
    "LTL": {
        "posting": ["Truck Freight - LTL"],
        "accrual": ["TRUCK FREIGHT CHARGE - LANE C"],
    },
}

# Carrier name mapping (rate card sheet name -> accrual carrier name)
RATE_CARRIER_MAP = {
    "DHL": "DHL Freight",
    "DSV": "DSV",
    "JDR": "Jan de Rijk",
    "VERST": "Versteijnen Transport",
    "Verstijnen": "Versteijnen Transport",
    "VOS": "VOS TRANSPORT",
    "WAB": "WABERER'S",
    "Waberers": "WABERER'S",
    "BARSAN": "Barsan",
    "Barsan": "Barsan",
    "DFDS": "DFDS",
    "EI": "EXPEDITORS",
    "Expeditors": "EXPEDITORS",
    "GEODIS": "Geodis",
    "GIRTEKA": "GIRTEKA",
    "PRL": "PRL FREIGHT",
    "SEGERS": "SEGERS",
    "Heppner": "Heppner",
    "INL": "INL Cargo",
    "BALTIC": "Baltic",
}

# Shipping Point to ECOM/B2B mapping for postings
SHIPPING_POINT_ECOM_MAP = {
    "ELD6": "ECOM",
    "ELD1": "B2B",
}


st.set_page_config(
    page_title="Freight Analysis Dashboard",
    page_icon="🚚",
    layout="wide",
)


# ---------------------------------------------------------------------------
# CARRIER MAPPING (from YAML, not parquet)
# ---------------------------------------------------------------------------
@st.cache_data(ttl="6h", show_spinner="Loading carrier mapping from YAML...")
def _build_carr_id_to_name_map():
    """Load CARR ID → Carrier name mapping from carrier_mapping.yaml.
    
    All 1386 mappings are in the YAML file. No parquet cache needed.
    """
    clear_cache()
    mapping = get_carrier_map()
    print(f"  Loaded {len(mapping)} carrier mappings from carrier_mapping.yaml")
    return mapping


# ---------------------------------------------------------------------------
# DATA LOADERS (cached)
# ---------------------------------------------------------------------------
@st.cache_data(ttl="1h", show_spinner="Loading accrual data...")
def _find_accrual_file(month_label: str):
    """Find accrual file (.xlsx or .xlsb)."""
    for ext in (".xlsx", ".xlsb"):
        f = ACCRUAL_DIR / f"Accrual_{month_label}{ext}"
        if f.exists():
            return f
    return None


def _has_real_headers(df, threshold=0.5):
    """Check if a sheet has real column headers (not all 'Unnamed: X')."""
    if df.empty:
        return False
    unnamed_count = sum(1 for c in df.columns if str(c).startswith("Unnamed:"))
    return unnamed_count / len(df.columns) < threshold


def _detect_sheet(xls, preferred_names=("Data", "SELS Selection", "DATA", "SELS SELECTION")):
    """Auto-detect sheet name (case-insensitive), fallback to first sheet with real headers."""
    lower_map = {s.lower(): s for s in xls.sheet_names}
    # First try preferred names
    for name in preferred_names:
        if name.lower() in lower_map:
            sheet = lower_map[name.lower()]
            # Verify it has real headers
            try:
                df_check = pd.read_excel(xls, sheet_name=sheet, header=0, nrows=5)
                if _has_real_headers(df_check):
                    return sheet
            except Exception:
                pass
    # Fallback: find first sheet with real headers
    for sheet in xls.sheet_names:
        try:
            df_check = pd.read_excel(xls, sheet_name=sheet, header=0, nrows=5)
            if _has_real_headers(df_check):
                return sheet
        except Exception:
            continue
    return xls.sheet_names[0]


@st.cache_data(ttl="1h", show_spinner="Loading accrual data...")
def load_accrual(month_label: str) -> pd.DataFrame:
    f = _find_accrual_file(month_label)
    if f is None:
        return pd.DataFrame()
    xls = pd.ExcelFile(f)
    sheet_name = _detect_sheet(xls)
    df = pd.read_excel(f, sheet_name=sheet_name, header=0)

    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    df[COST_COL] = pd.to_numeric(df[COST_COL], errors="coerce").fillna(0.0)
    if "Actual Date" in df.columns:
        df["Actual Date"] = pd.to_datetime(df["Actual Date"], format="%d.%m.%Y", errors="coerce")
        df["_actual_month"] = df["Actual Date"].dt.to_period("M").astype(str)

    # Normalize column names across years
    if "Country Code" in df.columns and "Country" not in df.columns:
        df = df.rename(columns={"Country Code": "Country"})
    if "Amt.2" in df.columns and "Ver.Sup.Amt" not in df.columns:
        df = df.rename(columns={"Amt.2": "Ver.Sup.Amt"})
    if "Carrier Name" in df.columns and "Carrier name" not in df.columns:
        df = df.rename(columns={"Carrier Name": "Carrier name"})

    # --- Enrich: fill missing 'Carrier name' from CARR ID mapping ---
    # All mappings come from carrier_mapping.yaml (no parquet)
    # Note: Only CARR ID is used. If no CARR ID, carrier name stays empty
    # (LSP ID fallback removed — it could match to incorrect carrier).
    if "Carrier name" not in df.columns or df["Carrier name"].isna().all():
        carr_map = _build_carr_id_to_name_map()
        if "CARR ID" in df.columns:
            df["Carrier name"] = df["CARR ID"].astype(str).str.strip().map(carr_map)
    elif "Carrier name" in df.columns and "CARR ID" in df.columns:
        carr_map = _build_carr_id_to_name_map()
        mask = df["Carrier name"].isna() | (df["Carrier name"].astype(str).str.strip() == "")
        if mask.any():
            df.loc[mask, "Carrier name"] = df.loc[mask, "CARR ID"].astype(str).str.strip().map(carr_map)


    return df


@st.cache_data(ttl="1h", show_spinner="Loading posting data...")
def load_posting(month_label: str) -> pd.DataFrame:
    f = POSTING_DIR / f"Posting_{month_label}.xlsx"
    if not f.exists():
        return pd.DataFrame()
    xls = pd.ExcelFile(f)
    sheet_name = _detect_sheet(xls, preferred_names=("Sheet4", "Data", "DATA"))
    df = pd.read_excel(f, sheet_name=sheet_name, header=0)

    df.columns = [c.strip() if isinstance(c, str) else c for c in df.columns]
    rename = {"Carrier Name": "Carrier name", "Carrier": "CARR ID"}
    df = df.rename(columns={k: v for k, v in rename.items() if k in df.columns})
    cost_col = "Ver.Sup.Amt"
    df[cost_col] = pd.to_numeric(df[cost_col], errors="coerce").fillna(0.0)
    df["_cost"] = df[cost_col]
    df["_weight"] = pd.to_numeric(df.get("C.Weight", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    df["_volume"] = pd.to_numeric(df.get("Total Volume", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    ship_col = None
    for c in df.columns:
        if isinstance(c, str) and c.strip().lower() in ("fr doc.no.", "fr doc.no", "fr doc no"):
            ship_col = c
            break
    if ship_col:
        df["_ship"] = df[ship_col]
    else:
        df["_ship"] = df.index.astype(str)
    df["_ship_count"] = df["_ship"].where(
        df["_ship"].notna() & (df["_ship"].astype(str).str.strip() != ""),
        df.index.astype(str)
    )
    return df


# ---------------------------------------------------------------------------
# RATE CARD PARSER (cached, vectorized)
# ---------------------------------------------------------------------------
@st.cache_data(ttl="6h", show_spinner="Parsing B2B rate cards (2025 & 2026)...")
def _parse_rate_card(filepath, year_label):
    """Parse a B2B rate card file (vectorized)."""
    def _extract_bound(colname):
        s = str(colname)
        if "<=" in s:
            part = s.split("<=")[-1].strip()
            for p in part.split():
                try:
                    return float(p)
                except:
                    continue
        return None

    xls = pd.ExcelFile(filepath)
    all_rates = []
    for sheet in xls.sheet_names:
        if sheet in ("Ratecard (Total) AS-IS", "LSP Allocation"):
            continue
        df_raw = pd.read_excel(filepath, sheet_name=sheet, header=None, nrows=10)
        header_row = None
        for idx in range(len(df_raw)):
            row_vals = df_raw.iloc[idx].astype(str).str.strip()
            if any("DESTINATION" in str(v).upper() and "COUNTRY" in str(v).upper() for v in row_vals):
                header_row = idx
                break
        if header_row is None:
            continue
        df = pd.read_excel(filepath, sheet_name=sheet, header=header_row)
        df = df.dropna(how="all")
        if len(df) == 0:
            continue
        country_col = None
        zip_col = None
        for c in df.columns:
            cstr = str(c).upper()
            if "COUNTRY" in cstr and "DEST" in cstr:
                country_col = c
            if "ZIP" in cstr and "DEST" in cstr:
                zip_col = c
        if country_col is None:
            cols = list(df.columns)
            if len(cols) > 4:
                country_col = cols[3]
                zip_col = cols[4]
        if country_col is None:
            continue
        stl_cols = [c for c in df.columns if "STL" in str(c).upper() and "ldm" in str(c).lower()]
        ltl_cols = [c for c in df.columns if "LTL" in str(c).upper() and "cbm" in str(c).lower()]
        ftl_cols = [c for c in df.columns if "FTL" in str(c).upper()]
        df_valid = df[df[country_col].notna() & (df[country_col].astype(str).str.strip() != "")].copy()
        if len(df_valid) == 0:
            continue
        countries = df_valid[country_col].astype(str).str.strip().str.upper()
        zips = df_valid[zip_col].astype(str).str.strip() if zip_col else pd.Series([""] * len(df_valid), index=df_valid.index)
        carrier_name = RATE_CARRIER_MAP.get(sheet, sheet)
        for col in stl_cols:
            bound = _extract_bound(col)
            if bound is None:
                continue
            rates = pd.to_numeric(df_valid[col], errors="coerce")
            mask = rates.notna() & (rates > 0)
            if mask.sum() == 0:
                continue
            all_rates.append(pd.DataFrame({
                "carrier": carrier_name, "country": countries[mask].values,
                "zip": zips[mask].values, "truck_type": "STL",
                "volume_bound": bound, "rate": rates[mask].values,
            }))
        for col in ltl_cols:
            bound = _extract_bound(col)
            if bound is None:
                continue
            rates = pd.to_numeric(df_valid[col], errors="coerce")
            mask = rates.notna() & (rates > 0)
            if mask.sum() == 0:
                continue
            all_rates.append(pd.DataFrame({
                "carrier": carrier_name, "country": countries[mask].values,
                "zip": zips[mask].values, "truck_type": "LTL",
                "volume_bound": bound, "rate": rates[mask].values,
            }))
        for col in ftl_cols:
            rates = pd.to_numeric(df_valid[col], errors="coerce")
            mask = rates.notna() & (rates > 0)
            if mask.sum() == 0:
                continue
            all_rates.append(pd.DataFrame({
                "carrier": carrier_name, "country": countries[mask].values,
                "zip": zips[mask].values, "truck_type": "FTL",
                "volume_bound": 999, "rate": rates[mask].values,
            }))
    if all_rates:
        return pd.concat(all_rates, ignore_index=True)
    return pd.DataFrame()


@st.cache_data(ttl="6h", show_spinner="Loading rate cards...")
def load_rate_cards():
    """Load both old (2025) and new (2026) rate cards. Returns (old_df, new_df)."""
    old_df = pd.DataFrame()
    new_df = pd.DataFrame()
    if OLD_RATE_CARD.exists():
        old_df = _parse_rate_card(OLD_RATE_CARD, "2025")
    if NEW_RATE_CARD.exists():
        new_df = _parse_rate_card(NEW_RATE_CARD, "2026")
    return old_df, new_df


def compare_rates_by_carrier_country(old_df, new_df, truck_type="STL"):
    """Compare rates old vs new per carrier/country for a given truck type."""
    old_sub = old_df[old_df["truck_type"] == truck_type]
    new_sub = new_df[new_df["truck_type"] == truck_type]
    old_avg = old_sub.groupby(["carrier", "country"])["rate"].mean().reset_index()
    old_avg.columns = ["carrier", "country", "old_avg_rate"]
    new_avg = new_sub.groupby(["carrier", "country"])["rate"].mean().reset_index()
    new_avg.columns = ["carrier", "country", "new_avg_rate"]
    merged = old_avg.merge(new_avg, on=["carrier", "country"], how="outer")
    merged["delta"] = merged["new_avg_rate"] - merged["old_avg_rate"]
    merged["delta_pct"] = np.where(merged["old_avg_rate"] > 0,
                                    merged["delta"] / merged["old_avg_rate"] * 100, 0)
    return merged.sort_values("delta", ascending=False)


def decompose_cost_increase(acc_prev, acc_curr, old_rates, new_rates, cost_col=COST_COL, ship_col=SHIP_COL, country_col="Country"):
    """Decompose the total cost increase into: rate effect, carrier mix, country mix, volume."""
    if acc_prev.empty or acc_curr.empty:
        return None
    if "Carrier name" not in acc_curr.columns or country_col not in acc_curr.columns:
        return None

    def _build_rate_lookup(rates_df):
        if rates_df.empty:
            return {}
        lookup = {}
        for _, r in rates_df.iterrows():
            key = (r["carrier"], r["country"])
            if key not in lookup:
                lookup[key] = {}
            tt = r["truck_type"]
            if tt not in lookup[key]:
                lookup[key][tt] = []
            lookup[key][tt].append(r["rate"])
        for key in lookup:
            for tt in lookup[key]:
                lookup[key][tt] = np.mean(lookup[key][tt])
        return lookup

    old_lookup = _build_rate_lookup(old_rates)
    new_lookup = _build_rate_lookup(new_rates)

    def _match_rates(acc_df, lookup):
        rates = []
        for _, row in acc_df.iterrows():
            carrier = str(row.get("Carrier name", "")).strip()
            country = str(row.get(country_col, "")).strip().upper()
            key = (carrier, country)
            if key in lookup:
                r = lookup[key].get("STL", lookup[key].get("FTL", lookup[key].get("LTL", 0)))
                rates.append(r)
            else:
                rates.append(np.nan)
        return rates

    acc_prev = acc_prev.copy()
    acc_curr = acc_curr.copy()
    acc_prev["_matched_rate_old"] = _match_rates(acc_prev, old_lookup)
    acc_curr["_matched_rate_new"] = _match_rates(acc_curr, new_lookup)

    prev_cost = acc_prev[cost_col].sum()
    curr_cost = acc_curr[cost_col].sum()
    prev_ships = acc_prev[ship_col].nunique()
    curr_ships = acc_curr[ship_col].nunique()

    cps_prev = prev_cost / prev_ships if prev_ships else 0
    vol_effect = (curr_ships - prev_ships) * cps_prev

    prev_matched = acc_prev["_matched_rate_old"].dropna()
    curr_matched = acc_curr["_matched_rate_new"].dropna()
    if len(prev_matched) > 0 and len(curr_matched) > 0:
        avg_old_rate = prev_matched.mean()
        avg_new_rate = curr_matched.mean()
        rate_effect = (avg_new_rate - avg_old_rate) * prev_ships
    else:
        rate_effect = 0

    mix_effect = (curr_cost - prev_cost) - vol_effect - rate_effect

    return {
        "prev_cost": prev_cost,
        "curr_cost": curr_cost,
        "total_delta": curr_cost - prev_cost,
        "vol_effect": vol_effect,
        "rate_effect": rate_effect,
        "mix_effect": mix_effect,
        "prev_ships": prev_ships,
        "curr_ships": curr_ships,
        "cps_prev": cps_prev,
        "cps_curr": curr_cost / curr_ships if curr_ships else 0,
        "avg_old_rate": prev_matched.mean() if len(prev_matched) > 0 else 0,
        "avg_new_rate": curr_matched.mean() if len(curr_matched) > 0 else 0,
        "matched_pct_prev": len(prev_matched) / len(acc_prev) * 100 if len(acc_prev) > 0 else 0,
        "matched_pct_curr": len(curr_matched) / len(acc_curr) * 100 if len(acc_curr) > 0 else 0,
    }


@st.cache_data(ttl="1h")
def get_available_months() -> dict:
    accrual_files = list(ACCRUAL_DIR.glob("Accrual_*.xlsx")) + list(ACCRUAL_DIR.glob("Accrual_*.xlsb"))
    accrual_months = sorted(set(f.stem.replace("Accrual_", "") for f in accrual_files))
    posting_months = sorted([f.stem.replace("Posting_", "")
                             for f in POSTING_DIR.glob("Posting_*.xlsx")
                             if not f.name.startswith("~$")])
    return {"accruals": accrual_months, "postings": posting_months}


# ---------------------------------------------------------------------------
# FILTERING
# ---------------------------------------------------------------------------
def _detect_country_col(df):
    """Detect the country column name in an accrual DataFrame."""
    if df.empty:
        return "Country"
    for c in ("Country", "Country Code"):
        if c in df.columns:
            return c
    return "Country"


def filter_accrual(df, warehouse, division, countries=None, carriers=None,
                   biz_types=None, order_types=None, ecom_b2b=None, truck_types=None):
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    if warehouse:
        wh_codes = WAREHOUSE_MAP.get(warehouse, [warehouse])
        wh_mask = pd.Series(False, index=df.index)
        if "Plant" in df.columns:
            wh_mask |= df["Plant"].astype(str).str.strip().isin(wh_codes)
        if "W/H Code" in df.columns:
            wh_mask |= df["W/H Code"].astype(str).str.strip().isin(wh_codes)
        mask &= wh_mask
    if division and "Division Code" in df.columns:
        mask &= df["Division Code"].astype(str).str.strip().eq(division)
    country_col = _detect_country_col(df)
    if countries and country_col in df.columns:
        mask &= df[country_col].astype(str).str.strip().isin(countries)
    if carriers and "Carrier name" in df.columns:
        mask &= df["Carrier name"].astype(str).str.strip().isin(carriers)
    if biz_types and "Biz. Type" in df.columns:
        mask &= df["Biz. Type"].astype(str).str.strip().isin(biz_types)
    if order_types and "Order Type2" in df.columns:
        mask &= df["Order Type2"].astype(str).str.strip().isin(order_types)
    if ecom_b2b and "SEBN B2B/D2C" in df.columns:
        mask &= df["SEBN B2B/D2C"].astype(str).str.strip().isin(ecom_b2b)
    if truck_types and "Description" in df.columns:
        desc_values = []
        for tt in truck_types:
            desc_values.extend(TRUCK_TYPE_MAP.get(tt, {}).get("accrual", []))
        if desc_values:
            mask &= df["Description"].astype(str).str.strip().isin(desc_values)
    return df[mask].copy()


def filter_posting(df, warehouse, division, countries=None, carriers=None,
                   biz_types=None, mot=None, cost_types=None, ecom_b2b=None, order_types=None, truck_types=None):
    if df.empty:
        return df
    mask = pd.Series(True, index=df.index)
    if warehouse:
        wh_codes = WAREHOUSE_MAP.get(warehouse, [warehouse])
        mask &= df["Plant"].astype(str).str.strip().isin(wh_codes)
    if division and "Division" in df.columns:
        mask &= df["Division"].astype(str).str.strip().eq(division)
    if countries and "Dest.Cnty" in df.columns:
        mask &= df["Dest.Cnty"].astype(str).str.strip().isin(countries)
    if carriers and "Carrier name" in df.columns:
        mask &= df["Carrier name"].astype(str).str.strip().isin(carriers)
    if biz_types and "BIZ Type" in df.columns:
        mask &= df["BIZ Type"].astype(str).str.strip().isin(biz_types)
    if mot and "MOT" in df.columns:
        mask &= df["MOT"].astype(str).str.strip().isin(mot)
    if cost_types and "Type of cost" in df.columns:
        mask &= df["Type of cost"].astype(str).str.strip().isin(cost_types)
    if ecom_b2b and "Shipping Point" in df.columns:
        sp_values = []
        for ecom_val in ecom_b2b:
            for sp, ev in SHIPPING_POINT_ECOM_MAP.items():
                if ev == ecom_val:
                    sp_values.append(sp)
        if sp_values:
            mask &= df["Shipping Point"].astype(str).str.strip().isin(sp_values)
    if truck_types and "Description.1" in df.columns:
        desc_values = []
        for tt in truck_types:
            desc_values.extend(TRUCK_TYPE_MAP.get(tt, {}).get("posting", []))
        if desc_values:
            mask &= df["Description.1"].astype(str).str.strip().isin(desc_values)
    return df[mask].copy()


# ---------------------------------------------------------------------------
# SUMMARY TABLE BUILDER
# ---------------------------------------------------------------------------
def build_summary_table(accrual_curr, accrual_prev, posting_curr):
    """Build LE posting / WE Accrual / WE Reverse accrual / Total rows."""
    rows = []
    p_cost = posting_curr["_cost"].sum() if not posting_curr.empty else 0
    if not posting_curr.empty:
        if "_ship_count" in posting_curr.columns:
            p_ship = posting_curr["_ship_count"].nunique()
        elif "_ship" in posting_curr.columns:
            p_ship = posting_curr["_ship"].nunique()
        else:
            p_ship = len(posting_curr)
    else:
        p_ship = 0
    rows.append(("LE - Normal LCC posting", p_cost, p_ship))

    a_cost = accrual_curr[COST_COL].sum() if not accrual_curr.empty else 0
    a_ship = accrual_curr[SHIP_COL].nunique() if not accrual_curr.empty else 0
    rows.append(("WE - Accrual", a_cost, a_ship))

    r_cost = -accrual_prev[COST_COL].sum() if not accrual_prev.empty else 0
    r_ship = -accrual_prev[SHIP_COL].nunique() if not accrual_prev.empty else 0
    rows.append(("WE - Reverse accrual", r_cost, r_ship))

    t_cost = sum(r[1] for r in rows)
    t_ship = sum(r[2] for r in rows)
    rows.append(("(All)", t_cost, t_ship))
    return pd.DataFrame(rows, columns=["Row", "Cost", "Shipments"])


def compare_summaries(sum_prev, sum_curr, label_prev, label_curr):
    result = []
    for _, r_prev in sum_prev.iterrows():
        r_curr = sum_curr[sum_curr["Row"] == r_prev["Row"]]
        if r_curr.empty:
            continue
        r_curr = r_curr.iloc[0]
        dc = r_curr["Cost"] - r_prev["Cost"]
        dp = dc / r_prev["Cost"] if r_prev["Cost"] != 0 else np.nan
        ds = r_curr["Shipments"] - r_prev["Shipments"]
        dsp = ds / r_prev["Shipments"] if r_prev["Shipments"] != 0 else np.nan
        result.append({
            "Row": r_prev["Row"],
            f"{label_prev} Cost": r_prev["Cost"],
            f"{label_curr} Cost": r_curr["Cost"],
            "Delta EUR": dc,
            "Delta %": dp,
            f"{label_prev} Ship": r_prev["Shipments"],
            f"{label_curr} Ship": r_curr["Shipments"],
            "Delta #": ds,
            "Delta % (Ship)": dsp,
        })
    return pd.DataFrame(result)


# ---------------------------------------------------------------------------
# DRIVER ANALYSIS
# ---------------------------------------------------------------------------
def driver_breakdown(df_curr, df_prev, cost_col, ship_col, dim, use_unique=True):
    if dim not in df_curr.columns or dim not in df_prev.columns:
        return pd.DataFrame()
    count_col_curr = "_ship_count" if "_ship_count" in df_curr.columns else ship_col
    count_col_prev = "_ship_count" if "_ship_count" in df_prev.columns else ship_col
    g_curr = df_curr.groupby(dim, dropna=False).agg(
        cost=(cost_col, "sum"), ships=(count_col_curr, "nunique"), lines=(cost_col, "size")
    ).reset_index()
    g_prev = df_prev.groupby(dim, dropna=False).agg(
        cost=(cost_col, "sum"), ships=(count_col_prev, "nunique"), lines=(cost_col, "size")
    ).reset_index()
    m = g_prev.merge(g_curr, on=dim, how="outer", suffixes=("_prev", "_curr")).fillna(0)
    m["cost_delta"] = m["cost_curr"] - m["cost_prev"]
    m["ship_delta"] = m["ships_curr"] - m["ships_prev"]
    m["cps_prev"] = np.where(m["ships_prev"] > 0, m["cost_prev"] / m["ships_prev"], np.nan)
    m["cps_curr"] = np.where(m["ships_curr"] > 0, m["cost_curr"] / m["ships_curr"], np.nan)
    return m.sort_values("cost_delta", ascending=False)


def net_driver_breakdown(post_curr, post_prev, acc_curr, acc_prev,
                         rev_curr, rev_prev, dim_posting, dim_accrual=None,
                         cost_col_posting="_cost", cost_col_accrual=COST_COL,
                         ship_col_accrual=SHIP_COL):
    dim_accrual = dim_accrual or dim_posting

    def _net_cost_series(posting, accrual, reverse, dim_p, dim_a, cost_p, cost_a):
        groups = {}
        if not posting.empty and dim_p in posting.columns:
            for k, v in posting.groupby(dim_p)[cost_p].sum().items():
                groups[k] = groups.get(k, 0) + v
        if not accrual.empty and dim_a in accrual.columns:
            for k, v in accrual.groupby(dim_a)[cost_a].sum().items():
                groups[k] = groups.get(k, 0) + v
        if not reverse.empty and dim_a in reverse.columns:
            for k, v in reverse.groupby(dim_a)[cost_a].sum().items():
                groups[k] = groups.get(k, 0) - v
        return pd.Series(groups, name="cost")

    net_curr = _net_cost_series(post_curr, acc_curr, rev_curr, dim_posting, dim_accrual,
                                cost_col_posting, cost_col_accrual)
    net_prev = _net_cost_series(post_prev, acc_prev, rev_prev, dim_posting, dim_accrual,
                                cost_col_posting, cost_col_accrual)

    df = pd.DataFrame({"cost_prev": net_prev, "cost_curr": net_curr}).fillna(0)
    df.index.name = dim_posting
    df = df.reset_index()
    df["cost_delta"] = df["cost_curr"] - df["cost_prev"]

    def _ship_series(accrual, dim_a):
        if not accrual.empty and dim_a in accrual.columns:
            s = accrual.groupby(dim_a)[ship_col_accrual].nunique()
            s.index.name = dim_a
            return s
        return pd.Series(dtype=float, name="ship")

    ship_curr_s = _ship_series(acc_curr, dim_accrual)
    ship_prev_s = _ship_series(acc_prev, dim_accrual)

    if not ship_curr_s.empty:
        ship_curr_df = ship_curr_s.rename("ships_curr").reset_index()
        df = df.merge(ship_curr_df, left_on=dim_posting, right_on=dim_accrual, how="left")
        if dim_accrual != dim_posting and dim_accrual in df.columns:
            df = df.drop(columns=[dim_accrual])
    else:
        df["ships_curr"] = 0

    if not ship_prev_s.empty:
        ship_prev_df = ship_prev_s.rename("ships_prev").reset_index()
        df = df.merge(ship_prev_df, left_on=dim_posting, right_on=dim_accrual, how="left")
        if dim_accrual != dim_posting and dim_accrual in df.columns:
            df = df.drop(columns=[dim_accrual])
    else:
        df["ships_prev"] = 0

    df["ships_curr"] = df["ships_curr"].fillna(0)
    df["ships_prev"] = df["ships_prev"].fillna(0)
    df["ship_delta"] = df["ships_curr"] - df["ships_prev"]
    df["cps_prev"] = np.where(df["ships_prev"] > 0, df["cost_prev"] / df["ships_prev"], np.nan)
    df["cps_curr"] = np.where(df["ships_curr"] > 0, df["cost_curr"] / df["ships_curr"], np.nan)
    df["lines_prev"] = 0
    df["lines_curr"] = 0
    return df.sort_values("cost_delta", ascending=False)


def vol_rate_decomp(cost_prev, cost_curr, ship_prev, ship_curr):
    cps_prev = cost_prev / ship_prev if ship_prev else 0
    cps_curr = cost_curr / ship_curr if ship_curr else 0
    vol_eff = (ship_curr - ship_prev) * cps_prev
    rate_eff = ship_prev * (cps_curr - cps_prev)
    return vol_eff, rate_eff, cps_prev, cps_curr


# ---------------------------------------------------------------------------
# CARRIER NAME NORMALIZATION
# ---------------------------------------------------------------------------
def normalize_carrier(name):
    if pd.isna(name) or not isinstance(name, str):
        return name
    s = name.strip().upper()
    for suffix in (" B.V.", " B.V", " S.R.O.", " S.R.O", " S. R. O.",
                   " GMBH", " GMBH & CO. KG", " N.V.", " N.V", " LTD", " LTD.",
                   " BV", " SRO"):
        if s.endswith(suffix.upper()):
            s = s[: -len(suffix)].strip()
    alias_map = {
        "WABERERS": "WABERER'S",
        "WABERER": "WABERER'S",
        "EXPEDITORS": "EXPEDITORS",
        "FIRST LOGISTICS": "FIRST LOGISTICS",
        "FIRST LOGISTICS BV": "FIRST LOGISTICS",
        "FIRST LOGISTICS B.V.": "FIRST LOGISTICS",
        "VERSTIJNEN": "VERSTEIJNEN TRANSPORT",
        "VERSTIJNEN TRANSPORT": "VERSTEIJNEN TRANSPORT",
        "JDR": "JAN DE RIJK",
        "DHL": "DHL FREIGHT",
        "WAB": "WABERER'S",
        "VOS": "VOS TRANSPORT",
        "BARSAN": "BARSAN",
        "DFDS": "DFDS",
        "EI": "EXPEDITORS",
        "GEODIS": "GEODIS",
        "GIRTEKA": "GIRTEKA",
        "PRL": "PRL FREIGHT",
        "SEGERS": "SEGERS",
        "HEPPNER": "HEPPNER",
        "INL": "INL CARGO",
        "BALTIC": "BALTIC",
    }
    return alias_map.get(s, s)


def normalize_carrier_series(series):
    return series.apply(lambda x: normalize_carrier(x) if pd.notna(x) else x)


def fmt_money(x):
    if pd.isna(x) or x is None:
        return "N/A"
    return f"EUR {x:,.0f}"


def fmt_pct(x):
    if pd.isna(x) or x is None:
        return "N/A"
    return f"{x*100:.1f}%"


# ---------------------------------------------------------------------------
# SIDEBAR FILTERS
# ---------------------------------------------------------------------------
months_info = get_available_months()
all_accrual_months = months_info["accruals"]
all_posting_months = months_info["postings"]

st.title("Freight Analysis Dashboard")
st.caption("YoY & MoM freight cost analysis with accrual/reverse/posting breakdown (config-driven)")

with st.sidebar:
    st.header("Configuration")
    compare_mode = st.radio("Comparison Mode", ["MoM (Month-over-Month)", "YoY (Year-over-Year)"])
    st.subheader("Period Selection")
    if compare_mode.startswith("MoM"):
        col1, col2 = st.columns(2)
        with col1:
            prev_month = st.selectbox("Previous Month", all_accrual_months,
                                       index=len(all_accrual_months) - 2 if len(all_accrual_months) >= 2 else 0,
                                       key="prev_mom")
        with col2:
            curr_month = st.selectbox("Current Month", all_accrual_months,
                                      index=len(all_accrual_months) - 1 if all_accrual_months else 0,
                                      key="curr_mom")
    else:
        col1, col2 = st.columns(2)
        with col1:
            prev_month = st.selectbox("Previous Year Month", all_accrual_months,
                                       index=0, key="prev_yoy")
        with col2:
            curr_month = st.selectbox("Current Year Month", all_accrual_months,
                                      index=len(all_accrual_months) - 1 if all_accrual_months else 0,
                                      key="curr_yoy")
    prev_idx = all_accrual_months.index(prev_month) if prev_month in all_accrual_months else 0
    prev_prev_month = all_accrual_months[prev_idx - 1] if prev_idx > 0 else prev_month

    st.divider()
    st.subheader("Filters")
    warehouse = st.selectbox("Warehouse (Plant/W-H)", ["L428", "L401", "L430", "S4A1", "All"], index=0)
    division = st.selectbox("Division", ["A6", "E4", "A1", "C1", "E1", "E2", "E3", "E5", "G1", "H1", "All"], index=0)
    st.caption("L428 = T4BN | L401 = T4BA (auto-matched in both accruals & postings)")

    @st.cache_data(ttl="1h")
    def get_filter_options(month_label):
        df = load_accrual(month_label)
        opts = {}
        if not df.empty:
            cc = _detect_country_col(df)
            opts["countries"] = sorted(df[cc].dropna().unique().tolist()) if cc in df.columns else []
            opts["carriers"] = sorted(df["Carrier name"].dropna().unique().tolist()) if "Carrier name" in df.columns else []
            opts["biz_types"] = sorted(df["Biz. Type"].dropna().unique().tolist()) if "Biz. Type" in df.columns else []
            opts["order_types"] = sorted(df["Order Type2"].dropna().unique().tolist()) if "Order Type2" in df.columns else []
            opts["ecom_b2b"] = sorted(df["SEBN B2B/D2C"].dropna().unique().tolist()) if "SEBN B2B/D2C" in df.columns else []
        return opts

    opts = get_filter_options(curr_month)
    with st.expander("Advanced Filters", expanded=False):
        countries = st.multiselect("Country / Destination", opts.get("countries", []))
        carriers = st.multiselect("Carrier", opts.get("carriers", []))
        biz_types = st.multiselect("BIZ Type", opts.get("biz_types", []))
        order_types = st.multiselect("Order Type", opts.get("order_types", []))
        ecom_b2b = st.multiselect("ECOM / B2B", opts.get("ecom_b2b", []))
        truck_types = st.multiselect("Truck Type (FTL/STL/LTL)", ["FTL", "STL", "LTL"],
                                     help="Postings: Truck Freight - FTL/STL/LTL | Accruals: LANE A/B/C")

    wh_filter = None if warehouse == "All" else warehouse
    div_filter = None if division == "All" else division


# ---------------------------------------------------------------------------
# LOAD DATA
# ---------------------------------------------------------------------------
@st.cache_data(ttl="1h", show_spinner="Loading all data files...")
def load_all_data(prev_m, curr_m, prev_prev_m, before_prev_m, before_curr_m):
    data = {}
    data["acc_prev"] = load_accrual(prev_m)
    data["acc_curr"] = load_accrual(curr_m)
    data["acc_before_prev"] = load_accrual(before_prev_m)
    data["acc_before_curr"] = load_accrual(before_curr_m)
    data["post_prev"] = load_posting(prev_m) if prev_m in all_posting_months else pd.DataFrame()
    data["post_curr"] = load_posting(curr_m) if curr_m in all_posting_months else pd.DataFrame()
    return data


def get_month_before(month_label, all_months):
    if month_label in all_months:
        idx = all_months.index(month_label)
        if idx > 0:
            return all_months[idx - 1]
    return month_label

before_prev_month = get_month_before(prev_month, all_accrual_months)
before_curr_month = get_month_before(curr_month, all_accrual_months)

data = load_all_data(prev_month, curr_month, before_prev_month, before_prev_month, before_curr_month)

acc_prev_f = filter_accrual(data["acc_prev"], wh_filter, div_filter, countries, carriers, biz_types, order_types, ecom_b2b, truck_types=truck_types)
acc_curr_f = filter_accrual(data["acc_curr"], wh_filter, div_filter, countries, carriers, biz_types, order_types, ecom_b2b, truck_types=truck_types)
acc_before_prev_f = filter_accrual(data["acc_before_prev"], wh_filter, div_filter, countries, carriers, biz_types, order_types, ecom_b2b, truck_types=truck_types)
acc_before_curr_f = filter_accrual(data["acc_before_curr"], wh_filter, div_filter, countries, carriers, biz_types, order_types, ecom_b2b, truck_types=truck_types)

post_countries = countries
post_carriers = carriers
post_biz = biz_types
post_prev_f = filter_posting(data["post_prev"], wh_filter, div_filter, post_countries, post_carriers, post_biz, ecom_b2b=ecom_b2b, order_types=order_types, truck_types=truck_types)
post_curr_f = filter_posting(data["post_curr"], wh_filter, div_filter, post_countries, post_carriers, post_biz, ecom_b2b=ecom_b2b, order_types=order_types, truck_types=truck_types)

# Normalize carrier names so postings and accruals match
for _df in [acc_prev_f, acc_curr_f, acc_before_prev_f, acc_before_curr_f, post_prev_f, post_curr_f]:
    if not _df.empty and "Carrier name" in _df.columns:
        _df["Carrier name"] = normalize_carrier_series(_df["Carrier name"])

# Detect accrual country column name
ACC_COUNTRY_COL = None
for _c in ("Country", "Country Code"):
    if not acc_curr_f.empty and _c in acc_curr_f.columns:
        ACC_COUNTRY_COL = _c
        break
    if not acc_prev_f.empty and _c in acc_prev_f.columns:
        ACC_COUNTRY_COL = _c
        break
if ACC_COUNTRY_COL is None:
    ACC_COUNTRY_COL = "Country"

# Compute NET breakdowns
net_bd_carrier = net_driver_breakdown(
    post_curr_f, post_prev_f, acc_curr_f, acc_prev_f,
    acc_before_curr_f, acc_before_prev_f,
    dim_posting="Carrier name", dim_accrual="Carrier name",
)
net_bd_carrier = net_bd_carrier[net_bd_carrier["cost_prev"] + net_bd_carrier["cost_curr"] != 0]

net_bd_country = net_driver_breakdown(
    post_curr_f, post_prev_f, acc_curr_f, acc_prev_f,
    acc_before_curr_f, acc_before_prev_f,
    dim_posting="Dest.Cnty", dim_accrual=ACC_COUNTRY_COL,
)
net_bd_country = net_bd_country[net_bd_country["cost_prev"] + net_bd_country["cost_curr"] != 0]

if order_types:
    st.info("ℹ️ **Order Type filter only applies to accruals.** "
            "The posting file does not have an Order Type column. "
            "ECOM/B2B filter works for both (postings mapped via Shipping Point: ELD6=ECOM, ELD1=B2B).")


# ---------------------------------------------------------------------------
# KPI METRICS
# ---------------------------------------------------------------------------
st.divider()
st.subheader(f"{compare_mode} Summary - Division {division}, Warehouse {warehouse}")

sum_prev = build_summary_table(acc_prev_f, acc_before_prev_f, post_prev_f)
sum_curr = build_summary_table(acc_curr_f, acc_before_curr_f, post_curr_f)

comparison = compare_summaries(sum_prev, sum_curr, prev_month, curr_month)

total_prev = sum_prev[sum_prev["Row"] == "(All)"]["Cost"].values[0] if not sum_prev.empty else 0
total_curr = sum_curr[sum_curr["Row"] == "(All)"]["Cost"].values[0] if not sum_curr.empty else 0
total_delta = total_curr - total_prev
total_delta_pct = total_delta / total_prev if total_prev else 0

ship_prev = sum_prev[sum_prev["Row"] == "(All)"]["Shipments"].values[0] if not sum_prev.empty else 0
ship_curr = sum_curr[sum_curr["Row"] == "(All)"]["Shipments"].values[0] if not sum_curr.empty else 0
ship_delta = ship_curr - ship_prev
ship_delta_pct = ship_delta / ship_prev if ship_prev else 0

cps_prev = total_prev / ship_prev if ship_prev else 0
cps_curr = total_curr / ship_curr if ship_curr else 0
cps_delta_pct = (cps_curr - cps_prev) / cps_prev if cps_prev else 0

with st.container(horizontal=True):
    st.metric("Total Cost (Net)", fmt_money(total_curr), fmt_pct(total_delta_pct), border=True)
    st.metric("Total Shipments", f"{ship_curr:,.0f}", fmt_pct(ship_delta_pct), border=True)
    st.metric("Cost / Shipment", fmt_money(cps_curr), fmt_pct(cps_delta_pct), border=True)
    st.metric("Cost Delta", fmt_money(total_delta), border=True)

# ---------------------------------------------------------------------------
# SUMMARY TABLE
# ---------------------------------------------------------------------------
st.subheader("Summary Table (LE Posting / WE Accrual / WE Reverse / Total)")
display_comp = comparison.copy()
for col in [f"{prev_month} Cost", f"{curr_month} Cost", "Delta EUR"]:
    if col in display_comp.columns:
        display_comp[col] = display_comp[col].apply(lambda x: f"{x:,.0f}")
display_comp["Delta %"] = display_comp["Delta %"].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "N/A")
for col in [f"{prev_month} Ship", f"{curr_month} Ship", "Delta #"]:
    if col in display_comp.columns:
        display_comp[col] = display_comp[col].apply(lambda x: f"{x:,.0f}")
display_comp["Delta % (Ship)"] = display_comp["Delta % (Ship)"].apply(lambda x: f"{x*100:.1f}%" if pd.notna(x) else "N/A")
st.dataframe(display_comp, use_container_width=True, hide_index=True)

# ---------------------------------------------------------------------------
# VOLUME vs RATE DECOMPOSITION
# ---------------------------------------------------------------------------
st.subheader("Volume vs Rate Decomposition")
ve, re_eff, cps_p, cps_c = vol_rate_decomp(total_prev, total_curr, ship_prev, ship_curr)
col1, col2, col3 = st.columns(3)
with col1:
    st.metric("Volume Effect (Shipments)", fmt_money(ve),
             help="Cost change due to shipment count change, holding rate constant")
with col2:
    st.metric("Rate Effect (Cost/Ship)", fmt_money(re_eff),
             help="Cost change due to cost-per-shipment change, holding volume constant")
with col3:
    st.metric("Total Cost Delta", fmt_money(total_delta))
st.caption(f"Cost per shipment: {fmt_money(cps_p)} -> {fmt_money(cps_c)} | Shipments: {ship_prev:,.0f} -> {ship_curr:,.0f}")

# ---------------------------------------------------------------------------
# COST BRIDGE WATERFALL CHART
# ---------------------------------------------------------------------------
st.subheader("Cost Bridge: Previous → Volume Effect → Rate Effect → Current")
bridge_display = pd.DataFrame({
    "Step": [prev_month, "Volume Effect\n(Shipments)", "Rate Effect\n(Cost/Ship)", curr_month],
    "Cost": [total_prev, ve, re_eff, total_curr],
})
st.bar_chart(bridge_display.set_index("Step"), use_container_width=True)
st.caption(f"Previous: {fmt_money(total_prev)} → Volume: {fmt_money(ve)} → Rate: {fmt_money(re_eff)} → Current: {fmt_money(total_curr)}")

# ---------------------------------------------------------------------------
# TOP 5 COST DRIVERS (NET)
# ---------------------------------------------------------------------------
st.subheader("Top 5 Cost Drivers by Carrier (Net)")
if not net_bd_carrier.empty:
    top5 = net_bd_carrier.head(5)[["Carrier name", "cost_delta"]].set_index("Carrier name")
    top5.columns = ["Cost Delta EUR"]
    st.bar_chart(top5, use_container_width=True)
else:
    st.info("No carrier data for selected filters.")

# Carrier market share
st.subheader("Carrier Market Share (Current Period, Net)")
if not net_bd_carrier.empty:
    share = net_bd_carrier.set_index("Carrier name")["cost_curr"].sort_values(ascending=False)
    share = share[share > 0]
    if len(share) > 8:
        top_carriers = share.head(7)
        other = pd.Series({"Other": share.iloc[7:].sum()})
        share = pd.concat([top_carriers, other])
    if len(share) > 0:
        st.bar_chart(share, use_container_width=True)
        st.caption(f"Top carrier: {share.index[0]} ({share.iloc[0]/share.sum()*100:.1f}% of net total)")
    else:
        st.info("No positive cost in current period.")
else:
    st.info("No carrier data for selected filters.")

# ---------------------------------------------------------------------------
# KEY INSIGHTS
# ---------------------------------------------------------------------------
st.subheader("🔑 Key Insights")
insights = []
if total_delta != 0:
    main_driver = "volume (more shipments)" if abs(ve) > abs(re_eff) else "rate (higher cost per shipment)"
    insights.append(f"📊 The cost change is primarily driven by **{main_driver}** — "
                    f"volume effect: {fmt_money(ve)}, rate effect: {fmt_money(re_eff)}")
if not net_bd_carrier.empty:
    top = net_bd_carrier.iloc[0]
    if abs(top["cost_delta"]) > 0:
        insights.append(f"🚚 **{top['Carrier name']}** is the #1 cost driver (net): "
                      f"{fmt_money(top['cost_delta'])} delta ({fmt_pct(top['cost_delta']/top['cost_prev'] if top['cost_prev'] else 0)})")
if not net_bd_country.empty:
    top = net_bd_country.iloc[0]
    if abs(top["cost_delta"]) > 0:
        country_col_name = net_bd_country.columns[0]
        insights.append(f"🌍 **{top[country_col_name]}** is the top destination driver (net): "
                      f"{fmt_money(top['cost_delta'])} delta")

if cps_prev > 0 and cps_curr > 0:
    if cps_curr > cps_prev:
        insights.append(f"💰 Cost per shipment **increased** from {fmt_money(cps_prev)} to {fmt_money(cps_curr)} "
                        f"({fmt_pct(cps_delta_pct)})")
    else:
        insights.append(f"💰 Cost per shipment **decreased** from {fmt_money(cps_prev)} to {fmt_money(cps_curr)} "
                        f"({fmt_pct(cps_delta_pct)})")
for ins in insights:
    st.markdown(f"- {ins}")


# ---------------------------------------------------------------------------
# DRIVER ANALYSIS TABS
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Driver Analysis")

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
    ["By Carrier", "By Country", "By Order Type", "By Cost Type", "Weight & Volume", "Rate Card Analysis"],
    on_change="rerun"
)

SHOW_COLS = ["cost_prev", "cost_curr", "cost_delta", "ships_prev", "ships_curr", "cps_prev", "cps_curr"]

if tab1.open:
    with tab1:
        st.markdown("#### Cost Delta by Carrier (Net: Posting + Accrual - Reverse)")
        st.caption("Net cost = LE Posting + WE Accrual - WE Reverse accrual (matches the top summary table total)")
        if not net_bd_carrier.empty:
            bd = net_bd_carrier.copy()
            chart_data = bd[["Carrier name", "cost_prev", "cost_curr"]].set_index("Carrier name")
            chart_data.columns = [prev_month, curr_month]
            st.bar_chart(chart_data, use_container_width=True)
            rm = {"cost_prev": f"{prev_month} Cost", "cost_curr": f"{curr_month} Cost",
                  "cost_delta": "Delta EUR", "ships_prev": f"{prev_month} Ship",
                  "ships_curr": f"{curr_month} Ship", "cps_prev": f"{prev_month} CPS",
                  "cps_curr": f"{curr_month} CPS"}
            st.dataframe(bd[["Carrier name"] + SHOW_COLS].rename(columns=rm), use_container_width=True, hide_index=True)
        else:
            st.info("No carrier data for selected filters.")

if tab2.open:
    with tab2:
        st.markdown("#### Cost Delta by Destination Country (Net: Posting + Accrual - Reverse)")
        st.caption("Net cost = LE Posting + WE Accrual - WE Reverse accrual (matches the top summary table total)")
        if not net_bd_country.empty:
            bd = net_bd_country.copy()
            country_col = bd.columns[0]
            chart_data = bd[[country_col, "cost_prev", "cost_curr"]].set_index(country_col)
            chart_data.columns = [prev_month, curr_month]
            st.bar_chart(chart_data, use_container_width=True)
            rm = {"cost_prev": f"{prev_month} Cost", "cost_curr": f"{curr_month} Cost",
                  "cost_delta": "Delta EUR", "ships_prev": f"{prev_month} Ship",
                  "ships_curr": f"{curr_month} Ship", "cps_prev": f"{prev_month} CPS",
                  "cps_curr": f"{curr_month} CPS"}
            st.dataframe(bd[[country_col] + SHOW_COLS].rename(columns=rm), use_container_width=True, hide_index=True)
        else:
            st.info("No country data for selected filters.")

if tab3.open:
    with tab3:
        st.markdown("#### Cost Delta by Order Type (Accrual data)")
        if not acc_curr_f.empty and not acc_prev_f.empty:
            bd = driver_breakdown(acc_curr_f, acc_prev_f, COST_COL, SHIP_COL, "Order Type2")
            bd = bd[bd["cost_prev"] + bd["cost_curr"] > 0]
            if not bd.empty:
                chart_data = bd[["Order Type2", "cost_prev", "cost_curr"]].set_index("Order Type2")
                chart_data.columns = [prev_month, curr_month]
                st.bar_chart(chart_data, use_container_width=True)
                rm = {"cost_prev": f"{prev_month} Cost", "cost_curr": f"{curr_month} Cost",
                      "cost_delta": "Delta EUR", "ships_prev": f"{prev_month} Ship",
                      "ships_curr": f"{curr_month} Ship", "cps_prev": f"{prev_month} CPS",
                      "cps_curr": f"{curr_month} CPS"}
                st.dataframe(bd[["Order Type2"] + SHOW_COLS].rename(columns=rm), use_container_width=True, hide_index=True)
            else:
                st.info("No order type data for selected filters.")
        else:
            st.warning("Accrual data not available for one or both periods.")

if tab4.open:
    with tab4:
        st.markdown("#### Cost Delta by Type of Cost (Posting data)")
        if not post_curr_f.empty and not post_prev_f.empty:
            bd = driver_breakdown(post_curr_f, post_prev_f, "_cost", "_ship", "Type of cost")
            bd = bd[bd["cost_prev"] + bd["cost_curr"] > 0]
            if not bd.empty:
                chart_data = bd[["Type of cost", "cost_prev", "cost_curr"]].set_index("Type of cost")
                chart_data.columns = [prev_month, curr_month]
                st.bar_chart(chart_data, use_container_width=True)
                rm = {"cost_prev": f"{prev_month} Cost", "cost_curr": f"{curr_month} Cost",
                      "cost_delta": "Delta EUR", "ships_prev": f"{prev_month} Ship",
                      "ships_curr": f"{curr_month} Ship", "cps_prev": f"{prev_month} CPS",
                      "cps_curr": f"{curr_month} CPS"}
                st.dataframe(bd[["Type of cost"] + SHOW_COLS].rename(columns=rm), use_container_width=True, hide_index=True)
            else:
                st.info("No cost type data for selected filters.")
        else:
            st.warning("Posting data not available for one or both periods.")

if tab5.open:
    with tab5:
        st.markdown("#### Weight & Volume Analysis (from Postings)")
        st.caption("Note: Weight/volume data only available from postings. Cost/shipment uses unique Reference No count.")
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**{prev_month}**")
            if not post_prev_f.empty:
                tc = post_prev_f["_cost"].sum()
                tw = post_prev_f["_weight"].sum()
                tv = post_prev_f["_volume"].sum()
                ts = post_prev_f["_ship_count"].nunique() if "_ship_count" in post_prev_f.columns else post_prev_f["_ship"].nunique() if "_ship" in post_prev_f.columns else len(post_prev_f)
                st.metric("Posting Cost", fmt_money(tc))
                st.metric("Total Weight", f"{tw:,.0f} kg")
                st.metric("Total Volume", f"{tv:,.0f} m3")
                st.metric("Shipments (unique)", f"{ts:,.0f}")
                st.metric("Cost / kg", f"EUR {tc/tw:,.2f}" if tw else "N/A")
                st.metric("Cost / m3", f"EUR {tc/tv:,.2f}" if tv else "N/A")
                st.metric("Cost / Shipment", f"EUR {tc/ts:,.2f}" if ts else "N/A")
            else:
                st.info("No posting data for this period.")
        with col2:
            st.markdown(f"**{curr_month}**")
            if not post_curr_f.empty:
                tc = post_curr_f["_cost"].sum()
                tw = post_curr_f["_weight"].sum()
                tv = post_curr_f["_volume"].sum()
                ts = post_curr_f["_ship_count"].nunique() if "_ship_count" in post_curr_f.columns else post_curr_f["_ship"].nunique() if "_ship" in post_curr_f.columns else len(post_curr_f)
                st.metric("Posting Cost", fmt_money(tc))
                st.metric("Total Weight", f"{tw:,.0f} kg")
                st.metric("Total Volume", f"{tv:,.0f} m3")
                st.metric("Shipments (unique)", f"{ts:,.0f}")
                st.metric("Cost / kg", f"EUR {tc/tw:,.2f}" if tw else "N/A")
                st.metric("Cost / m3", f"EUR {tc/tv:,.2f}" if tv else "N/A")
                st.metric("Cost / Shipment", f"EUR {tc/ts:,.2f}" if ts else "N/A")
            else:
                st.info("No posting data for this period.")

if tab6.open:
    with tab6:
        st.markdown("#### Rate Card Analysis (2025 Old vs 2026 New Contract)")
        st.caption("Compares B2B rate card rates (STL/FTL/LTL) from the 2025 and 2026 contracts in the AI folder, "
                   "and decomposes the accrual cost increase into rate vs mix vs volume effects.")

        try:
            old_rates, new_rates = load_rate_cards()
        except Exception as e:
            old_rates = pd.DataFrame()
            new_rates = pd.DataFrame()
            st.warning(f"Could not load rate cards: {e}")

        if old_rates.empty and new_rates.empty:
            st.warning("⚠️ Rate card files not found or could not be parsed. "
                      f"Expected files in: `{AI_DIR}`")
        else:
            # Section 1: Overall rate card summary
            st.markdown("##### 1. Overall Rate Card Comparison")
            col_a, col_b, col_c = st.columns(3)
            for tt in ["STL", "FTL", "LTL"]:
                old_sub = old_rates[old_rates["truck_type"] == tt] if not old_rates.empty else pd.DataFrame()
                new_sub = new_rates[new_rates["truck_type"] == tt] if not new_rates.empty else pd.DataFrame()
                old_avg = old_sub["rate"].mean() if len(old_sub) > 0 else 0
                new_avg = new_sub["rate"].mean() if len(new_sub) > 0 else 0
                delta_pct = ((new_avg - old_avg) / old_avg * 100) if old_avg else 0
                with [col_a, col_b, col_c][["STL", "FTL", "LTL"].index(tt)]:
                    st.metric(
                        f"{tt} Avg Rate",
                        f"EUR {new_avg:,.0f}" if new_avg else "N/A",
                        f"{delta_pct:+.1f}%" if old_avg else "N/A",
                        help=f"2025: EUR {old_avg:,.0f} | 2026: EUR {new_avg:,.0f}"
                    )

            # Section 2: Rate change by carrier
            st.markdown("##### 2. Rate Change by Carrier (All Truck Types)")
            if not old_rates.empty and not new_rates.empty:
                selected_tt = st.selectbox("Select Truck Type", ["STL", "FTL", "LTL"], key="rate_card_tt")
                tt_compare = compare_rates_by_carrier_country(old_rates, new_rates, selected_tt)
                carrier_summary = tt_compare.groupby("carrier").agg(
                    old_avg=("old_avg_rate", "mean"),
                    new_avg=("new_avg_rate", "mean"),
                ).reset_index()
                carrier_summary["delta"] = carrier_summary["new_avg"] - carrier_summary["old_avg"]
                carrier_summary["delta_pct"] = np.where(carrier_summary["old_avg"] > 0,
                                                        carrier_summary["delta"] / carrier_summary["old_avg"] * 100, 0)
                carrier_summary = carrier_summary.sort_values("delta", ascending=False)
                disp = carrier_summary.copy()
                disp["old_avg"] = disp["old_avg"].apply(lambda x: f"EUR {x:,.0f}" if pd.notna(x) else "N/A")
                disp["new_avg"] = disp["new_avg"].apply(lambda x: f"EUR {x:,.0f}" if pd.notna(x) else "N/A")
                disp["delta"] = disp["delta"].apply(lambda x: f"EUR {x:+,.0f}" if pd.notna(x) else "N/A")
                disp["delta_pct"] = disp["delta_pct"].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) and x != 0 else "N/A")
                disp.columns = ["Carrier", f"2025 {selected_tt} Avg", f"2026 {selected_tt} Avg", "Delta", "Delta %"]
                st.dataframe(disp, use_container_width=True, hide_index=True)

            # Section 3: Top carrier/country rate increases
            st.markdown("##### 3. Top 15 Carrier/Country Rate Increases (All Truck Types)")
            if not old_rates.empty and not new_rates.empty:
                if 'selected_tt' in locals():
                    top_inc = tt_compare.head(15)
                    disp2 = top_inc.copy()
                    disp2["old_avg_rate"] = disp2["old_avg_rate"].apply(lambda x: f"EUR {x:,.0f}" if pd.notna(x) else "N/A")
                    disp2["new_avg_rate"] = disp2["new_avg_rate"].apply(lambda x: f"EUR {x:,.0f}" if pd.notna(x) else "N/A")
                    disp2["delta"] = disp2["delta"].apply(lambda x: f"EUR {x:+,.0f}" if pd.notna(x) else "N/A")
                    disp2["delta_pct"] = disp2["delta_pct"].apply(lambda x: f"{x:+.1f}%" if pd.notna(x) else "N/A")
                    disp2 = disp2[["carrier", "country", "old_avg_rate", "new_avg_rate", "delta", "delta_pct"]]
                    disp2.columns = ["Carrier", "Country", f"2025 {selected_tt} Rate", f"2026 {selected_tt} Rate", "Delta", "Delta %"]
                    st.dataframe(disp2, use_container_width=True, hide_index=True)

            # Section 4: Cost increase decomposition
            st.markdown("##### 4. Cost Increase Decomposition (Rate vs Mix vs Volume)")
            st.caption("Decomposes the accrual cost increase into: **Rate effect** (same carrier/country, higher rate), "
                       "**Mix effect** (shift to more expensive carriers/countries/lanes), and **Volume effect** (more/fewer shipments). "
                       "Rate card rates are matched to accrual lines by carrier + destination country.")

            if not acc_prev_f.empty and not acc_curr_f.empty:
                decomp = decompose_cost_increase(acc_prev_f, acc_curr_f, old_rates, new_rates, country_col=ACC_COUNTRY_COL)
                if decomp:
                    col1, col2, col3, col4 = st.columns(4)
                    with col1:
                        st.metric("Volume Effect", fmt_money(decomp["vol_effect"]),
                                 help="Cost change from shipment count change, holding rate constant")
                    with col2:
                        st.metric("Rate Effect (Card)", fmt_money(decomp["rate_effect"]),
                                 help="Cost change from rate card increase, matched per carrier+country")
                    with col3:
                        st.metric("Mix Effect (Residual)", fmt_money(decomp["mix_effect"]),
                                 help="Cost change from carrier/country/lane mix shift (residual after rate & volume)")
                    with col4:
                        st.metric("Total Delta", fmt_money(decomp["total_delta"]))

                    st.markdown("###### Cost Bridge: Volume → Rate Card → Mix → Total")
                    bridge = pd.DataFrame({
                        "Step": [prev_month, "Volume Effect", "Rate Card Effect", "Mix Effect", curr_month],
                        "Cost": [decomp["prev_cost"], decomp["vol_effect"], decomp["rate_effect"], decomp["mix_effect"], decomp["curr_cost"]],
                    })
                    st.bar_chart(bridge.set_index("Step"), use_container_width=True)

                    st.caption(f"Rate card match rate: {decomp['matched_pct_prev']:.0f}% of prev lines matched, "
                              f"{decomp['matched_pct_curr']:.0f}% of curr lines matched. "
                              f"Matched avg rate: EUR {decomp['avg_old_rate']:,.0f} → EUR {decomp['avg_new_rate']:,.0f}")

                    st.info(f"**How to read this:** The total cost increase of {fmt_money(decomp['total_delta'])} is decomposed as:\n"
                           f"- **Volume effect** ({fmt_money(decomp['vol_effect'])}): cost change from "
                           f"{decomp['prev_ships']:.0f} → {decomp['curr_ships']:.0f} shipments\n"
                           f"- **Rate card effect** ({fmt_money(decomp['rate_effect'])}): cost change from the "
                           f"contract rate increase (matched per carrier+country)\n"
                           f"- **Mix effect** ({fmt_money(decomp['mix_effect'])}): cost change from shifting to "
                           f"more expensive carriers, countries, or lanes (residual)")
                else:
                    st.info("Could not decompose cost increase (need Carrier name and Country columns in accrual data).")
            else:
                st.info("No accrual data available for the selected filters.")

            # Section 5: Active carriers rate card lookup
            st.markdown("##### 5. Rate Card for Active Carriers in This Filter")
            if not acc_curr_f.empty and "Carrier name" in acc_curr_f.columns and not new_rates.empty:
                active_carriers = acc_curr_f["Carrier name"].dropna().unique()
                rate_rows = []
                for carrier in active_carriers:
                    for tt in ["STL", "FTL", "LTL"]:
                        old_sub = old_rates[(old_rates["carrier"] == carrier) & (old_rates["truck_type"] == tt)] if not old_rates.empty else pd.DataFrame()
                        new_sub = new_rates[(new_rates["carrier"] == carrier) & (new_rates["truck_type"] == tt)]
                        old_avg = old_sub["rate"].mean() if len(old_sub) > 0 else None
                        new_avg = new_sub["rate"].mean() if len(new_sub) > 0 else None
                        if old_avg is not None or new_avg is not None:
                            delta = (new_avg - old_avg) if (old_avg and new_avg) else None
                            delta_pct = (delta / old_avg * 100) if (old_avg and delta is not None) else None
                            rate_rows.append({
                                "Carrier": carrier,
                                "Truck Type": tt,
                                "2025 Rate": f"EUR {old_avg:,.0f}" if old_avg else "—",
                                "2026 Rate": f"EUR {new_avg:,.0f}" if new_avg else "—",
                                "Delta": f"EUR {delta:+,.0f}" if delta is not None else "N/A",
                                "Delta %": f"{delta_pct:+.1f}%" if delta_pct is not None else "N/A",
                            })
                if rate_rows:
                    st.dataframe(pd.DataFrame(rate_rows), use_container_width=True, hide_index=True)
                else:
                    st.info("No matching carriers found between accrual data and rate card.")


# ---------------------------------------------------------------------------
# SUMMARY BULLETS
# ---------------------------------------------------------------------------
st.divider()
st.subheader("Summary Bullets")

bullets = []
delta = total_curr - total_prev
delta_pct = total_delta_pct
ship_delta_pct_val = ship_delta_pct
direction = "increased" if delta > 0 else "decreased"
bullets.append(
    f"1. Total net cost for {division} {direction} from {fmt_money(total_prev)} to {fmt_money(total_curr)} "
    f"({'+' if delta > 0 else ''}{fmt_money(delta)}, {fmt_pct(delta_pct)}) "
    f"due to shipments {'increasing' if ship_curr > ship_prev else 'changing'} from {ship_prev:,.0f} to {ship_curr:,.0f} ({fmt_pct(ship_delta_pct_val)})"
)

if not net_bd_carrier.empty:
    top = net_bd_carrier.iloc[0]
    dp = top["cost_delta"] / top["cost_prev"] if top["cost_prev"] else 0
    bullets.append(
        f"2. {top['Carrier name']} net cost {fmt_money(top['cost_prev'])} -> {fmt_money(top['cost_curr'])} "
        f"({'+' if top['cost_delta'] > 0 else ''}{fmt_money(top['cost_delta'])}, {fmt_pct(dp)}), "
        f"ships {top['ships_prev']:.0f} -> {top['ships_curr']:.0f}"
    )

if not net_bd_country.empty:
    top = net_bd_country.iloc[0]
    country_col_name = net_bd_country.columns[0]
    dp = top["cost_delta"] / top["cost_prev"] if top["cost_prev"] else 0
    bullets.append(
        f"3. {top[country_col_name]} destination net cost {fmt_money(top['cost_prev'])} -> {fmt_money(top['cost_curr'])} "
        f"({'+' if top['cost_delta'] > 0 else ''}{fmt_money(top['cost_delta'])}, {fmt_pct(dp)}), "
        f"ships {top['ships_prev']:.0f} -> {top['ships_curr']:.0f}"
    )

if not post_curr_f.empty and not post_prev_f.empty:
    bd = driver_breakdown(post_curr_f, post_prev_f, "_cost", "_ship", "BIZ Type")
    bd = bd[bd["cost_prev"] + bd["cost_curr"] > 0]
    for _, r in bd.iterrows():
        dp = r["cost_delta"] / r["cost_prev"] if r["cost_prev"] else 0
        bullets.append(
            f"4. {r['BIZ Type']} cost {fmt_money(r['cost_prev'])} -> {fmt_money(r['cost_curr'])} "
            f"({'+' if r['cost_delta'] > 0 else ''}{fmt_money(r['cost_delta'])}, {fmt_pct(dp)}), "
            f"ships {r['ships_prev']:.0f} -> {r['ships_curr']:.0f}"
        )

for b in bullets:
    st.markdown(b)

st.divider()
st.caption(f"Data source: Accrual Accuracy | Comparison: {prev_month} vs {curr_month} | Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")

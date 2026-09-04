"""
Config Loader — Shared configuration reader for all workers.
=============================================================
Loads all YAML config files from the configs/ directory.
Every worker imports this instead of hardcoding values.

This is what makes the platform scalable: change YAML, not Python.
"""
import yaml
from pathlib import Path
from functools import lru_cache

CONFIGS_DIR = Path(__file__).parent.parent / "configs"


@lru_cache(maxsize=1)
def load_warehouse_mapping():
    """Load warehouse/division/BIZ type mappings."""
    with open(CONFIGS_DIR / "warehouse_mapping.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


@lru_cache(maxsize=1)
def load_carrier_mapping():
    """Load carrier ID → name mapping (full mapping + overrides merged).
    
    All 1386 carrier mappings live in carrier_mapping.yaml.
    Overrides take priority over the carriers section.
    No parquet cache needed.
    """
    with open(CONFIGS_DIR / "carrier_mapping.yaml", "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    # Start with full mapping, then apply overrides on top
    mapping = dict(cfg.get("carriers", {}))
    mapping.update(cfg.get("overrides", {}))
    return mapping



@lru_cache(maxsize=1)
def load_charge_type_mapping():
    """Load charge type → type of cost mapping."""
    with open(CONFIGS_DIR / "charge_type_mapping.yaml", "r", encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg


@lru_cache(maxsize=1)
def load_column_structure():
    """Load target column layout for posting files."""
    with open(CONFIGS_DIR / "column_structure.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


@lru_cache(maxsize=1)
def load_formula_config():
    """Load month-end cost formula and data source definitions."""
    with open(CONFIGS_DIR / "formula.yaml", "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def get_target_columns():
    """Return the ordered list of target column names."""
    cfg = load_column_structure()
    return cfg["target_columns"]


def get_warehouse_map():
    """Return dict: file_prefix → plant_code (e.g., {'T4BA': 'L401'})."""
    cfg = load_warehouse_mapping()
    return cfg.get("warehouses", {})


def get_division_map():
    """Return dict: division_code → display_name."""
    cfg = load_warehouse_mapping()
    return cfg.get("divisions", {})


def get_carrier_map():
    """Return dict: carrier_id → carrier_name (full mapping: carriers + overrides merged).
    
    All 1386 mappings from carrier_mapping.yaml. No parquet cache needed.
    """
    return load_carrier_mapping()


def get_carrier_overrides():
    """Return dict: carrier_id → carrier_name (manual overrides only)."""
    with open(CONFIGS_DIR / "carrier_mapping.yaml", "r", encoding="utf-8") as f:
        c = yaml.safe_load(f)
    return c.get("overrides", {})



def get_charge_type_map():
    """Return dict: charge_type_str → type_of_cost."""
    cfg = load_charge_type_mapping()
    return cfg.get("charge_types", {})


def get_cost_categories():
    """Return ordered list of cost category names."""
    cfg = load_charge_type_mapping()
    return cfg.get("cost_categories", [])


def clear_cache():
    """Clear all cached configs (use after editing YAML during development)."""
    load_warehouse_mapping.cache_clear()
    load_carrier_mapping.cache_clear()
    load_charge_type_mapping.cache_clear()
    load_column_structure.cache_clear()
    load_formula_config.cache_clear()


if __name__ == "__main__":
    # Quick test: print all loaded configs
    print("=== Warehouse Mapping ===")
    print(get_warehouse_map())
    print("\n=== Division Mapping ===")
    print(get_division_map())
    print("\n=== Carrier Overrides ===")
    print(get_carrier_overrides())
    print(f"\n=== Charge Types ({len(get_charge_type_map())} mapped) ===")
    print(f"=== Target Columns ({len(get_target_columns())} columns) ===")
    print("\n=== Formula ===")
    fcfg = load_formula_config()
    print(f"  {fcfg['formula']['expression']}")
    print("\nAll configs loaded successfully!")

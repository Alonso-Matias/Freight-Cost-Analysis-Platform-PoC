"""Convert all .xlsb files to .xlsx so we don't need pyxlsb on Streamlit Cloud."""
import pandas as pd
from pathlib import Path

DATA = Path(__file__).parent / "data"

files = [
    DATA / "Accruals" / "Accrual_2025-08.xlsb",
    DATA / "rate_cards" / "rate_card_2025.xlsb",
    DATA / "rate_cards" / "rate_card_2026.xlsb",
]

for f in files:
    if not f.exists():
        print(f"SKIP (not found): {f}")
        continue
    out = f.with_suffix(".xlsx")
    print(f"Converting {f.name} -> {out.name} ...")
    xls = pd.ExcelFile(f, engine="pyxlsb")
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        for sheet in xls.sheet_names:
            df = pd.read_excel(f, sheet_name=sheet, engine="pyxlsb", header=None)
            df.to_excel(writer, sheet_name=sheet[:31], index=False, header=False)
    print(f"  Done: {out} ({out.stat().st_size / 1024 / 1024:.1f} MB)")

print("\nAll conversions done!")

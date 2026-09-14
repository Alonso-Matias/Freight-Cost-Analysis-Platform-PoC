"""Rewrite each xlsx zip container (data untouched) to force git/streamlit-cloud
to re-checkout fresh copies. Adds a zip archive comment so bytes change."""
import zipfile
import os

FILES = [
    r"data\Accruals\Accrual_2025-08.xlsx",
    r"data\Accruals\Accrual_2026-07.xlsx",
    r"data\Accruals\Accrual_2026-08.xlsx",
    r"data\Postings\Posting_2025-08.xlsx",
    r"data\Postings\Posting_2026-07.xlsx",
    r"data\Postings\Posting_2026-08.xlsx",
    r"data\rate_cards\rate_card_2025.xlsx",
    r"data\rate_cards\rate_card_2026.xlsx",
]

BASE = r"C:\Users\alonso.m\Projects\freight-cost-analysis-demo"

for rel in FILES:
    path = os.path.join(BASE, rel)
    tmp = path + ".tmp"
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
            zout.comment = b"refresh-v2"
            for item in zin.infolist():
                zout.writestr(item, zin.read(item.filename))
    os.replace(tmp, path)
    # verify still readable
    ok = zipfile.is_zipfile(path)
    print(f"{rel}: rewritten, valid={ok}, size={os.path.getsize(path):,}")

print("\nAll files rewritten with fresh zip containers.")

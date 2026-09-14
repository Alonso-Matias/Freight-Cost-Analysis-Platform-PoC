"""Check every .xlsx blob stored in Git HEAD for zip validity (BadZipFile detection)."""
import subprocess
import zipfile
import tempfile
import os

REPO = r"C:\Users\alonso.m\Projects\freight-cost-analysis-demo"

files = subprocess.run(
    ["git", "-C", REPO, "ls-files", "*.xlsx"],
    capture_output=True, text=True
).stdout.strip().split("\n")

print(f"Checking {len(files)} xlsx files in git HEAD...\n")
bad = []
for f in files:
    blob = subprocess.run(
        ["git", "-C", REPO, "cat-file", "blob", f"HEAD:{f}"],
        capture_output=True
    ).stdout
    # Write blob to temp file and test zip validity
    tmp = os.path.join(tempfile.gettempdir(), "blobcheck.xlsx")
    with open(tmp, "wb") as fh:
        fh.write(blob)
    ok = zipfile.is_zipfile(tmp)
    status = "VALID" if ok else "CORRUPT (not a zip!)"
    print(f"{f}  [{len(blob):,} bytes]  -> {status}")
    if not ok:
        bad.append(f)

print()
if bad:
    print("CORRUPT FILES IN GIT:")
    for f in bad:
        print(f"  - {f}")
else:
    print("All git blobs are valid zips.")

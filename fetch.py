#!/usr/bin/env python3
"""
Download the Census C30 time-series workbooks and decide whether a new release is available.

Usage:  python3 fetch.py            (writes new_release=true|false and release=<date> to $GITHUB_OUTPUT)

Behaviour
  1. Downloads every workbook listed in segments.json "sources" from DOWNLOAD_BASE into a temp folder.
  2. Rejects anything that is not a real .xlsx (an error page or a bot-block page fails here).
  3. Reads the release date from each workbook's footer ("Construction Spending, <Month d, yyyy>").
  4. If all files share one release date that is newer than the latest folder in vintages/,
     replaces raw/ with the new files and reports new_release=true.
     If the files are from mixed releases (Census mid-publish), reports new_release=false; the next run retries.
     If the release is already archived, reports new_release=false.
Exit code is non-zero only when a download fails or a file is invalid, so GitHub emails the repo owner.
"""
import os, re, sys, json, shutil, tempfile, zipfile, datetime, urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
DOWNLOAD_BASE = os.environ.get("C30_BASE", "https://www.census.gov/construction/c30/xlsx/")
UA = ("Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/128.0 Safari/537.36 (C30 dashboard monthly update)")

def out(key, val):
    print(f"{key}={val}")
    p = os.environ.get("GITHUB_OUTPUT")
    if p:
        with open(p, "a") as f: f.write(f"{key}={val}\n")

def release_of(path):
    import openpyxl
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).worksheets[0]
    for row in ws.iter_rows(values_only=True):
        txt = " ".join(str(x) for x in row if x)
        m = re.search(r"Construction Spending,\s+([A-Z][a-z]+ \d{1,2}, \d{4})", txt)
        if m:
            return datetime.datetime.strptime(m.group(1), "%B %d, %Y").date().isoformat()
    return None

cfg = json.load(open(os.path.join(HERE, "segments.json")))
names = sorted({f for src in cfg["sources"].values() for k, f in src.items() if k in ("nsa", "sa") and f})
tmp = tempfile.mkdtemp(prefix="c30_")
releases = {}
for n in names:
    url = DOWNLOAD_BASE + n
    dest = os.path.join(tmp, n)
    req = urllib.request.Request(url, headers={"User-Agent": UA, "Accept": "*/*"})
    try:
        with urllib.request.urlopen(req, timeout=60) as r, open(dest, "wb") as f:
            shutil.copyfileobj(r, f)
    except Exception as e:
        sys.exit(f"ERROR: download failed for {url}: {e}")
    if not zipfile.is_zipfile(dest):
        head = open(dest, "rb").read(300)
        sys.exit(f"ERROR: {url} did not return an .xlsx file (first bytes: {head[:120]!r}). "
                 "Census may be blocking automated downloads or the file moved.")
    releases[n] = release_of(dest)
    print(f"downloaded {n}: {os.path.getsize(dest):,} bytes, release {releases[n]}")

if None in releases.values():
    sys.exit(f"ERROR: release date not found in: {[k for k, v in releases.items() if v is None]}")
rels = set(releases.values())
vdir = os.path.join(HERE, "vintages")
archived = sorted(d for d in os.listdir(vdir)) if os.path.isdir(vdir) else []
latest_archived = archived[-1] if archived else ""

if len(rels) > 1:
    print(f"Files are from different releases {releases}; Census is likely mid-publish. Will retry next run.")
    out("new_release", "false"); out("release", latest_archived); sys.exit(0)
rel = rels.pop()
if rel <= latest_archived:
    print(f"No new release (latest on Census {rel}, archived {latest_archived}).")
    out("new_release", "false"); out("release", rel); sys.exit(0)

raw = os.path.join(HERE, "raw")
os.makedirs(raw, exist_ok=True)
for n in names:
    shutil.copy2(os.path.join(tmp, n), os.path.join(raw, n))
print(f"New release {rel} (previous {latest_archived or 'none'}); raw/ updated.")
out("new_release", "true"); out("release", rel)

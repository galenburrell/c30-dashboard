#!/usr/bin/env python3
"""
Build the nonresidential-buildings dataset from Census C30 "time series" workbooks.

Usage:  python3 build.py [raw_dir]          (default raw_dir = ./raw)

Inputs  (raw_dir): the workbooks named in segments.json "sources" (privtime, pubtime, privsatime, pubsatime, sltime, slsatime)
Config:            segments.json  (which columns make up each segment, and what is excluded)
Outputs (./out):
  data.json        raw layer (every source cell used, with file/sheet/cell address) + segment definitions
  derived.csv      every derived monthly value: segment x sector x adjustment x month
  lineage.csv      one row per source cell feeding a derived value (segment, component, sign, file!cell, value)
  reconciliation.csv  checks: subsegments vs parent, building total vs Census nonresidential total
  transform_log.md human-readable list of every transformation applied in this run
  revisions.csv    month-by-month changes vs the previous vintage (if one exists)
Vintages:          ./vintages/<release-date>/  copy of the raw files + manifest.json with SHA-256
"""
import sys, os, re, json, csv, hashlib, shutil, datetime
import openpyxl
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = sys.argv[1] if len(sys.argv) > 1 else os.path.join(HERE, "raw")
OUT = os.path.join(HERE, "out"); os.makedirs(OUT, exist_ok=True)
CFG = json.load(open(os.path.join(HERE, "segments.json")))
MON = {m: i + 1 for i, m in enumerate("Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split())}
log = []

def norm(h):
    """Header label normalisation: collapse Excel line breaks, drop trailing footnote digits."""
    if h is None: return None
    h = str(h).replace("_x000D_", " ").replace("\n", " ")
    h = re.sub(r"\s+", " ", h).strip()
    return re.sub(r"(?<=[A-Za-z\)])\d+$", "", h).strip()

def parse_date(s):
    m = re.fullmatch(r"([A-Z][a-z]{2})-(\d{2})([pr]?)", str(s).strip())
    if not m: return None, None
    yy = int(m.group(2)); year = 1900 + yy if yy >= 90 else 2000 + yy
    return f"{year}-{MON[m.group(1)]:02d}", (m.group(3) or "")

def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 16), b""): h.update(b)
    return h.hexdigest()

# ---------------------------------------------------------------- 1. read raw workbooks
books = {}      # (sector, adj) -> {file, sheet, sha256, headers{label: letter}, rows{month: (rownum, flag)}, vals{label: {month: v}}}
vintage = None
SRC_SECTOR = {}
for sector, files in CFG["sources"].items():
    SRC_SECTOR[sector] = files.get("sector", sector)
    for adj, fname in files.items():
        if adj not in ("nsa", "sa"): continue
        if not fname:
            log.append(f"Source `{sector}` has no {adj.upper()} workbook configured; its {adj.upper()} values are unavailable.")
            continue
        path = os.path.join(RAW, fname)
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        hdr_idx = next(i for i, r in enumerate(rows) if r and str(r[0]).strip() == "Date")
        headers = {}
        for j, h in enumerate(rows[hdr_idx]):
            if j and h is not None: headers[norm(h)] = j
        months, vals = {}, {lab: {} for lab in headers}
        for i in range(hdr_idx + 1, len(rows)):
            r = rows[i]
            if not r or r[0] is None:
                continue
            ym, flag = parse_date(r[0])
            if ym is None:
                src = " ".join(str(x) for x in r if x)
                m = re.search(r"Construction Spending,\s+([A-Z][a-z]+ \d{1,2}, \d{4})", src)
                if m:
                    v = datetime.datetime.strptime(m.group(1), "%B %d, %Y").date().isoformat()
                    if vintage and v != vintage:
                        sys.exit(f"ERROR: files are from different releases ({vintage} vs {v} in {fname})")
                    vintage = v
                continue
            months[ym] = (i + 1, flag)
            for lab, j in headers.items():
                v = r[j] if j < len(r) else None
                vals[lab][ym] = v if isinstance(v, (int, float)) else None
        books[(sector, adj)] = dict(file=fname, sheet=ws.title, sha256=sha256(path),
                                    headers={k: get_column_letter(v + 1) for k, v in headers.items()},
                                    rows=months, vals=vals)
        log.append(f"Read `{fname}` sheet `{ws.title}`: header row {hdr_idx+1}, {len(headers)} columns, "
                   f"{len(months)} monthly rows ({min(months)} to {max(months)}). SHA-256 {books[(sector,adj)]['sha256'][:16]}…")
if not vintage: sys.exit("ERROR: release date not found in workbook footers")
log.insert(0, f"Census release (vintage): {vintage}")

latest = max(books[("private", "nsa")]["rows"])
for k, b in books.items():
    if max(b["rows"]) != latest: sys.exit(f"ERROR: {b['file']} ends {max(b['rows'])}, expected {latest}")
start = CFG["window_start"]
months = sorted(m for m in books[("private", "nsa")]["rows"] if m >= start)
flags = {m: books[("private", "nsa")]["rows"][m][1] for m in months if books[("private", "nsa")]["rows"][m][1]}
log.append(f"Window: {months[0]} to {months[-1]} ({len(months)} months). Preliminary/revised flags: {flags}")

# ---------------------------------------------------------------- 2. resolve every referenced column
src = lambda c: c.get("source", c["sector"])
def need(sector, col):
    for adj in ("nsa", "sa"):
        if (sector, adj) not in books: continue
        if col not in books[(sector, adj)]["headers"]:
            sys.exit(f"ERROR: column '{col}' not found in {books[(sector, adj)]['file']}. "
                     f"Census may have renamed it; update segments.json. Available: {sorted(books[(sector, adj)]['headers'])}")

refs = set()
for s in CFG["segments"]:
    for c in s["components"]: refs.add((src(c), c["col"]))
    for ss in s["subsegments"] + s.get("breakdown", []):
        for c in ss["components"]: refs.add((src(c), c["col"]))
context_cols = [("private", "Nonresidential"), ("public", "Public Nonresidential"),
                ("private", "Total Private Construction"), ("public", "Total Public Construction")]
for x in context_cols: refs.add(x)
for x in refs: need(*x)

raw = {}   # key "sector|adj|col" -> {file, sheet, col_letter, values[], rows[]}
for (sector, col) in sorted(refs):
    for adj in ("nsa", "sa"):
        b = books.get((sector, adj))
        if b is None:
            raw[f"{sector}|{adj}|{col}"] = dict(file=None, sheet=None, letter=None, label=col,
                values=[None] * len(months), rows=[None] * len(months))
            continue
        raw[f"{sector}|{adj}|{col}"] = dict(
            file=b["file"], sheet=b["sheet"], letter=b["headers"][col], label=col,
            values=[b["vals"][col].get(m) for m in months],
            rows=[b["rows"][m][0] for m in months])
log.append(f"Extracted {len(refs)} source columns × 2 adjustments × {len(months)} months = "
           f"{len(refs)*2*len(months)} source cells.")

# ---------------------------------------------------------------- 3. derive segment series
def combine(components, adj, sector_filter=None):
    out = []
    for i in range(len(months)):
        tot, ok = 0, True
        for c in components:
            if sector_filter and c["sector"] != sector_filter: continue
            v = raw[f'{src(c)}|{adj}|{c["col"]}']["values"][i]
            if v is None: ok = False; break
            tot += c["sign"] * v
        out.append(tot if ok else None)
    return out

derived_rows, lineage_rows, recon_rows = [], [], []
def emit(seg_id, seg_name, level, components):
    for adj in ("nsa", "sa"):
        for sf in ("private", "public", None):
            if sf and not any(c["sector"] == sf for c in components): continue
            series = combine(components, adj, sf)
            for i, m in enumerate(months):
                derived_rows.append([seg_id, seg_name, level, sf or "total", adj, m, series[i]])
        for c in components:
            r = raw[f'{src(c)}|{adj}|{c["col"]}']
            if r["file"] is None: continue
            for i, m in enumerate(months):
                lineage_rows.append([seg_id, seg_name, level, adj, m, c["sector"], c["col"], c["sign"],
                                     r["file"], r["sheet"], f'{r["letter"]}{r["rows"][i]}', r["values"][i]])

for s in CFG["segments"]:
    emit(s["id"], s["name"], "segment", s["components"])
    desc = " ".join(f'{"+" if c["sign"]>0 else "−"} {src(c)}:"{c["col"]}"' for c in s["components"])
    log.append(f'Segment **{s["name"]}** = {desc.lstrip("+ ")}' + (f' — {s["note"]}' if s["note"] else ""))
    for ss in s["subsegments"]:
        emit(ss["id"], ss["name"], f'sub:{s["id"]}', ss["components"])
    if s.get("breakdown"):
        for b in s["breakdown"]:
            emit(b["id"], b["name"], f'slice:{s["id"]}', b["components"])
        parent = combine(s["components"], "nsa"); parts = [combine(b["components"], "nsa") for b in s["breakdown"]]
        for i, m in enumerate(months):
            recon_rows.append([f'{s["id"]}: segment − Σ pie slices', "nsa", m, parent[i], sum(x[i] for x in parts), parent[i] - sum(x[i] for x in parts)])
        log.append(f'  Pie breakdown of {s["name"]}: ' + "; ".join(f'{b["name"]} [{"private" if all(c["sector"]=="private" for c in b["components"]) else "public"}]' for b in s["breakdown"]) + ". Slices sum exactly to the segment.")
    if s["subsegments"]:
        # reconciliation: parent − Σ subsegments = residual ("not shown separately" by Census)
        for adj in ("nsa",):
            parent = combine(s["components"], adj)
            subs = [combine(ss["components"], adj) for ss in s["subsegments"]]
            for i, m in enumerate(months):
                ssum = sum(x[i] for x in subs)
                recon_rows.append([f'{s["id"]}: parent − Σ subsegments', adj, m, parent[i], ssum, parent[i] - ssum])
        log.append(f'  Subsegments of {s["name"]}: ' + "; ".join(ss["name"] for ss in s["subsegments"]) +
                   ". Residual (parent − Σ subs) is reported as 'Other / not shown separately'.")

# federal slice (Public − state & local) vs Census Federal column
for cat in ("Health care", "Educational"):
    if f"Federal {cat}" not in books[("public", "nsa")]["headers"]: continue
    pub = books[("public","nsa")]["vals"][f"Public {cat}"]; fed = books[("public","nsa")]["vals"][f"Federal {cat}"]
    sl = books[("statelocal","nsa")]["vals"][cat]
    for m in months:
        recon_rows.append([f"Federal {cat}: (Public − state & local) vs Census 'Federal {cat}'", "nsa", m, fed[m], pub[m] - sl[m], fed[m] - (pub[m] - sl[m])])

# building total vs Census nonresidential totals
all_comps = [c for s in CFG["segments"] for c in s["components"]]
for adj in ("nsa",):
    for sf, ctx in (("private", "Nonresidential"), ("public", "Public Nonresidential")):
        bt = combine(all_comps, adj, sf)
        ref = raw[f"{sf}|{adj}|{ctx}"]["values"]
        for i, m in enumerate(months):
            recon_rows.append([f"{sf}: buildings vs Census '{ctx}'", adj, m, ref[i], bt[i], ref[i] - bt[i]])
for e in CFG["excluded"]:
    log.append(f'Excluded {e["sector"]}:"{e["col"]}" — {e["reason"]}')
log.append("Annual and YTD figures are computed from NOT seasonally adjusted monthly values (sum of months), "
           "consistent with Census annual/YTD methodology. SAAR values are used only for month-level trend views "
           "and are never summed.")
log.append("All values are nominal (current) dollars, millions, as published. No inflation adjustment.")

# ---------------------------------------------------------------- 4. vintage archive + revisions
vdir = os.path.join(HERE, "vintages", vintage)
prev_vintages = sorted(d for d in os.listdir(os.path.join(HERE, "vintages")) if d < vintage) \
    if os.path.isdir(os.path.join(HERE, "vintages")) else []
os.makedirs(vdir, exist_ok=True)
manifest = {}
for (sector, adj), b in books.items():
    shutil.copy2(os.path.join(RAW, b["file"]), os.path.join(vdir, b["file"]))
    manifest[b["file"]] = b["sha256"]
json.dump(dict(release=vintage, latest_month=latest, sha256=manifest), open(os.path.join(vdir, "manifest.json"), "w"), indent=2)

rev_rows = []
if prev_vintages:
    pv = prev_vintages[-1]
    prev_derived = os.path.join(HERE, "vintages", pv, "derived.csv")
    if os.path.exists(prev_derived):
        old = {}
        for r in csv.DictReader(open(prev_derived)):
            if r["sector"] == "total" and r["adjustment"] == "nsa" and r["level"] == "segment":
                old[(r["segment_id"], r["month"])] = float(r["value"]) if r["value"] else None
        for r in derived_rows:
            if r[3] == "total" and r[4] == "nsa" and r[2] == "segment":
                o = old.get((r[0], r[5]))
                if o is not None and r[6] is not None and o != r[6]:
                    rev_rows.append([r[0], r[1], r[5], pv, o, vintage, r[6], r[6] - o])
        log.append(f"Revisions vs previous vintage {pv}: {len(rev_rows)} segment-months changed (see revisions.csv).")
else:
    log.append("No previous vintage archived; revision comparison starts with the next update.")

# ---------------------------------------------------------------- 5. write outputs
def wcsv(name, header, rows):
    with open(os.path.join(OUT, name), "w", newline="") as f:
        w = csv.writer(f); w.writerow(header); w.writerows(rows)
wcsv("derived.csv", ["segment_id", "segment", "level", "sector", "adjustment", "month", "value"], derived_rows)
shutil.copy2(os.path.join(OUT, "derived.csv"), os.path.join(vdir, "derived.csv"))
wcsv("lineage.csv", ["segment_id", "segment", "level", "adjustment", "month", "source_sector", "source_column",
                     "sign", "file", "sheet", "cell", "value"], lineage_rows)
wcsv("reconciliation.csv", ["check", "adjustment", "month", "reference", "computed", "difference"], recon_rows)
wcsv("revisions.csv", ["segment_id", "segment", "month", "prev_vintage", "prev_value", "new_vintage", "new_value", "change"], rev_rows)

data = dict(
    meta=dict(release=vintage, latest=latest, built=datetime.datetime.now().isoformat(timespec="seconds"),
              flags=flags, units="Millions of nominal US dollars",
              files={b["file"]: dict(sheet=b["sheet"], sha256=b["sha256"]) for b in books.values()},
              source_url="https://www.census.gov/construction/c30/historical_data.html",
              revisions=len(rev_rows), previous_vintage=prev_vintages[-1] if prev_vintages else None),
    months=months, raw=raw,
    segments=CFG["segments"], excluded=CFG["excluded"],
    context={"private": "Nonresidential", "public": "Public Nonresidential"},
    revision_rows=rev_rows,
    log=log)
json.dump(data, open(os.path.join(OUT, "data.json"), "w"), separators=(",", ":"))
with open(os.path.join(OUT, "transform_log.md"), "w") as f:
    f.write("# Transformation log\n\n" + "\n".join(f"- {l}" for l in log) + "\n")
print("\n".join(log))
print(f"\nOK  release {vintage}, latest {latest}, {len(derived_rows)} derived values, {len(lineage_rows)} lineage rows")

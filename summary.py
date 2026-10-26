#!/usr/bin/env python3
"""
Print a short plain-text summary of the current build (out/data.json + out/derived.csv).
Used by the monthly Claude task to report what changed. Also flags a stale build.

Usage: python3 summary.py
"""
import os, csv, json, datetime, collections

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, "out", "data.json")))
meta = D["meta"]
latest = meta["latest"]; LY, LM = int(latest[:4]), int(latest[5:7])
vals = collections.defaultdict(dict)
names = {}
for r in csv.DictReader(open(os.path.join(HERE, "out", "derived.csv"))):
    if r["sector"] == "total" and r["adjustment"] == "nsa" and r["value"]:
        vals[r["segment_id"]][r["month"]] = float(r["value"]); names[r["segment_id"]] = r["segment"]

def ytd(sid, y):
    ms = [f"{y}-{m:02d}" for m in range(1, LM + 1)]
    return sum(vals[sid][m] for m in ms) if all(m in vals[sid] for m in ms) else None

seg_ids = [s["id"] for s in D["segments"]]
tot = {y: sum(ytd(s, y) or 0 for s in seg_ids) for y in (LY, LY - 1)}
fmt = lambda v: f"${v/1000:,.1f}B"
pct = lambda a, b: f"{(a/b-1)*100:+.1f}%" if a and b else "n/a"

rel = datetime.date.fromisoformat(meta["release"])
age = (datetime.date.today() - rel).days
print(f"Census release: {meta['release']} | data through {latest} | flags: {meta.get('flags')}")
print(f"Previous archived release: {meta.get('previous_vintage')} | segment-months revised: {meta.get('revisions')}")
print(f"STALE: newest release is {age} days old; Census normally publishes monthly. Check the GitHub Actions log." if age > 40 else f"Release age: {age} days")
print(f"\nYTD Jan-{datetime.date(2000, LM, 1):%b} {LY} vs {LY-1} (NSA, private + public)")
print(f"  Total buildings: {fmt(tot[LY])} vs {fmt(tot[LY-1])} ({pct(tot[LY], tot[LY-1])})")
for sid in ["data_center", "health_care", "aviation", "av_terminal", "educational"]:
    if sid in vals:
        a, b = ytd(sid, LY), ytd(sid, LY - 1)
        if a is not None:
            print(f"  {names[sid]}: {fmt(a)} vs {fmt(b)} ({pct(a, b)})")
rows = meta.get("revisions") and list(csv.DictReader(open(os.path.join(HERE, "out", "revisions.csv")))) or []
if rows:
    big = sorted(rows, key=lambda r: -abs(float(r["change"])))[:5]
    print("\nLargest revisions vs previous release ($M):")
    for r in big:
        print(f"  {r['segment']} {r['month']}: {float(r['prev_value']):,.0f} -> {float(r['new_value']):,.0f} ({float(r['change']):+,.0f})")

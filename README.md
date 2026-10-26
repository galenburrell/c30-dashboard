# C30 nonresidential buildings dashboard

Tracks US nonresidential **building** construction spending from the Census Bureau's Value of Construction Put in Place (C30), with data centers, health care, aviation (incl. state & local air passenger terminals) and education isolated. Every dashboard figure traces back to a workbook cell.

## Monthly update (automatic)

1. **GitHub Actions** (`.github/workflows/update.yml`) runs every weekday at 16:30 UTC.
   `fetch.py` downloads the six workbooks from `https://www.census.gov/construction/c30/xlsx/`, checks they are real .xlsx files, and reads the release date from their footers.
   - Same release as the newest folder in `vintages/` → stops, no commit.
   - Files from mixed releases (Census mid-publish) → stops, retries next run.
   - New release → `build.py` + `render.py` run and the result is committed as `C30 release YYYY-MM-DD`.
   - Download blocked or file invalid → the run fails and GitHub emails the repo owner.
2. **Claude scheduled task** pulls this repo, and if `out/data.json` holds a newer release than the published dashboard, republishes `out/dashboard.html` to the same link and sends a summary (`summary.py`).

### If Census blocks the download
Upload the six workbooks into `raw/` through GitHub's web interface ("Add file → Upload files"). The push triggers the same build and commit.

## Files

| Path | Purpose |
|---|---|
| `segments.json` | Scope: every segment, subsegment and pie slice as a signed sum of Census columns (matched by header label), plus exclusions with reasons |
| `fetch.py` | Download + new-release detection |
| `build.py` | Parses workbooks → `out/` (data.json, derived.csv, lineage.csv, reconciliation.csv, revisions.csv, transform_log.md); archives inputs with SHA-256 in `vintages/<release>/` |
| `render.py` | Injects `out/data.json` into `template.html` → `out/dashboard.html` |
| `summary.py` | Plain-text YTD summary, revisions and a staleness warning |
| `raw/` | Workbooks of the current release |
| `vintages/<release>/` | Archived inputs + manifest (hashes) + derived values for each release; used for revision tracking |

Workbooks: `privtime`, `pubtime`, `privsatime`, `pubsatime`, `sltime`, `slsatime` (.xlsx).

## Run locally
```
pip install openpyxl
python3 fetch.py      # optional; or place the workbooks in raw/ yourself
python3 build.py raw && python3 render.py
python3 summary.py
```
The build stops if any mapped column is missing or renamed, or if the workbooks come from different releases.

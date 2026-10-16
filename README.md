# Xiaomi WOD Campaign Coverage Dashboard

Live, interactive coverage dashboard for the Xiaomi WOD campaign, hosted on Vercel.

## How it works

```
Browser (index.html) ──every 5 min──▶ /api/data (Vercel function) ──▶ SharePoint WOD.xlsx
                     ◀── compact JSON ──┘   (edge-cached 5 min)
        └── fallback: data/snapshot.json (if SharePoint is unreachable)
```

- `api/data.js` downloads the workbook from the SharePoint share link (following its cookie redirects), parses the **WOD** sheet, and returns compact JSON. The response is cached at Vercel's edge for 5 minutes (`s-maxage=300`), so SharePoint is checked at most every 5 minutes.
- `index.html` polls `/api/data` every 5 minutes (and again when the tab regains focus after that). When the data changes it re-renders and shows a "New data loaded" toast.
- **Data Updated Upto** is read from cell **A1** of the WOD sheet. **Last Synced** is when the dashboard last fetched the file.
- If SharePoint can't be reached, the header pill turns amber ("Offline snapshot") and the dashboard uses `data/snapshot.json`.
- No npm dependencies: the XLSX reader (`lib/parse-wod.js`) uses only Node's built-in `zlib`. Charts are plain SVG.

## Metric definitions

| Metric | Definition |
|---|---|
| Coverage % | stores with `WOD Covered = 1` ÷ target stores |
| VM / Non-VM | split on the `Coverage` column |
| GT / OT | split on the `Source` column |
| Element coverage % | covered stores carrying the element (qty > 0) ÷ stores targeted for that element |
| Zone chart | coverage % per `Final Zone`, grouped by `Final Region` |
| OT partner table | OT stores only, grouped by `Partner Name` |

Filters: Source, Region (`Final Region`), Zone (`Final Zone`), Covered By (`Coverage`), Element (stores where that element qty > 0).

## Deploy (Vercel via GitHub)

1. Import this repo in Vercel → Framework preset **Other**, no build command, output directory = root.
2. Optional: set the env var `WOD_SOURCE_URL` to override the SharePoint link.
3. The SharePoint link must stay shared as **"Anyone with the link"**, otherwise the function receives a sign-in page and the dashboard falls back to the snapshot.

## Refresh the fallback snapshot

```bash
node scripts/build-snapshot.js path/to/WOD.xlsx
```

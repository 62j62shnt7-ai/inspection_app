# Master Inspection Plan — Web App

A local web app for tracking inspection due dates, overdue items, and
generating yearly inspection plans — replacing the master Excel workbook.

The app uses a local database, **`inspection_plan.db`**. For a new checkout,
copy your existing database into the app directory or create one by following
the Excel import instructions below.

## Running it (on your Mac, right now)

```
cd inspection_app
python3 app.py inspection_plan.db
```

Your browser opens automatically to `http://127.0.0.1:8642`. Close the
terminal window (or Ctrl+C) to stop it.

That's it — no pip installs needed to *run* the app itself.

## Day-to-day use

- **Dashboard** — overdue items, what's due in the next 30/90 days, a
  breakdown by source sheet.
- **Assets** — search/filter everything, click any row to open, edit, or
  archive it.
- **Yearly Plan** — pick a year, get every asset with a due date in that
  year, printable.
- **Log an inspection** — inside an asset's detail view, add a dated
  inspection record; if you give it a new next-due date, that becomes the
  asset's new tracked due date automatically.
- **Export CSV** — anytime, from the top bar (optionally filtered to one
  sheet) — a safety-net snapshot you can open in Excel.

From here on, you never touch the original Excel file again — the database
is the single source of truth.

## Moving to your Windows work machine later

1. Copy the whole `inspection_app` folder (code + `inspection_plan.db`) to
   your company share folder.
2. On the Windows machine, you only need **Python** installed — no other
   packages. Most corporate Windows images already have it, or IT can
   install the official python.org build without admin rights via the
   "Install for me only" option. If IT blocks even that, ask about the
   [embeddable Python zip](https://www.python.org/downloads/windows/) —
   it's a folder you unzip, no installer/admin rights required.
3. Run the same command as above (or double-click a `.bat` file — see
   below).
4. Because the `.db` file lives on the share, this behaves like your Excel
   file does today: whoever has it open is the one editing. Avoid two
   people running the app against the same `.db` file at the same time —
   same discipline you already use with the spreadsheet.

### 1-Click Launch on Windows (Zero Installation Required)

1. Double-click **`start_app.bat`**.
2. It automatically detects any Python installed on the PC (System Python, User Python 3.10-3.13) **OR** portable embeddable Python.

#### Setting up Portable Embeddable Python (Zero Admin Rights):
If you want to run the app on a locked-down Windows PC without installing Python:
1. Download **Python Embeddable zip** from python.org (e.g. `python-3.11.x-embed-amd64.zip`).
2. Unzip it directly into a subfolder named `python` inside `inspection_app/` (so `inspection_app/python/python.exe` exists).
3. Copy the `inspection_app` folder to any USB drive or company shared folder.
4. Double-click `start_app.bat` on **any PC** — it will run instantly using the portable `python/` folder without requiring administrative rights or installer actions.

## Re-importing from Excel later

You only need this if you want to bulk-load a spreadsheet again (e.g. a
one-time cleanup pass). Needs `openpyxl`:

```
pip3 install openpyxl --break-system-packages
python3 import_excel.py "/path/to/workbook.xlsx" inspection_plan.db
```

This **overwrites** the database — anything you'd added/edited in the app
since the last import would be lost, so treat it as a reset, not a sync.

## Creating a lossless Master Register draft

The updated multi-sheet workbook can be copied with a normalized `Master
Register` sheet added. Original sheets are retained, every non-empty source
row is represented, and each register row includes its source sheet/row,
raw source values, and review flags. This generator does not modify the
source workbook or the database:

```
python3 build_master_register.py \
  "/path/to/1. Master Inspection Plan - Updated 25-10-2026.xlsx" \
  "/path/to/1. Master Inspection Plan - Register Draft.xlsx"
```

Review rows by filtering `Review Flags`. Asset IDs use the strongest source
identifier available. Rows without one use a generated ID based on the
asset name and are marked `ID_FROM_ASSET_NAME_REVIEW`. Treat the reviewed
register as the one-time migration: retain its Asset ID values and edit the
Master Register for future updates instead of regenerating it from the old
source sheets. The generator refuses to overwrite an existing output file
so it cannot erase review edits. Choose a new output filename for a new
migration draft. Import only after the mappings and flagged rows have been
checked.

## What got imported, and what to double check

- **Fully structured** (all fields mapped): Vessels & TKs, Coolers, OP
  Piping, GP Piping, Turbines Piping, Critical Assets.
- **Best-effort generic import** (data preserved, but column mapping was
  inferred by keyword-matching rather than hand-verified): OP/GP Dead Legs,
  WD-33 Piping, EPFs, GP Inlet Lines, MFDs, TLs, FLs, GL Lines. Worth
  spot-checking a few rows from each of these in the Assets tab.
- **Skipped on purpose** (reference/lookup tables, not trackable assets):
  DWG, Evaluation Criteria, Inspection Sequence, Anodes Reporting, API-574
  Tables & Pipe Schedule, HT Press Tables, Temp-Repair, FF (this last one
  was just a folder link, no data rows).
- **818 of 1,420 rows** have a clean, parseable next-due date and drive the
  dashboard/overdue logic. The rest had placeholder text in their due-date
  cells (`W/Eval`, `Never`, `OOS`, `Next SD`, etc.) — exactly as they were
  in the original sheet — and show as "No due date" until you open them and
  set a real date.

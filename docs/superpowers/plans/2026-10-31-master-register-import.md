# Master Register Import Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Generate a reviewable Master Register from every non-empty row in the updated workbook and import it safely by stable Asset ID.

**Architecture:** A standalone register generator maps known sheet layouts, retains source rows and raw values, and writes a formatted `Master Register` sheet with review flags. The app validates and previews that table, then updates assets by Asset ID while retaining history; legacy imports remain separate during transition.

**Tech Stack:** Python 3 standard library, bundled `openpyxl`, SQLite, existing vanilla JS UI.

**Spec:** [2026-10-31-master-register-design.md](../specs/2026-10-31-master-register-design.md)

## Global Constraints

- Include every non-empty source row from all ten source sheets.
- Preserve raw values and source sheet/row; flag uncertainty rather than dropping or silently correcting it.
- Internal inspections apply only to tanks and vessels; lines are OSI-only.
- Imports match by Asset ID and never silently delete missing rows.
- Draft generation and preview do not mutate the current database.
- Preserve standard-library runtime compatibility and bundled `openpyxl` support.

## Review Focus

- Banner, section, and notes rows are retained and marked non-scheduleable/reference.
- Duplicate or absent source identifiers produce unique deterministic Asset IDs and review flags.
- Status strings in date fields remain raw/status values and do not become dates.
- Unclear GL Lines rows are retained and flagged.
- Re-import with the same IDs updates mapped fields without deleting absent rows or inspection history.

---

## Files

- Create `build_master_register.py` to map all workbook sheets to a normalized register and retain raw source values.
- Modify the workbook copy only; never mutate the original input workbook.
- Create `tests/test_master_register.py` for generator mapping, flags, and round-trip retention.
- Modify `app.py` for register validation, preview API, and non-destructive Asset ID update behavior.
- Modify `static/index.html` and `static/app.js` for preview/validation and explicit apply workflow.
- Modify `import_excel.py` only if a clear adapter is needed to reuse existing persistence without invoking the legacy multi-sheet parser.
- Update `README.md` with the register contract and safe import procedure.

## Task 1: Build a full-row, review-flagged register generator

**Interfaces:** `build_register(input_path: str, output_path: str) -> dict` writes a copy containing `Master Register` and returns counts by source sheet, scheduleable status, and review flag. `map_source_row(sheet_name: str, row_num: int, headers: list[str], values: list[object]) -> dict` returns normalized fields plus raw source data and flags.

- [ ] Write tests using small in-memory/generated workbooks for every non-empty row retention, deterministic IDs, tank/vessel internal applicability, line OSI-only applicability, date/status separation, duplicate IDs, and note-row flagging.
- [ ] Run `python3 -m unittest tests.test_master_register` and confirm expected failures.
- [ ] Implement header detection and explicit mappings for all ten sheets; preserve otherwise-unmapped columns in JSON and source coordinates.
- [ ] Generate a formatted `Master Register` with frozen header, autofilter, date formats, validation lists where practical, and highlighted `Review Flags` cells.
- [ ] Run targeted tests to pass, then generate a draft from the attached workbook to a new output file in the project workspace.
- [ ] Inspect output counts and sample records; confirm source workbook untouched and every non-empty source row represented.

## Task 2: Add a non-destructive import preview and apply path

**Interfaces:** `api_preview_master_register(payload) -> dict` returns row totals, additions, updates, duplicate/invalid rows, and row-level flags without DB writes. `api_apply_master_register(payload) -> dict` updates valid rows by `asset_id`, creates absent IDs, preserves logs, and leaves unmatched existing assets unchanged.

- [ ] Add tests against a temporary SQLite database: preview causes no writes; repeated IDs are rejected; valid rows update by ID; absent rows remain; inspection log rows remain; internal fields are rejected/cleared for line assets only through explicit validation behavior.
- [ ] Run targeted tests and confirm failures before implementation.
- [ ] Implement parser/validator for the `Master Register` contract, retaining input row numbers and error messages.
- [ ] Implement preview/apply API routes and transactionally apply valid records; create a backup before writes.
- [ ] Add UI preview table with row errors and an explicit apply button; default to no writes until confirmation.
- [ ] Run targeted tests; manually review preview against generated workbook.

## Task 3: Document transition and verify end-to-end workbook handling

- [ ] Update README with register columns, status/date rules, generated IDs, preview, backup, and missing-row behavior.
- [ ] Run `python3 -m unittest tests.test_master_register` and the existing `python3 test_engineering.py`.
- [ ] Verify workbook row counts by sheet equal register source-row counts; verify all flagged rows show raw values and source coordinates.
- [ ] Verify the pre-existing `inspection_plan.db` remains byte-for-byte untouched by generation and preview.
- [ ] Review Git diff and summarize unresolved flags by sheet for user review.

# Master Register and Reliable Import Design

## Intent

Make the inspection app a dependable master register for individual assets and their OSI/internal schedules. The updated workbook remains the source material during migration, while a normalized `Master Register` sheet becomes the stable import contract.

## Goals

- Include every non-empty source row from all ten sheets in `1. Master Inspection Plan - Updated 25-10-2026.xlsx`, including explanatory notes and summary rows.
- Keep one row per scheduleable asset where possible. Preserve rows that cannot confidently be classified as assets and flag them for review rather than dropping them.
- Assign deterministic `Asset ID` values where source identifiers are absent or duplicated; retain source tags, asset numbers, sheet, and row. Where an ID depends on an asset name, flag it and retain the assigned ID in the register for every later update.
- Track OSI for all applicable assets; track internal inspections only for tanks and vessels.
- Separate real dates from statuses/hold text. Preserve source cell values in raw columns and highlight invalid/ambiguous values.
- Import by `Asset ID` with a preview and row-level errors; retain app inspection history on updates.
- Never silently delete records absent from a workbook import. Any replace/archive operation must be explicit and backed up.

## Non-goals

- Reconcile or overwrite the existing database as part of generating the first draft register.
- Infer engineering risk scores or turn free-text recommendations into engineering judgments.
- Delete original worksheets or their content.
- Decide unresolved classifications or repair source dates automatically without showing the raw value and a review flag.

## Approaches considered

1. Keep parsing all source worksheets in the app. This leaves multiple fragile layouts and parsers, so it does not provide a stable future import contract.
2. Add a normalized register sheet and import only it. This needs an initial mapping pass, but provides one predictable table and allows source worksheets to remain intact. **Selected** because the user approved it and asked that every sheet be migrated with errors highlighted.

## Register and migration behavior

The generated sheet has one header row and one row per source record. It includes `Asset ID`, `Asset Class`, `Asset Name`, `Tag`, `Asset Number`, location fields, `OSI Applicable`, OSI dates/status, `Internal Applicable`, internal dates/status, remarks, `Source Sheet`, `Source Row`, `Review Flags`, and raw source values. Extra source columns are preserved in `Source Data JSON` so no non-empty value is discarded.

All ten sheets are included. Known header rows may be mapped to standard fields. Descriptive/banners, section headings, and notes are retained as records classified `Review / Reference` when they are not scheduleable assets. Duplicate or weak identifiers, unparseable dates, conflicting fields, and uncertain class/applicability are highlighted with explicit flags. IDs use the best available source identity and collision handling; weak/name-derived IDs are flagged. The first register is a one-time migration artifact. Later updates use its stored Asset IDs and do not regenerate IDs from the legacy source sheets.

Tanks and vessels are eligible for internal inspections. Pipes, flowlines, trunklines, and other line assets are OSI-only. Classification not established by a reliable sheet/row cue is flagged for review rather than guessed. Date columns contain only parsed dates; status/hold fields preserve normalized status plus original text.

## Import behavior

The app validates the `Master Register` header contract and shows a preview before applying changes. It reports new, matched, invalid, and duplicate IDs, and includes row references. Update mode matches by `Asset ID`, updates register-owned fields, and keeps inspection logs and app-maintained history. Missing rows are not deleted automatically. The legacy multi-sheet import path remains available during transition but is clearly labeled as legacy; it is not silently used for a Master Register workbook.

## Constraints and calls

- Keep the app compatible with its existing standard-library runtime and bundled `openpyxl`.
- Preserve all source sheet cells and use the attached updated workbook as input.
- Draft generation and import preview must not mutate the current database.
- Workbook review errors, name-derived IDs, and future update dates are visible in a `Review Flags` column with conspicuous conditional formatting and filters.
- The initial draft may contain unresolved issues; flagged data is still retained for later resolution.

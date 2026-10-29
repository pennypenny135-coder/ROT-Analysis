# Record of Transaction Summary 交易紀錄摘要

A 100% client-side web app for processing property sales "Record of
Transaction" Excel workbooks. Upload a `.xlsx` / `.xlsm` file containing
`PL`, `SA` and `ROT` sheets, and the app will:

- Cross-reference the three sheets by **BFF** (Block & Flat & Floor)
- Auto-fill SA Saleable Area / No. of Rooms and ROT No. of Rms / ROT 呎價
- Build 8 fixed Summary tables using native Excel formulas
  (`AVERAGEIF`, `COUNTIFS`, `MINIFS`, `MAXIFS`, `AVERAGEIFS`, …)
- Provide a cross-sheet Pivot Builder (row / column / value / aggregation / filters)
- Flag data problems both in the web UI and as red `ERROR` cells in the
  downloaded workbook
- Never send your file anywhere — everything runs in your browser.

Built with React 19 + TypeScript + Tailwind CSS v4 + ExcelJS, bundled with
Vite (single-file output).

## Running locally 本機執行

```bash
npm install
npm run dev       # start the dev server (http://localhost:5173)
npm run build     # production build -> dist/
npm run preview   # preview the production build locally
```

No environment variables, backend, or database are required — the app is
fully static and client-side.

## Deploying to Vercel 部署至 Vercel

1. Push this project to a Git repository (GitHub / GitLab / Bitbucket).
2. In Vercel, "Import Project" and select the repository.
3. Framework preset: **Vite** (or "Other"). Build command: `npm run build`.
   Output directory: `dist`.
4. Deploy — no environment variables are needed since there is no backend.

You can also deploy the `dist/` folder directly with the Vercel CLI:

```bash
npm run build
npx vercel deploy --prebuilt
```

## Excel template format Excel 範本格式

Click **"Download Sample Template 下載範本"** in the app to get a fully
worked example (12 PL units, 2 Tenders, 9 ROT transactions) that exercises
every rule below.

### Sheet 1 · PL

| Col | Field | Notes |
|---|---|---|
| A | Block | |
| B | Floor | |
| C | Flat | |
| D | BFF | May be a formula, e.g. `=A2&C2&B2`. The app reads the **calculated value**, not the formula text. |
| E | 實用面積 Saleable Area | Free text, e.g. `45.664 ( 492 )\n露台 Balcony: 2.020 ( 22 )`. The app extracts the number inside the **first** `(...)` group (tolerating spaces) — here `492`. If it cannot be extracted, the row is flagged `ERROR`. |
| F | Price ($) | |
| G | 呎價 Price per sq. ft. | Not used for Discounted 呎價 calculations. |
| H | Max Discount (%) | Accepts a decimal fraction (`0.05`) or a percentage string (`"5%"`). |
| I | Discounted 呎價 | Used directly — the app does **not** re-derive it from Column G. |
| J | Rms | `0` = studio, up to `5`. |
| K | PL No. | Required — blank values are flagged `ERROR`. |

### Sheet 2 · SA

| Col | Field | Notes |
|---|---|---|
| A | BFF | Same format as PL. |
| B | SA No. | `1` / `2` / `3` = PL unit. `1T` / `2T` / `3T` = Tender. Normalized by trimming spaces, removing internal spaces, and upper-casing the trailing `T` (`" 1t "` → `1T`). |
| C | Saleable Area (sq. ft.) | **Auto-filled** from PL for non-tender rows (never overwritten for Tender rows — enter manually). |
| D | No. of Rooms | Same auto-fill rule as Column C. |

### Sheet 3 · ROT

| Col | Field | Notes |
|---|---|---|
| A–C | Date of PASP / ASP / termination | Accepts an Excel date cell **or** `DD-MM-YYYY` text. |
| D–F | Block / Floor / Flat | |
| G | BFF | May be a formula, e.g. `=D2&F2&E2`. Calculated value is used. |
| H | Transaction Price | Required numeric value. |
| I | Terms of Payment | Free text. |
| J | (unused) | |
| K | No. of Rms | **Auto-filled** from the matching SA row. |
| L | ROT 呎價 | **Auto-computed** as `Transaction Price ÷ SA Saleable Area (sq.ft.)`. |

Tender vs. PL classification for ROT rows always comes from the matching
**SA No.** (Column B of SA), never from the free-text "Terms of Payment".

## BFF format BFF 格式說明

`BFF = Block & Flat & Floor` (concatenated in that order, as plain text),
e.g. `Block=2, Flat=B, Floor=9` → `BFF = "2B9"`.

The app normalizes BFF values for **matching only** (trims whitespace and
upper-cases), but always preserves and displays the original text.

## How the Summary is calculated 計算方法

The `Summary` sheet is deleted and rebuilt on every export. A small hidden
helper column (`ROT!M`, "SA No. (Join Helper)") is added so that Summary
formulas can tell whether a given ROT row belongs to a PL or a Tender unit
— it looks up the matching SA No. via `VLOOKUP` and normalizes it
(uppercase, no spaces). Tender rows are identified as any SA No. containing
the letter `T`.

| # | Table | Formula approach |
|---|---|---|
| 1 | PL: Avg Discounted Price | `AVERAGEIF(PL!K, PL No., PL!I)` per distinct PL No. |
| 2 | Tender: Avg Discounted Price | `AVERAGEIF(ROT!M, "1T", ROT!L)` per distinct Tender label |
| 3 | Overall: Avg Discounted Price | Weighted combination of `SUM/COUNT(PL!I)` and `SUMIF/COUNTIF(ROT!M,"*T*",ROT!L)` |
| 4 | PL-Units Sold | `COUNTIFS` on SA (Released) / ROT (Sold, PASP present) grouped by Rooms, non-tender only |
| 5 | Tender-Units Sold | Same as (4) but filtered to Tender (`SA!B` contains `T`) |
| 6 | PL-Price | `MINIFS` / `MAXIFS` / `AVERAGEIFS` on `ROT!L` grouped by Rooms, non-tender only |
| 7 | Tender-Price | Same as (6), Tender only |
| 8 | Overall-Price | `MIN` / `MAX` / `AVERAGE` of all of `ROT!L` |

All monetary / 呎價 values are formatted to 2 decimal places; take-up rates
use a `%` number format. Headers are bold with a dark fill; totals rows are
shaded.

## Pivot Builder 樞紐分析

The Pivot Builder joins PL + SA + ROT in-memory (by BFF) into one flat
dataset. You can choose any Row / Column / Value field from the three
sheets, pick an aggregation (Sum / Average / Count / Min / Max), and add
filters (category multi-select, or date-range for date fields). Results are
shown live in the browser, and can be exported as a new `Pivot` sheet in
the downloaded workbook (the PL / SA / ROT / Summary sheets are preserved).

## Known limitations 已知限制

- Only one workbook is processed at a time; re-uploading replaces the
  previous session's data.
- Duplicate-BFF detection is applied to **PL** and **SA** (each BFF should
  be unique there). It is intentionally *not* applied to ROT, since a unit
  can legitimately appear more than once over time (e.g. terminated then
  re-sold) — the spec asks users to delete terminated rows themselves.
- Summary formulas assume the sheets are named exactly `PL`, `SA`, `ROT`
  (matched case-insensitively); if any sheet is missing the app will show
  an error instead of generating a Summary.
- A small helper column (`ROT!M`) is added to support Tender/PL detection
  in Summary formulas. It's clearly labelled "SA No. (Join Helper)" and
  does not alter any of the original A–L columns.
- Very large workbooks (tens of thousands of rows) may be slower since all
  parsing/formula generation happens in the browser's main thread.
- Cell formatting/styles that exist only in advanced Excel features (e.g.
  pivot tables, macros, conditional formatting rules) in the *original*
  file are preserved only as far as the underlying `exceljs` library
  supports when round-tripping a workbook.

## Project structure 專案結構

```text
src/
  lib/excel/
    types.ts            shared TypeScript types
    normalize.ts         pure normalization helpers (area/discount/date/SA No.)
    bff.ts                BFF build + normalize-for-matching
    readWorkbook.ts       parses PL / SA / ROT sheets from an ExcelJS workbook
    lookup.ts             cross-sheet BFF join, auto-fill, validation
    constants.ts          styling palette + pivot field metadata
    summary.ts            builds the 8 fixed Summary tables (Excel formulas)
    pivot.ts              cross-sheet join + pivot aggregation (pure functions)
    exportWorkbook.ts      autofill/error-marking + Summary/Pivot sheet writers + download
    sampleTemplate.ts      generates the fictional sample workbook
  components/
    UploadPanel.tsx
    StatsPanel.tsx
    ErrorPanel.tsx
    PivotBuilder.tsx
    ThemeToggle.tsx
  App.tsx                 page composition / state orchestration
```

All business logic in `lib/excel/*` is written as pure, unit-testable
functions with no DOM or React dependency.

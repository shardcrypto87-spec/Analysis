# Yearly Refresh via Power Query — Setup Guide

**Status: UNTESTED in real Excel** (no real Excel in this environment), but
the approach it automates — reading one consolidated CaseWare export, one tab
per entity — is proven: this is exactly how the FY2026 data in this workbook
was added, done by hand in Python against the real file
`Sandhu_G__2026.xlsx`, with every result cross-checked against
already-validated numbers before being trusted. Treat the first Power Query
refresh as a trial anyway: validate its output the same way (see "How to
validate" below) before trusting it for real work.

## What this does

Reads every tab of **one consolidated CaseWare export file** — the single
workbook your accountant can prepare each year-end with one tab per entity
(confirmed shape, from `Sandhu_G__2026.xlsx`: 16 tabs, each laid out exactly
like the old individual WTB files) — applies the exact same categorization
rules used to build this workbook (GIFI code lookup → Map No fallback → the
known account-level exceptions), and produces a table shaped exactly like
`Fact_TrialBalance` — so once it's wired up, everything downstream (Entity
Register, Group Summary, Exec Summary, Tax Position) keeps working unchanged.

This replaces the older approach of gathering 16 separate `*_WTB.xlsx`
files into a folder — one file is simpler, and it can't end up with a
mismatched or missing company. If your accountant instead gives you 16
separate files in some year, the folder-based version of this code (used
before FY2026) is still in this project's git history.

It does **not** trust the file's own printed date column headers (e.g.
"Final: 2025-03-31") — confirmed stale by a year in both the FY2026 example
file and the consolidated file itself (the column labelled "2025-03-31" was
actually FY2026 for 15 of its 16 tabs). It reads by column **position**
instead (same approach the Python pipeline uses), and you tell it which
fiscal year the "Final" column represents via one parameter.

It also does **not** trust a tab's name to be the right entity just because
it's sitting in the file. FY2026 caught a real example: one tab
("Sheet16") was never renamed to its usual entity code, and turned out —
after manually inspecting its account names — to be Sandhu & Sandhu Enr.
Worse, its data was a byte-for-byte duplicate of last year's file: not
refreshed at all, despite sitting inside the new year's export. Excel
Power Query can't detect "this looks like leftover data" on its own — see
the extra validation step 5 below, which exists specifically because of
this.

## What it can't do

- Catch a **new kind of anomaly** next year that isn't one of the 6 known
  exceptions (e.g. a different account miscoded with a wrong GIFI code in a
  different entity). It will categorize confidently and wrongly. The
  `Uncategorized_Check` query below only catches accounts that don't match
  *any* rule — not accounts that match the *wrong* rule.
- Handle a **new entity** (a 17th company) automatically — a file for one
  won't be silently miscategorized (it's caught by
  `Unmapped_FilePrefix_Check`, see below), but you do need to add a row to
  `PQ_Dim_FilePrefix` and `Dim_Entity` yourself before its data will flow
  through.
- Pull Taxprep data (SBD, GRIP, Income Tax cross-checks) — this only rebuilds
  `Fact_TrialBalance` from WTB files, same as the Python pipeline's WTB half.

## One-time setup, in Excel

1. **Get this year's consolidated file** from your accountant — one workbook,
   one tab per entity, e.g. `Sandhu_G__2027.xlsx`. Save it somewhere fixed,
   e.g. `C:\Sandhu\Sandhu_G__2027.xlsx`.

2. Open this workbook in Excel. Go to **Data → Get Data → Launch Power Query
   Editor** (or **Data → Queries & Connections → New Query → Blank Query**).

3. In the Power Query Editor: **Home → New Source → Blank Query**. Then
   **Home → Advanced Editor**, delete whatever's there, and paste the whole
   contents of `Fact_TrialBalance_Refresh.pq` (next to this file). Rename the
   query to `Fact_TrialBalance_Refresh` (right-click it in the Queries pane →
   Rename).

4. **Edit the two lines at the top of the pasted code** before you load it:
   ```
   FilePath  = "C:\Sandhu\Sandhu_G__2027.xlsx",
   CurrentFY = "FY2027",
   ```
   Point `FilePath` at the file from step 1, and set `CurrentFY` to the
   fiscal year label the "Final" column in this year's file represents —
   don't trust the column's own printed date; confirm it the same way this
   session did, by cross-checking against a figure you already trust (see
   "How to validate", steps 3 and 5 below).

5. Also paste `Uncategorized_Check.pq` and `Unmapped_FilePrefix_Check.pq` as
   two more blank queries (same steps, matching names). These are your
   safety net — see below. Update the `FilePath` line in each to match.

6. **Load `Fact_TrialBalance_Refresh`**: Home → Close & Load To... → Table →
   Existing worksheet → point it at the `Fact_TrialBalance` sheet's `A1`
   cell, **replacing the existing table**. (Excel will ask to replace the
   existing connection/table — confirm.) Load the two check queries each to
   their own new sheet.

## How to validate before trusting it

Before relying on a refresh for real work:

1. Check **`Unmapped_FilePrefix_Check`** is empty. Anything listed there is
   a tab whose name wasn't recognized — its data was silently left out of
   `Fact_TrialBalance_Refresh` entirely, not guessed into the wrong entity.
   Usually means: a new company (add it to `PQ_Dim_FilePrefix` and
   `Dim_Entity`), or CaseWare left a tab with its default name (e.g.
   "Sheet16") instead of renaming it — open that tab, read its account
   names to work out which entity it actually is, then add/fix the row.
2. Check **`Uncategorized_Check`** is empty (or only contains genuinely
   $0 accounts). Anything else there is real money not sitting anywhere in
   your P&L — investigate before proceeding. A brand-new account with no
   GIFI code your mapping table recognizes (this happened for real in
   FY2026 — a $300,000 one-time "Gain on sale") needs a judgment call on
   which category it belongs in, not an automatic guess.
3. Open **02 Profitability**, pick an entity you can cross-check against its
   real financial statements or last year's confirmed figure, and confirm
   Net Income matches.
4. Check **01 Exec Summary**'s Group Revenue / Net Income against a rough
   expectation (e.g. "about the same as last year, not 10x off").
5. **For every tab, confirm it's actually new data, not a leftover copy of
   last year's file.** This is the step FY2026 proved necessary: one tab
   (Sandhu & Sandhu Enr.) was a byte-for-byte duplicate of the prior year's
   standalone file — same numbers, every account, every column — simply
   never refreshed before being dropped into this year's export. None of
   the other checks above catch this (the tab has a valid name, every
   account categorizes fine, the totals are internally consistent — they're
   just last year's totals). The tell: that entity's new "Prior" column
   (which should be last year's confirmed figure) matches last year's
   *Prior* column too, not last year's *Final* — i.e. everything is shifted
   one year further back than it should be. If you have last year's
   confirmed Net Income for each entity handy, spot-check at least the
   entities you don't check every year anyway in step 3.
6. Only once all six look right, treat the refreshed numbers as reliable.

## Every year after that

Repeat setup steps 1 and 4 (get the new file, update the two parameter lines)
only: get the new year's consolidated file, then update `FilePath` (all three
queries) and `CurrentFY` (the two that have it — `Fact_TrialBalance_Refresh`
and `Uncategorized_Check`) via Home → Advanced Editor on each query. Then
**Data → Refresh All**. Re-run the 6-point validation checklist above every
time — don't skip it just because
it worked last year.

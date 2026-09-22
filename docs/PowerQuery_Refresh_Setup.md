# Yearly Refresh via Power Query — Setup Guide

**Status: UNTESTED.** I have no real Excel in this environment, so this M code
has never actually been run. Treat the first refresh as a trial: validate its
output against the numbers already in this workbook (see "How to validate"
below) before trusting it for real work. If it disagrees with what's already
here, the manual Python pipeline (what produced this workbook) is still the
source of truth.

## What this does

Reads every `*_WTB.xlsx` file in a folder you choose, applies the exact same
categorization rules used to build this workbook (GIFI code lookup → Map No
fallback → the 6 known account-level exceptions), and produces a table shaped
exactly like `Fact_TrialBalance` — so once it's wired up, everything
downstream (Entity Register, Group Summary, Exec Summary, Tax Position) keeps
working unchanged.

It does **not** trust the WTB's own printed date headers (e.g. "Final:
2025-03-31") — those can be stale by a year, as the FY2026 example file for
Restaurant Indiarosa showed. It reads by column **position** instead (same
approach the Python pipeline uses), and you tell it which fiscal year the
"Final" column represents via one parameter.

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

1. **Create a folder** for this year's WTB files, e.g. `C:\Sandhu\WTB_2026\`.
   Put all 16 (or 17, 18...) `*_WTB.xlsx` files in it — same naming pattern
   as this year's (`<FilePrefix>_<year>_WTB.xlsx`).

2. Open this workbook in Excel. Go to **Data → Get Data → Launch Power Query
   Editor** (or **Data → Queries & Connections → New Query → Blank Query**).

3. In the Power Query Editor: **Home → New Source → Blank Query**. Then
   **Home → Advanced Editor**, delete whatever's there, and paste the whole
   contents of `Fact_TrialBalance_Refresh.pq` (next to this file). Rename the
   query to `Fact_TrialBalance_Refresh` (right-click it in the Queries pane →
   Rename).

4. **Edit the two lines at the top of the pasted code** before you load it:
   ```
   FolderPath = "C:\Sandhu\WTB_2026\",
   CurrentFY  = "FY2026",
   ```
   Point `FolderPath` at the folder from step 1, and set `CurrentFY` to the
   fiscal year label the "Final" column in this year's files represents.

5. Also paste `Uncategorized_Check.pq` and `Unmapped_FilePrefix_Check.pq` as
   two more blank queries (same steps, matching names). These are your
   safety net — see below. Update the `FolderPath` line in each to match.

6. **Load `Fact_TrialBalance_Refresh`**: Home → Close & Load To... → Table →
   Existing worksheet → point it at the `Fact_TrialBalance` sheet's `A1`
   cell, **replacing the existing table**. (Excel will ask to replace the
   existing connection/table — confirm.) Load the two check queries each to
   their own new sheet.

## How to validate before trusting it

Before relying on a refresh for real work:

1. Check **`Unmapped_FilePrefix_Check`** is empty. Anything listed there is
   a file whose company wasn't recognized — its data was silently left out
   of `Fact_TrialBalance_Refresh` entirely, not guessed into the wrong
   entity. Usually means: a new company (add it to `PQ_Dim_FilePrefix` and
   `Dim_Entity`), or a renamed/misspelled file.
2. Check **`Uncategorized_Check`** is empty (or only contains genuinely
   $0 accounts). Anything else there is real money not sitting anywhere in
   your P&L — investigate before proceeding.
3. Open **02 Profitability**, pick an entity you can cross-check against its
   real financial statements or last year's confirmed figure, and confirm
   Net Income matches.
4. Check **01 Exec Summary**'s Group Revenue / Net Income against a rough
   expectation (e.g. "about the same as last year, not 10x off").
5. Only once all four look right, treat the refreshed numbers as reliable.

## Every year after that

Repeat steps 1 and 4 only: drop the new year's files in a new folder, then
update `FolderPath` (all three queries) and `CurrentFY` (the two that have
it — `Fact_TrialBalance_Refresh` and `Uncategorized_Check`) via Home →
Advanced Editor on each query. Then **Data → Refresh All**. Re-run the
5-point validation checklist above every time — don't skip it just because
it worked last year.

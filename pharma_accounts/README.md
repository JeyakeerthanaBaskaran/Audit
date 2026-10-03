# Supplier Accounts (Streamlit + SQLite)

Builds the year-end supplier accounts sheet by comparing several Excel files,
and exports it as PDF or Excel.

## Setup (once)
    pip install -r requirements.txt
If pip fails with an SSL certificate error, use:
    python -m pip install --use-feature=truststore -r requirements.txt

## Run
Double-click `run_app.bat`, or:
    streamlit run app.py
It opens in the browser at http://localhost:8501

To try it with made-up data: `python make_sample_data.py` creates 3 example
files in `sample_data/`.

## Pages
1. **Excel files** – upload many Excel files at once; every sheet in them is
   listed and saved (choose which, their names, and the headings row).
2. **Filter & compare** – pick a sheet and filter it with the ▼ on each column
   heading (like Excel: sort, tick values, search, or a condition). Then check a
   column against another sheet so only rows found (or not found) there stay.
   Click column headings to choose the columns to keep, and save them as a new
   sheet (marked ⭐), or add them to a sheet you made before — rows already in
   it can be skipped, and an "Added on" date can be filled in.
3. **Parameters** – the columns of the Main accounts sheet. Add, rename,
   reorder, change type (Text / Money / Number / Date) or delete.
4. **Main accounts** – view, search, edit, add by hand, or import from Excel.
   When importing, file columns are matched to parameters automatically by name
   (ignoring capitals, spaces, symbols); unmatched columns can be skipped or
   turned into a new parameter. "Update" mode matches rows by a key such as
   Supplier Code; blank cells never overwrite saved values.
5. **Year-end sheet** – pick the main list, compare with other files by a key
   column, bring in only the columns you need, add filters (e.g. *Account End ·
   date is in year · 2026*, or *Amount · differs from column · Paid*), choose
   which columns to show and total, then download PDF/Excel or save the rows
   into Main accounts.
6. **History & backup** – change log, and download of the database file.

## Accuracy
- Money is stored as exact text and added up with Python `Decimal` (no float
  rounding errors). Totals are shown to 2 decimals with Indian grouping.
- Cells that are not valid numbers/dates are flagged before import, and are
  never silently dropped from totals (a warning says how many were left out).
- Repeated keys and rows not found in a compared file are reported as warnings.
- Each import is saved in one transaction: it either fully succeeds or
  changes nothing.

## Files
- `app.py` – the screens
- `db.py` – SQLite storage (`accounts.db` is created on first run)
- `utils.py` – Excel reading, parsing, matching, comparing, filters
- `filter_grid.py` – the Excel-style table with ▼ filters in the headings
- `pdf_export.py` – PDF layout
- `accounts.db` – ALL the data. Back it up.

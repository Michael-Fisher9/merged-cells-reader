# merged-cells-reader

Turns "presentation" spreadsheets, the kind built to be printed rather than analyzed, into one long table where every number knows which cell it came from.

The test case is three tables from the Census Bureau's Annual Retail Trade Survey: retail sales, inventories and operating expenses by kind of business. They have most of the usual problems:

- a title and notes stacked above the table in merged cells
- headers two rows deep, with "Annual payroll" merged over "Amount" and "Percent of total"
- labels padded with dot leaders (`Automobile dealers ………………...`) and indented with spaces, inconsistently
- footnote numbers glued onto words (`GAFO1`, `fringe benefits2`) and year headers flagged as revised (`2022r`)
- letters where numbers should be (`S`, `D`, `Z`), explained in a legend at the bottom
- footnotes and notes below the data that look like rows

## Run it

```
python -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python mcreader.py data
.venv/bin/python examples/quick_look.py
```

```
Detailed operating expenses of retailers (2022): 936 numbers out of 936
End-of-year retail inventories (2022 benchmarked): 961 numbers out of 961
Annual retail sales by kind of business (2022 benchmarked): 2139 numbers out of 2139
4783 rows -> output/2026-10-29/tidy.csv
```

## What comes out

`output/<date>/tidy.csv`, one row per cell:

| file | naics | kind_of_business | level | measure | measure_flag | value | code | source_row | source_col |
|---|---|---|---|---|---|---|---|---|---|
| Detailed operating expenses… | 441 | Motor vehicle and parts dealers | 1 | Annual payroll \| Amount | | 114805 | | 7 | 4 |
| Annual retail sales… | 4411 | Automobile dealers | 2 | 2022 | r | 1345925 | | 9 | 3 |
| Detailed operating expenses… | 443 | Electronics and appliance stores | 1 | Expensed equipment \| Amount | | | S | 11 | 14 |

and `output/<date>/qa/`:

- `reconciliation.csv`: numbers found in the sheet and numbers written out, per file
- `suppressed_cells.csv`: every cell that held a code instead of a number, with what the code means
- `skipped_rows.csv`: blank rows inside the table
- `notes.txt`: the title, which rows were the header, and every note and footnote

`config/tables.csv` names each file and, if a sheet needs it, which tab and how many ID columns it has. A file not listed there gets the defaults.

## Why it's built this way

- **Merged ranges come from the sheet's XML, not openpyxl.** openpyxl only reports merged cells when it loads the whole workbook into memory; its streaming mode doesn't see them. The ranges are sitting in the XML as `<mergeCell ref="D4:E4"/>`, so I read them with `zipfile` and stream the values. These files are small, but on a 150k-row export with a hundred thousand merged ranges this was the difference between seconds and several minutes.
- **No forward-fill.** The usual fix for merged headers is to fill blanks from the left. But a blank header cell is sometimes part of a merge and sometimes really blank, and forward-fill can't tell which. Copying each merged value across its actual range can.
- **Level comes from the NAICS code, not the indentation.** The labels are indented to show the hierarchy, but with three spaces in one place and four in another. Code length (3 digits a subsector, 4 a group, 5 an industry) is reliable.
- **The run stops before writing if anything doesn't reconcile.** Every number in the table has to appear exactly once in the output, and every letter code has to be in the sheet's own legend. A file that half-worked is worse than one that failed, because nobody looks twice at it.
- **Outputs go in a dated folder.** Re-running next year doesn't overwrite this year, and the QA files sit next to the data they describe.

## Limits

- It expects ID columns on the left and values to the right. A table turned sideways needs `id_cols` set, or a different reader.
- Header detection is a rule of thumb (first row with three filled cells, until a row has numbers in it). If a header row has numbers in it, set the header rows by hand.

Data: U.S. Census Bureau, Annual Retail Trade Survey (public domain). `get_data.sh` downloads it again.

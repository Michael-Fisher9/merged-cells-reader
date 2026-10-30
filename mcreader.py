"""Read "presentation" spreadsheets (merged headers, dot leaders, footnotes, notes at the
bottom) into one long, tidy table where every value knows which cell it came from.

    python mcreader.py data/            # every .xlsx listed in config/tables.csv
    python mcreader.py data/sales.xlsx  # one file, settings guessed

Writes output/<today>/tidy.csv plus a qa/ folder: rows skipped and why, cells that held a
suppression code instead of a number, the notes and footnotes, and a reconciliation that
proves every number in the sheet made it into the output exactly once.
"""

import csv
import re
import sys
import zipfile
from datetime import date
from pathlib import Path
from xml.etree import ElementTree

import openpyxl
import pandas as pd
from openpyxl.utils import range_boundaries

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "pr": "http://schemas.openxmlformats.org/package/2006/relationships"}

LEADERS = re.compile(r"[\s.…]*[…][\s.…]*|\.{3,}[\s.…]*")   # "Automobile dealers ………...……" -> "Automobile dealers"
FOOTNOTE = re.compile(r"(?<=[A-Za-z)])(\d{1,2})$")          # "GAFO1" -> ("GAFO", "1")
YEAR = re.compile(r"^(\d{4})([a-z]*)$")                     # "2022r" -> 2022, revised
LEGEND = re.compile(r"^\s*([A-Z]{1,3})\s+-\s+(.+)")          # "S - Estimate does not meet..."
NOTE_START = re.compile(r"^\s*(footnotes?:|notes?:|source:|\d+[A-Za-z]|[a-z][A-Z])")


class StopRun(Exception):
    """A check failed. Nothing has been written."""


def merged_ranges(path, sheet_name):
    """Merged ranges straight from the sheet's XML.

    openpyxl only reports merged cells when it loads the whole workbook into memory. In
    read-only (streaming) mode it doesn't see them at all. The ranges are sitting in the
    sheet XML as <mergeCell ref="D4:E4"/>, so read them from there and keep streaming the
    values. On these Census files it doesn't matter; on a 150k-row export it's the
    difference between seconds and minutes.
    """
    with zipfile.ZipFile(path) as z:
        wb = ElementTree.fromstring(z.read("xl/workbook.xml"))
        rid = next(s.get(f"{{{NS['r']}}}id") for s in wb.find("m:sheets", NS) if s.get("name") == sheet_name)
        rels = ElementTree.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == rid).lstrip("/")
        sheet = ElementTree.fromstring(z.read(target if target.startswith("xl/") else "xl/" + target))
    cells = sheet.find("m:mergeCells", NS)
    return [range_boundaries(m.get("ref")) for m in cells] if cells is not None else []


def load_grid(path, sheet_name=None):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet_name] if sheet_name else wb.worksheets[0]
    grid = [list(row) for row in ws.iter_rows(values_only=True)]
    name = ws.title
    wb.close()
    width = max((len(r) for r in grid), default=0)
    return [r + [None] * (width - len(r)) for r in grid], name


def text(v):
    return "" if v is None else str(v).strip()


def is_number(v):
    if isinstance(v, (int, float)):
        return True
    try:
        float(text(v).replace(",", ""))
        return text(v) != ""
    except ValueError:
        return False


def clean_label(v):
    s = LEADERS.sub(" ", text(v)).strip()
    m = FOOTNOTE.search(s)
    return (s[:m.start()].strip(), m.group(1)) if m else (s, "")


def find_header(grid, id_cols):
    """The header starts at the first row with 3+ filled cells and runs until a row has
    numbers in the value columns. Title rows above it are one long merged cell each."""
    start = next(i for i, r in enumerate(grid) if sum(text(v) != "" for v in r) >= 3)
    end = start + 1
    while end < len(grid) and not any(is_number(v) for v in grid[end][id_cols:]):
        end += 1
    return start, end


def header_names(grid, start, end, merges, id_cols):
    """One name per column, built down through the header rows. A merged header cell
    (Annual payroll over Amount | Percent of total) is copied across its whole span first,
    which forward-filling can't do safely: a blank header cell is sometimes a merge and
    sometimes really blank."""
    rows = [list(grid[i]) for i in range(start, end)]
    for c1, r1, c2, r2 in merges:
        if start < r1 <= end:
            v = grid[r1 - 1][c1 - 1]
            for r in range(r1, r2 + 1):
                for c in range(c1, c2 + 1):
                    if start < r <= end:
                        rows[r - 1 - start][c - 1] = v
    names, footnotes = [], {}
    for c in range(len(grid[start])):
        parts = []
        for row in rows:
            part, note = clean_label(row[c])
            if note:
                footnotes[part] = note
            if part and (not parts or parts[-1] != part):
                parts.append(part)
        names.append(" | ".join(parts))
    return names, footnotes


def read_sheet(path, sheet_name=None, id_cols=2, file_label=None):
    path = Path(path)
    grid, sheet = load_grid(path, sheet_name)
    merges = merged_ranges(path, sheet)
    start, end = find_header(grid, id_cols)
    names, head_notes = header_names(grid, start, end, merges, id_cols)
    title = " ".join(text(r[0]) for r in grid[:start] if text(r[0]))

    # a row merged across most of the sheet is a note, not data
    wide = {r1 for c1, r1, c2, r2 in merges if c2 - c1 >= 5 and r1 > end}
    records, skipped, notes, codes_seen = [], [], [], []
    in_notes = False
    for i in range(end, len(grid)):
        row, excel_row = grid[i], i + 1
        first = text(row[0])
        if in_notes or excel_row in wide or (NOTE_START.match(first) and not any(is_number(v) for v in row[id_cols:])):
            in_notes = True
            if first:
                notes.append({"row": excel_row, "text": first})
            continue
        if not any(text(v) for v in row):
            skipped.append({"row": excel_row, "reason": "blank"})
            continue
        code = text(row[0])
        label, label_note = clean_label(row[1]) if id_cols > 1 else ("", "")
        for c in range(id_cols, len(names)):
            if not names[c]:
                continue  # stray cells past the last header
            v = row[c]
            measure, measure_note = names[c], ""
            m = YEAR.match(measure)
            if m:
                measure, measure_note = m.group(1), m.group(2)  # "r" = revised, per the footnotes
            if text(v) == "":
                continue
            rec = {"file": file_label or path.name, "sheet": sheet, "naics": code, "kind_of_business": label,
                   # the source indents labels with spaces, but not consistently (3 here, 4 there),
                   # so the level comes from the NAICS code: 3 digits is a subsector, 4 a group, and so on
                   "level": max(len(code) - 2, 0) if code.isdigit() else 0,
                   "label_footnote": label_note, "measure": measure, "measure_flag": measure_note,
                   "value": float(text(v).replace(",", "")) if is_number(v) else None,
                   "code": "" if is_number(v) else text(v), "source_row": excel_row, "source_col": c + 1}
            records.append(rec)
            if rec["code"]:
                codes_seen.append(rec)

    legend = {m.group(1): m.group(2).strip() for n in notes for m in [LEGEND.match(n["text"])] if m}
    numbers_in_sheet = sum(is_number(grid[r][c]) for r in range(end, len(grid)) for c in range(id_cols, len(names))
                           if names[c] and r + 1 not in {n["row"] for n in notes})
    numbers_out = sum(r["value"] is not None for r in records)
    unknown = sorted({r["code"] for r in codes_seen} - set(legend))
    return {"records": records, "skipped": skipped, "notes": notes, "legend": legend, "codes": codes_seen,
            "title": title, "header_rows": (start + 1, end), "header_footnotes": head_notes,
            "check": {"file": file_label or path.name, "numbers_in_sheet": numbers_in_sheet,
                      "numbers_out": numbers_out, "unknown_codes": unknown}}


def run(paths, out_root="output", settings="config/tables.csv"):
    known = {}
    if Path(settings).exists():
        with open(settings, newline="") as f:
            known = {r["file"]: r for r in csv.DictReader(f)}
    results = []
    for p in paths:
        s = known.get(Path(p).name, {})
        results.append(read_sheet(p, s.get("sheet") or None, int(s.get("id_cols") or 2), s.get("label") or None))

    # hard stop before anything is written: every number in, every number out, no mystery codes
    checks = [r["check"] for r in results]
    bad = [c for c in checks if c["numbers_in_sheet"] != c["numbers_out"] or c["unknown_codes"]]
    if bad:
        raise StopRun(f"reconciliation failed: {bad}")

    out = Path(out_root) / date.today().isoformat()
    (out / "qa").mkdir(parents=True, exist_ok=True)
    tidy = pd.DataFrame([rec for r in results for rec in r["records"]])
    tidy.to_csv(out / "tidy.csv", index=False)
    pd.DataFrame(checks).to_csv(out / "qa" / "reconciliation.csv", index=False)
    pd.DataFrame([{"file": r["check"]["file"], **s} for r in results for s in r["skipped"]],
                 columns=["file", "row", "reason"]).to_csv(out / "qa" / "skipped_rows.csv", index=False)
    pd.DataFrame([{**c, "meaning": r["legend"].get(c["code"], "")} for r in results for c in r["codes"]],
                 columns=list(tidy.columns) + ["meaning"]).to_csv(out / "qa" / "suppressed_cells.csv", index=False)
    with open(out / "qa" / "notes.txt", "w") as f:
        for r in results:
            f.write(f"# {r['check']['file']}: {r['title']}\n# header rows {r['header_rows'][0]}-{r['header_rows'][1]}\n")
            f.writelines(f"row {n['row']}: {n['text']}\n" for n in r["notes"])
            f.write("\n")
    return out, tidy, checks


if __name__ == "__main__":
    args = [Path(a) for a in sys.argv[1:]] or [Path("data")]
    files = sorted(f for a in args for f in (a.glob("*.xlsx") if a.is_dir() else [a]))
    where, tidy, checks = run(files)
    for c in checks:
        print(f"{c['file']}: {c['numbers_out']} numbers out of {c['numbers_in_sheet']}")
    print(f"{len(tidy)} rows -> {where / 'tidy.csv'}")

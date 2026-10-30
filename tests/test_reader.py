from pathlib import Path

import openpyxl
import pytest

import mcreader


def make_book(path, codes=("S",), legend=True):
    """A small sheet built like the Census ones: a title, a two-row header with merges,
    dot leaders, a footnote glued to a word, a suppression code, notes at the bottom."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Expenses"
    ws["A1"] = "Estimated Operating Expenses of Shops: 2025"
    ws.merge_cells("A1:F1")
    ws.append([])
    ws.append(["Code", "Kind of business", "Payroll", None, "Rent2", None])
    ws.append([None, None, "Amount", "Percent of total", "Amount", "Percent of total"])
    ws.merge_cells("C3:D3")
    ws.merge_cells("E3:F3")
    ws.merge_cells("A3:A4")
    ws.append([None, "All shops, total ………………...", 500, 40.0, 120, 9.6])
    ws.append(["441", "  Bike shops1 …………", 200, 38.5, codes[0], codes[0]])
    ws.append([])
    ws.append(["4411", "    Used bike shops ...........", "1,050", 41.2, 30, 1.2])
    if legend:
        ws.append(["S - Estimate does not meet publication standards."])
        ws.merge_cells(f"A{ws.max_row}:F{ws.max_row}")
    ws.append(["Footnotes:"])
    ws.append(["1Includes repair-only shops."])
    ws.append(["2Rent excludes owned buildings."])
    wb.save(path)
    return path


def test_merged_headers_become_one_name_per_column(tmp_path):
    r = mcreader.read_sheet(make_book(tmp_path / "shops.xlsx"))
    assert r["header_rows"] == (3, 4)
    assert sorted({x["measure"] for x in r["records"]}) == [
        "Payroll | Amount", "Payroll | Percent of total", "Rent | Amount", "Rent | Percent of total"]
    assert r["header_footnotes"] == {"Rent": "2"}


def test_labels_lose_leaders_keep_footnotes_and_level_comes_from_the_code(tmp_path):
    r = mcreader.read_sheet(make_book(tmp_path / "shops.xlsx"))
    by_label = {x["kind_of_business"]: x for x in r["records"]}
    assert set(by_label) == {"All shops, total", "Bike shops", "Used bike shops"}
    assert by_label["Bike shops"]["label_footnote"] == "1"
    assert [by_label[k]["level"] for k in ("All shops, total", "Bike shops", "Used bike shops")] == [0, 1, 2]


def test_every_value_keeps_its_cell_and_codes_are_explained(tmp_path):
    r = mcreader.read_sheet(make_book(tmp_path / "shops.xlsx"))
    used = next(x for x in r["records"] if x["kind_of_business"] == "Used bike shops" and x["measure"] == "Payroll | Amount")
    assert (used["value"], used["source_row"], used["source_col"]) == (1050.0, 8, 3)
    assert [(c["code"], c["source_row"]) for c in r["codes"]] == [("S", 6), ("S", 6)]
    assert r["legend"] == {"S": "Estimate does not meet publication standards."}
    assert r["skipped"] == [{"row": 7, "reason": "blank"}]
    assert r["check"]["numbers_in_sheet"] == r["check"]["numbers_out"] == 10


def test_a_code_nobody_explained_stops_the_run_before_anything_is_written(tmp_path):
    path = make_book(tmp_path / "shops.xlsx", codes=("Q",))
    with pytest.raises(mcreader.StopRun):
        mcreader.run([path], out_root=tmp_path / "out", settings=tmp_path / "none.csv")
    assert not (tmp_path / "out").exists()


def test_a_full_run_writes_the_tidy_table_and_qa(tmp_path):
    path = make_book(tmp_path / "shops.xlsx")
    out, tidy, checks = mcreader.run([path], out_root=tmp_path / "out", settings=tmp_path / "none.csv")
    assert len(tidy) == 12 and checks[0]["numbers_out"] == 10
    assert {p.name for p in (out / "qa").iterdir()} == {"reconciliation.csv", "skipped_rows.csv",
                                                        "suppressed_cells.csv", "notes.txt"}


@pytest.mark.skipif(not Path("data/bes.xlsx").exists(), reason="Census files not downloaded")
def test_the_real_census_tables_reconcile():
    for f in ("sales.xlsx", "invent.xlsx", "bes.xlsx"):
        c = mcreader.read_sheet(Path("data") / f)["check"]
        assert c["numbers_in_sheet"] == c["numbers_out"] > 900 and not c["unknown_codes"]

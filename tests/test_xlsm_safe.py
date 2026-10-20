import os
import zipfile
from pathlib import Path

import pytest

from xlsm_safe import Workbook, XlsmError, apply, extract, verify

DATA = Path(__file__).resolve().parent.parent / "data"
PAIRS = {
    "rack": ("INSP_2205_PALLET-RACK_A_J_VER1700.xlsm", "例)94715タカラスタンダード／新潟_2205_PALLET-RACK_A_J_VER1700.xlsm"),
    "crane": ("INSP_240G_PC-CRANE_A_J_VER2300.xlsm", "例)94715_タカラスタンダード様_002_INSP_240G_PC-CRANE_A_J_VER2300.xlsm"),
    "elec": ("INSP_X102_ELECTRICAL-WORK_A_J_VER1200.xlsm", "例)94715タカラスタンダード／新潟_INSP_X102_ELECTRICAL-WORK_A_J_VER1200.xlsm"),
}
RACK = str(DATA / PAIRS["rack"][0])

pytestmark = pytest.mark.skipif(not os.path.exists(RACK), reason="sample data not available")


@pytest.mark.parametrize("name", PAIRS)
def test_roundtrip_matches_filled_example(name, tmp_path):
    template, filled = (str(DATA / f) for f in PAIRS[name])
    spec, _ = extract(template, filled)
    out = str(tmp_path / "out.xlsm")
    apply(template, spec, out)
    result = verify(template, out, expect=filled)
    assert result["problems"] == []
    assert result["expect_mismatches"] == []


def test_untouched_parts_are_byte_identical(tmp_path):
    out = str(tmp_path / "out.xlsm")
    apply(RACK, {"sheets": {"表紙": {"cells": {"Z11": 94715}}}}, out)
    with zipfile.ZipFile(RACK) as a, zipfile.ZipFile(out) as b:
        assert a.namelist() == b.namelist()
        changed = {n for n in a.namelist() if a.read(n) != b.read(n)}
    assert changed == {"xl/worksheets/sheet4.xml", "xl/workbook.xml"}


def test_japanese_text_and_newlines_survive(tmp_path):
    out = str(tmp_path / "out.xlsm")
    text = "タカラスタンダード新潟工場\n①　全角スペース・記号（※）"
    apply(RACK, {"sheets": {"表紙": {"cells": {"F9": text}}}}, out)
    assert Workbook(out).sheet("表紙").get("F9")[0] == text


def test_formula_cells_are_protected(tmp_path):
    wb = Workbook(RACK)
    sheet = wb.sheet("M01-1 ラック")
    ref = next(r for r, _, _, isf in sheet.iter_cells() if isf)
    with pytest.raises(XlsmError, match="formula"):
        apply(RACK, {"sheets": {"M01-1 ラック": {"cells": {ref: 1}}}}, str(tmp_path / "x.xlsm"))
    assert not (tmp_path / "x.xlsm").exists()


def test_errors_abort_without_writing(tmp_path):
    out = tmp_path / "x.xlsm"
    with pytest.raises(XlsmError, match="not found"):
        apply(RACK, {"sheets": {"M01-2 (9)": {"cells": {"A1": 1}}}}, str(out))
    assert not out.exists()


def test_similar_sheet_names_are_never_confused(tmp_path):
    out = str(tmp_path / "out.xlsm")
    apply(RACK, {"sheets": {"M01-2 (3)": {"state": "hidden"}}}, out)
    states = {s.name: s.state for s in Workbook(out).sheets}
    assert states["M01-2 (3)"] == "hidden"
    assert states["M01-2"] == states["M01-2 (2)"] == states["M01-2 (4)"] == "visible"


def test_fullwidth_name_fallback_is_reported():
    wb = Workbook(RACK)
    assert wb.sheet("M01-2（2）").name == "M01-2 (2)"
    assert wb.warnings
    with pytest.raises(XlsmError):
        wb.sheet("M01-2 (5)")


def test_merged_cell_requires_anchor(tmp_path):
    wb = Workbook(RACK)
    sheet = wb.sheet("表紙")
    mc = sheet.root.find("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}mergeCells")[0].get("ref")
    first, last = mc.split(":")
    with pytest.raises(XlsmError, match="merged"):
        sheet.set(last, "x")


def test_radio_selection_is_exclusive(tmp_path):
    out = str(tmp_path / "out.xlsm")
    apply(RACK, {"sheets": {"M01-1 ラック": {"controls": {"option01": True}}}}, out)
    apply(out, {"sheets": {"M01-1 ラック": {"controls": {"option02": True}}}}, str(tmp_path / "out2.xlsm"))
    sheet = Workbook(str(tmp_path / "out2.xlsm")).sheet("M01-1 ラック")
    on = {c.name for c in sheet.controls() if c.kind == "Radio" and c.group == "Group Box 23" and c.is_on}
    assert on == {"option02"}
    assert sheet.get("AL4")[0] == 2


def test_duplicate_control_names_need_disambiguation():
    sheet = Workbook(RACK).sheet("M04-2")
    with pytest.raises(XlsmError, match="Several controls"):
        sheet.find_control("Check Box 1")
    assert sheet.find_control("link:BE15").link == "$BE$15"


def test_delete_sheet_reindexes_names_and_removes_parts(tmp_path):
    template = str(DATA / PAIRS["crane"][0])
    out = str(tmp_path / "out.xlsm")
    apply(template, {"delete_sheets": ["A08(走行ビームラック)"]}, out)
    wb = Workbook(out)
    assert "A08(走行ビームラック)" not in wb.sheet_names()
    xml = wb.pkg.read("xl/workbook.xml").decode()
    assert "A08(走行ビームラック)" not in xml
    assert verify(template, out)["ok"]


# -- Python API -------------------------------------------------------------

from xlsm_safe import Template  # noqa: E402

ELEC = str(DATA / PAIRS["elec"][0])


def test_api_edit_and_save(tmp_path):
    t = Template(ELEC)
    cover = t["表紙"]
    cover["Z11"] = 94715
    cover.update({"F9": "タカラスタンダード新潟工場向け自動倉庫", "AB41": "落合"})
    out = tmp_path / "out.xlsm"
    report = t.save(out)
    assert report.cells == 3 and report.verify["ok"]
    saved = Template(out)["表紙"]
    assert saved["Z11"] == 94715 and saved["AB41"] == "落合"


def test_api_matches_spec_roundtrip(tmp_path):
    template, filled = (str(DATA / f) for f in PAIRS["crane"])
    spec, _ = extract(template, filled)
    t = Template(template)
    t.apply_spec(spec)
    out = str(tmp_path / "out.xlsm")
    t.save(out)
    assert verify(template, out, expect=filled)["expect_mismatches"] == []


def test_api_controls(tmp_path):
    t = Template(RACK)
    sheet = t["M01-1 ラック"]
    keys = {c.key for c in sheet.controls()}
    assert "option01" in keys
    sheet.select("option01")
    t["M04-2"].check("link:BE15")
    out = tmp_path / "out.xlsm"
    t.save(out)
    saved = Template(out)
    assert saved["M01-1 ラック"]["AL4"] == 1
    assert saved["M04-2"]["BE15"] is True
    with pytest.raises(XlsmError, match="not a radio"):
        saved["M04-2"].select("link:BE15")


def test_api_formula_and_deleted_sheet_guards(tmp_path):
    t = Template(str(DATA / PAIRS["crane"][0]))
    m01 = t["M01"]
    ref = next(r for r, _, _, isf in m01._sheet.iter_cells() if isf)
    with pytest.raises(XlsmError, match="formula"):
        m01[ref] = 1
    t.delete("A08(走行ビームラック)")
    assert "A08(走行ビームラック)" not in t.sheet_names
    with pytest.raises(XlsmError, match="scheduled for deletion"):
        t["A08(走行ビームラック)"]
    report = t.save(tmp_path / "out.xlsm")
    assert report.deleted == 1


def test_api_never_overwrites_template_and_saves_once(tmp_path):
    t = Template(ELEC)
    with pytest.raises(XlsmError, match="new file"):
        t.save(ELEC)
    t.save(tmp_path / "a.xlsm")
    with pytest.raises(XlsmError, match="already saved"):
        t.save(tmp_path / "b.xlsm")
    assert not list(tmp_path.glob(".xlsm-safe-*"))

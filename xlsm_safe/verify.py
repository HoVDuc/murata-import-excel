"""Structural checks: prove that an output file kept the template's structure."""

from __future__ import annotations

import posixpath

from lxml import etree

from .refs import replace_sheet_refs
from .spec import CONTROL_KINDS, _control_value
from .workbook import NS, Workbook, q

# Sheet children that hold data or view state and may legitimately change.
_DATA_TAGS = {q("m:sheetData"), q("m:dimension"), q("m:sheetViews")}
# Parts that must be byte-identical unless they belonged to a deleted sheet.
_FROZEN_DIRS = ("xl/vbaProject.bin", "xl/drawings/drawing", "xl/media/", "xl/activeX/",
                "xl/embeddings/", "xl/printerSettings/", "xl/styles.xml", "xl/theme/")


def _c14n(el) -> bytes:
    return etree.tostring(el, method="c14n")


def verify(template: str, output: str, expect: str | None = None) -> dict:
    t, o = Workbook(template), Workbook(output)
    problems: list[str] = []
    notes: list[str] = []
    deleted = [n for n in t.sheet_names() if n not in o.sheet_names()]

    # 1. Untouched binary/drawing parts must be identical.
    for name in t.pkg.names():
        if not name.startswith(_FROZEN_DIRS):
            continue
        if not o.pkg.exists(name):
            notes.append(f"part removed: {name}") if deleted else problems.append(f"part missing: {name}")
        elif t.pkg.read(name) != o.pkg.read(name):
            problems.append(f"part modified: {name}")

    # 2. Sheet order is preserved (apart from deleted sheets).
    if [n for n in t.sheet_names() if n not in deleted] != o.sheet_names():
        problems.append("sheet order changed")

    for os_ in o.sheets:
        ts = t.sheet(os_.name)
        # 3. Layout, drawings, controls, merges, column widths, page setup unchanged.
        tchildren = {c.tag: c for c in ts.root if c.tag not in _DATA_TAGS}
        ochildren = {c.tag: c for c in os_.root if c.tag not in _DATA_TAGS}
        for tag in set(tchildren) | set(ochildren):
            a, b = tchildren.get(tag), ochildren.get(tag)
            if a is None or b is None or _c14n(a) != _c14n(b):
                problems.append(f"{os_.name}: <{etree.QName(tag).localname}> changed")
        # 4. Row heights / hidden rows unchanged (shapes are anchored to rows).
        trows = {r.get("r"): dict(r.attrib) for r in ts.root.find("m:sheetData", NS)}
        orows = {r.get("r"): dict(r.attrib) for r in os_.root.find("m:sheetData", NS)}
        for rn, attrs in trows.items():
            oa = {k: v for k, v in orows.get(rn, {}).items() if k != "spans"}
            ta = {k: v for k, v in attrs.items() if k != "spans"}
            if oa != ta:
                problems.append(f"{os_.name}: row {rn} attributes changed {ta} -> {oa}")
        # 5. Every formula survives (modulo #REF! for deleted sheets).
        ofx = {r: fx for r, _, fx, isf in os_.iter_cells() if isf}
        for r, _, fx, isf in ts.iter_cells():
            if not isf:
                continue
            want = fx
            for d in deleted:
                want = replace_sheet_refs(want, d) if want else want
            if r not in ofx:
                problems.append(f"{os_.name}!{r}: formula removed ({fx})")
            elif ofx[r] != want:
                problems.append(f"{os_.name}!{r}: formula changed {fx!r} -> {ofx[r]!r}")
        # 6. Control names unchanged (macros address them by name).
        tn = sorted((c.shape_id, c.name, c.kind, c.link or "", c.anchor) for c in ts.controls())
        on = sorted((c.shape_id, c.name, c.kind, c.link or "", c.anchor) for c in os_.controls())
        if tn != on:
            problems.append(f"{os_.name}: form controls differ")
        # 7. Radio groups have at most one selected button.
        groups: dict[str, int] = {}
        for c in os_.controls():
            if c.kind == "Radio" and c.is_on:
                groups[c.group] = groups.get(c.group, 0) + 1
        for g, n in groups.items():
            if n > 1:
                problems.append(f"{os_.name}: radio group {g!r} has {n} selected buttons")

    result = {"ok": not problems, "problems": problems, "notes": notes, "deleted_sheets": deleted}
    if expect:
        result["expect_mismatches"] = compare_content(output, expect)
        result["ok"] = result["ok"] and not result["expect_mismatches"]
    return result


def compare_content(actual: str, expected: str) -> list[str]:
    """Cell values (non-formula), sheet states and control states must match."""
    a, e = Workbook(actual), Workbook(expected)
    out: list[str] = []
    if a.sheet_names() != e.sheet_names():
        out.append(f"sheet list differs: {a.sheet_names()} vs {e.sheet_names()}")
    for es in e.sheets:
        if es.name not in a.sheet_names():
            continue
        as_ = a.sheet(es.name)
        if as_.state != es.state:
            out.append(f"{es.name}: state {as_.state} != {es.state}")
        av = {r: v for r, v, _, isf in as_.iter_cells() if not isf}
        ev = {r: v for r, v, _, isf in es.iter_cells() if not isf}
        for r in set(av) | set(ev):
            x, y = av.get(r), ev.get(r)
            if x != y and not (x in (None, "") and y in (None, "")):
                anchor = es.merge_anchor(r)
                if anchor and anchor != r:
                    continue
                out.append(f"{es.name}!{r}: {x!r} != {y!r}")
        ac = {c.shape_id: _control_value(c) for c in as_.controls() if c.kind in CONTROL_KINDS}
        for c in es.controls():
            if c.kind in CONTROL_KINDS and ac.get(c.shape_id) != _control_value(c):
                out.append(f"{es.name}: control {c.name!r} (id {c.shape_id}) {ac.get(c.shape_id)} != {_control_value(c)}")
    return out


def inspect(path: str) -> dict:
    wb = Workbook(path)
    sheets = []
    for i, s in enumerate(wb.sheets):
        ctrls = s.control_keys()
        sheets.append({
            "index": i, "name": s.name, "state": s.state, "codeName": s.code_name,
            "part": posixpath.basename(s.part),
            "controls": [
                {"name": k, "type": c.kind, "link": c.link or c.extra.get("group_link"),
                 "checked": _control_value(c) if c.kind in CONTROL_KINDS else None,
                 **({"group": c.group, "index": c.group_index} if c.kind == "Radio" else {})}
                for k, c in ctrls.items() if c.kind in CONTROL_KINDS or c.kind == "ActiveX"
            ],
        })
    return {"file": path, "sheets": sheets}

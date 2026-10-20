"""Import spec: a JSON document describing what to write into a template.

{
  "sheets": {
    "表紙":        {"cells": {"Z11": 94715, "F9": "タカラ..."}},
    "M01-1 ラック": {"cells": {"S6": 4805}, "controls": {"option01": true}},
    "M01-2 (2)":   {"state": "hidden"},
    "A10(能力)":   {"formula_overwrites": {"X5": 12}}
  },
  "delete_sheets": ["A08(走行ビームラック)"]
}

cells:              value = string | number | bool | null (null clears the cell)
controls:           CheckBox -> true/false/"mixed", Radio -> true (selects it in its group)/false
state:              "visible" | "hidden" | "veryHidden"
formula_overwrites: like cells, but allowed to replace a formula (use sparingly)
"""

from __future__ import annotations

from .workbook import Workbook

CONTROL_KINDS = ("CheckBox", "Radio")


def _control_value(c):
    if c.checked == "Mixed":
        return "mixed"
    return c.is_on


def extract(template: str, filled: str) -> tuple[dict, list[str]]:
    """Derive the spec that turns `template` into `filled`."""
    t, f = Workbook(template), Workbook(filled)
    warnings: list[str] = []
    spec: dict = {"sheets": {}, "delete_sheets": []}
    fnames = set(f.sheet_names())
    spec["delete_sheets"] = [n for n in t.sheet_names() if n not in fnames]

    for fs in f.sheets:
        if fs.name not in t.sheet_names():
            warnings.append(f"Sheet {fs.name!r} does not exist in the template; adding sheets is not supported")
            continue
        ts = t.sheet(fs.name)
        entry: dict = {}
        if fs.state != ts.state:
            entry["state"] = fs.state

        tc = {r: (v, fx, isf) for r, v, fx, isf in ts.iter_cells()}
        fc = {r: (v, fx, isf) for r, v, fx, isf in fs.iter_cells()}
        cells, overwrites = {}, {}
        for ref in sorted(set(tc) | set(fc), key=lambda r: (len(r), r)):
            tv, tf, tisf = tc.get(ref, (None, None, False))
            fv, ff, fisf = fc.get(ref, (None, None, False))
            if fisf:
                if not tisf or tf != ff:
                    warnings.append(f"{fs.name}!{ref}: formula differs from template ({tf!r} -> {ff!r}); not imported")
                continue
            if tv == fv and not tisf:
                continue
            anchor = ts.merge_anchor(ref)
            if anchor and anchor != ref:
                if fv not in (None, ""):
                    warnings.append(f"{fs.name}!{ref}: value in hidden part of merged range ignored")
                continue
            (overwrites if tisf else cells)[ref] = fv
        if cells:
            entry["cells"] = cells
        if overwrites:
            entry["formula_overwrites"] = overwrites
            warnings.append(f"{fs.name}: {len(overwrites)} formula cell(s) were replaced by constants in the filled file")

        # Shape ids are stable across saves, names are not (Excel renames copies).
        tctl = {c.shape_id: c for c in ts.controls() if c.kind in CONTROL_KINDS}
        controls = {}
        for c in fs.controls():
            tc_ = tctl.get(c.shape_id)
            if tc_ is None or tc_.kind != c.kind:
                if c.kind in CONTROL_KINDS:
                    warnings.append(f"{fs.name}: control {c.name!r} (id {c.shape_id}) not found in template")
                continue
            if _control_value(c) != _control_value(tc_):
                controls[ts.preferred_key(tc_)] = _control_value(c)
        if controls:
            entry["controls"] = controls
        if entry:
            spec["sheets"][fs.name] = entry
    return spec, warnings


def apply(template: str, spec: dict, output: str, *, verify: bool = True) -> dict:
    """Write spec into a copy of template. Raises XlsmError before saving on any problem."""
    from .api import Template

    t = Template(template)
    t.apply_spec(spec)
    r = t.save(output, verify=verify)
    stats = {"cells": r.cells, "controls": r.controls, "states": r.states, "deleted": r.deleted}
    return {"stats": stats, "warnings": r.warnings, "changed_parts": r.changed_parts, "verify": r.verify}

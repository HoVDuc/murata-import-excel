"""Workbook model on top of Package: sheets, cells, form controls, hide/delete.

Design rules (each one prevents a symptom seen with openpyxl & co.):
  * Only the XML parts that really change are re-serialized; everything else
    (vbaProject.bin, drawings, VML, ctrlProps, ActiveX) is copied as-is,
    so shapes and option buttons keep their positions and macros keep working.
  * Cells keep their existing style index; new strings are *appended* to the
    shared string table so existing indices (and phonetic data) never move.
  * Formula cells are never overwritten unless explicitly allowed, and the
    workbook is flagged for full recalculation on open.
  * Sheets are addressed by exact name only (NFKC fallback must be unique),
    never by index, so "M01-2" / "M01-2 (2)" / "M01-2 (3)" cannot be mixed up
    and deleting a sheet cannot shift data onto the wrong one.
  * Form controls are updated consistently in ctrlProp, VML and linked cell.
"""

from __future__ import annotations

import math
import re
import unicodedata
from dataclasses import dataclass, field

from lxml import etree

from .package import Package, rels_path
from .refs import (col_to_num, norm_ref, num_to_col, parse_range,
                   refers_to_sheet, replace_sheet_refs, split_ref)

M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
X14 = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
XM = "http://schemas.microsoft.com/office/excel/2006/main"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"
RT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": M, "r": R, "x14": X14, "xdr": XDR, "xm": XM, "vt": VT}

_ILLEGAL_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")
MAX_CELL_CHARS = 32767


class XlsmError(Exception):
    pass


def q(tag: str) -> str:
    prefix, _, local = tag.partition(":")
    return f"{{{NS[prefix]}}}{local}"


def clean_text(s: str) -> str:
    """Normalize text so Excel displays it exactly as given."""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = _ILLEGAL_XML.sub("", s)
    # Drop lone surrogates that would corrupt the UTF-8 stream.
    s = s.encode("utf-8", "replace").decode("utf-8")
    if len(s) > MAX_CELL_CHARS:
        raise XlsmError(f"Text longer than Excel's {MAX_CELL_CHARS} character limit")
    return s


def si_text(si: etree._Element) -> str:
    """Plain text of a shared string item, without phonetic (rPh) runs."""
    parts = []
    for child in si:
        if child.tag == q("m:t"):
            parts.append(child.text or "")
        elif child.tag == q("m:r"):
            t = child.find("m:t", NS)
            parts.append(t.text or "" if t is not None else "")
    return "".join(parts)


def _set_text(parent: etree._Element, text: str) -> None:
    t = etree.SubElement(parent, q("m:t"))
    t.text = text
    if text != text.strip(" \t\n") or "\n" in text:
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


class SharedStrings:
    def __init__(self, pkg: Package, part: str | None):
        self.pkg, self.part = pkg, part
        self.items: list[str] = []
        self._plain: dict[str, int] = {}
        if part:
            for i, si in enumerate(pkg.xml(part).findall("m:si", NS)):
                txt = si_text(si)
                self.items.append(txt)
                # Only reuse items without rich-text runs or phonetics.
                if len(si) == 1 and si[0].tag == q("m:t"):
                    self._plain.setdefault(txt, i)

    def get(self, idx: int) -> str:
        return self.items[idx]

    def add(self, text: str) -> int:
        if text in self._plain:
            return self._plain[text]
        if not self.part:
            raise XlsmError("Workbook has no shared string table")
        root = self.pkg.xml(self.part)
        si = etree.SubElement(root, q("m:si"))
        _set_text(si, text)
        idx = len(self.items)
        self.items.append(text)
        self._plain[text] = idx
        root.set("uniqueCount", str(len(self.items)))
        root.set("count", str(int(root.get("count", "0")) + 1))
        self.pkg.touch(self.part)
        return idx


@dataclass
class Control:
    sheet: str
    name: str
    shape_id: int
    kind: str  # Radio, CheckBox, GBox, Button, Drop, List, ... or ActiveX
    prop_part: str | None
    link: str | None
    checked: str | None  # None / "Checked" / "Mixed"
    first_button: bool
    anchor: tuple[tuple[int, int, int, int], tuple[int, int, int, int]] | None
    group: str | None = None
    group_index: int | None = None
    extra: dict = field(default_factory=dict)

    @property
    def is_on(self) -> bool:
        return self.checked == "Checked"


class Sheet:
    def __init__(self, wb: "Workbook", name: str, sheet_id: str, rid: str, part: str):
        self.wb, self.name, self.sheet_id, self.rid, self.part = wb, name, sheet_id, rid, part
        self._controls: list[Control] | None = None

    @property
    def root(self) -> etree._Element:
        return self.wb.pkg.xml(self.part)

    @property
    def el(self) -> etree._Element:
        """The <sheet> element in workbook.xml."""
        for s in self.wb.root.find("m:sheets", NS):
            if s.get(q("r:id")) == self.rid:
                return s
        raise XlsmError(f"Sheet {self.name!r} vanished from workbook.xml")

    @property
    def state(self) -> str:
        return self.el.get("state", "visible")

    @property
    def code_name(self) -> str | None:
        pr = self.root.find("m:sheetPr", NS)
        return pr.get("codeName") if pr is not None else None

    # -- cells ------------------------------------------------------------
    def _merged(self) -> list[tuple[int, int, int, int]]:
        mc = self.root.find("m:mergeCells", NS)
        return [parse_range(m.get("ref")) for m in mc] if mc is not None else []

    def merge_anchor(self, ref: str) -> str | None:
        """Top-left cell of the merged range containing ref (None if not merged)."""
        c, r = split_ref(ref)
        for c1, r1, c2, r2 in self._merged():
            if c1 <= c <= c2 and r1 <= r <= r2:
                return f"{num_to_col(c1)}{r1}"
        return None

    def _find_cell(self, ref: str, create: bool) -> etree._Element | None:
        col, rown = split_ref(ref)
        ref = f"{num_to_col(col)}{rown}"
        data = self.root.find("m:sheetData", NS)
        row = None
        for i, rw in enumerate(data):
            rn = int(rw.get("r"))
            if rn == rown:
                row = rw
                break
            if rn > rown:
                if not create:
                    return None
                row = etree.Element(q("m:row"), r=str(rown))
                data.insert(i, row)
                break
        if row is None:
            if not create:
                return None
            row = etree.SubElement(data, q("m:row"), r=str(rown))
        for i, c in enumerate(row):
            cc = col_to_num(re.match(r"[A-Z]+", c.get("r")).group())
            if cc == col:
                return c
            if cc > col:
                if not create:
                    return None
                cell = etree.Element(q("m:c"), r=ref)
                row.insert(i, cell)
                break
        else:
            if not create:
                return None
            cell = etree.SubElement(row, q("m:c"), r=ref)
        self._init_new_cell(cell, row, col)
        return cell

    def _init_new_cell(self, cell, row, col) -> None:
        style = row.get("s") if row.get("customFormat") == "1" else None
        if style is None:
            cols = self.root.find("m:cols", NS)
            for c in cols if cols is not None else []:
                if int(c.get("min")) <= col <= int(c.get("max")) and c.get("style"):
                    style = c.get("style")
        if style and style != "0":
            cell.set("s", style)
        spans = row.get("spans")
        if spans and ":" in spans:
            lo, hi = map(int, spans.split(":"))
            row.set("spans", f"{min(lo, col)}:{max(hi, col)}")

    def get(self, ref: str):
        """(value, formula) of a cell; value typed as str/float/int/bool/None."""
        c = self._find_cell(ref, create=False)
        if c is None:
            return None, None
        return self.wb.cell_value(c), (c.find("m:f", NS).text if c.find("m:f", NS) is not None else None)

    def iter_cells(self):
        for row in self.root.find("m:sheetData", NS):
            for c in row:
                f = c.find("m:f", NS)
                yield c.get("r"), self.wb.cell_value(c), (f.text if f is not None else None), f is not None

    def set(self, ref: str, value, *, allow_formula: bool = False) -> None:
        ref = norm_ref(ref)
        anchor = self.merge_anchor(ref)
        if anchor and anchor != ref:
            raise XlsmError(f"{self.name}!{ref} is inside a merged range; write to {anchor} instead")
        cell = self._find_cell(ref, create=True)
        f = cell.find("m:f", NS)
        if f is not None:
            if not allow_formula:
                raise XlsmError(f"{self.name}!{ref} contains a formula ({f.text}); refusing to overwrite it")
            if f.get("t") in ("shared", "array") and f.get("ref"):
                raise XlsmError(f"{self.name}!{ref} is the master of a shared/array formula used by {f.get('ref')}")
            self.wb.formulas_removed = True
        for child in list(cell):
            if child.tag in (q("m:f"), q("m:v"), q("m:is")):
                cell.remove(child)
        for attr in ("t", "cm", "vm"):
            cell.attrib.pop(attr, None)

        if value is None or value == "":
            pass
        elif isinstance(value, bool):
            cell.set("t", "b")
            etree.SubElement(cell, q("m:v")).text = "1" if value else "0"
        elif isinstance(value, (int, float)):
            if isinstance(value, float) and not math.isfinite(value):
                raise XlsmError(f"{self.name}!{ref}: cannot store {value}")
            etree.SubElement(cell, q("m:v")).text = repr(value) if isinstance(value, float) else str(value)
        elif isinstance(value, str):
            idx = self.wb.sst.add(clean_text(value))
            cell.set("t", "s")
            etree.SubElement(cell, q("m:v")).text = str(idx)
        else:
            raise XlsmError(f"{self.name}!{ref}: unsupported value type {type(value).__name__}")
        self._grow_dimension(ref)
        self.wb.pkg.touch(self.part)
        self.wb.values_changed = True

    def _grow_dimension(self, ref: str) -> None:
        dim = self.root.find("m:dimension", NS)
        if dim is None:
            return
        c, r = split_ref(ref)
        c1, r1, c2, r2 = parse_range(dim.get("ref"))
        if not (c1 <= c <= c2 and r1 <= r <= r2):
            c1, r1, c2, r2 = min(c1, c), min(r1, r), max(c2, c), max(r2, r)
            dim.set("ref", f"{num_to_col(c1)}{r1}:{num_to_col(c2)}{r2}")

    # -- form controls ----------------------------------------------------
    def vml_part(self) -> str | None:
        ld = self.root.find("m:legacyDrawing", NS)
        if ld is None:
            return None
        rel = self.wb.pkg.rel_by_id(self.part, ld.get(q("r:id")))
        return rel.part if rel else None

    def controls(self) -> list[Control]:
        if self._controls is not None:
            return self._controls
        pkg = self.wb.pkg
        out: dict[int, Control] = {}
        for ctl in self.root.iter(q("m:control")):
            sid = int(ctl.get("shapeId"))
            rel = pkg.rel_by_id(self.part, ctl.get(q("r:id")))
            if rel is None or sid in out and out[sid].anchor:
                continue
            anchor = None
            a = ctl.find("m:controlPr/m:anchor", NS)
            if a is not None:
                def pt(e):
                    return tuple(int(e.find(f"xdr:{k}", NS).text) for k in ("col", "colOff", "row", "rowOff"))
                anchor = (pt(a.find("m:from", NS)), pt(a.find("m:to", NS)))
            if rel.type.endswith("/ctrlProp"):
                pr = pkg.xml(rel.part)
                out[sid] = Control(self.name, ctl.get("name"), sid, pr.get("objectType"), rel.part,
                                   pr.get("fmlaLink"), pr.get("checked"), pr.get("firstButton") == "1", anchor)
            else:
                out[sid] = Control(self.name, ctl.get("name"), sid, "ActiveX", rel.part, None, None, False, anchor)
        self._controls = list(out.values())
        self._assign_radio_groups()
        return self._controls

    def _vml_zorder(self) -> dict[int, int]:
        vp = self.vml_part()
        if not vp:
            return {}
        data = self.wb.pkg.read(vp)
        order = {}
        for i, m in enumerate(re.finditer(rb'<v:shape\b[^>]*?"_x0000_s(\d+)"', data)):
            order.setdefault(int(m.group(1)), i)
        return order

    def _assign_radio_groups(self) -> None:
        """Excel groups option buttons by the group box that contains them;
        buttons outside every group box form one sheet-level group."""
        def inside(inner, box):
            (ifrom, ito), (bfrom, bto) = inner, box
            cols = bfrom[0:2] <= ifrom[0:2] and ito[0:2] <= bto[0:2]
            rows = bfrom[2:4] <= ifrom[2:4] and ito[2:4] <= bto[2:4]
            return cols and rows

        boxes = [c for c in self._controls if c.kind == "GBox" and c.anchor]
        zorder = self._vml_zorder()
        groups: dict[str, list[Control]] = {}
        for c in self._controls:
            if c.kind != "Radio":
                continue
            gid = "<sheet>"
            if c.anchor:
                best = None
                for b in boxes:
                    if inside(c.anchor, b.anchor):
                        (fc, _, fr, _), (tc, _, tr, _) = b.anchor
                        area = (tc - fc + 1) * (tr - fr + 1)
                        if best is None or area < best[0]:
                            best = (area, b.name)
                if best:
                    gid = best[1]
            c.group = gid
            groups.setdefault(gid, []).append(c)
        for members in groups.values():
            members.sort(key=lambda c: zorder.get(c.shape_id, c.shape_id))
            link = next((m.link for m in members if m.link), None)
            for i, m in enumerate(members, 1):
                m.group_index = i
                m.extra["group_link"] = link

    def control_keys(self) -> dict[str, Control]:
        """Unique keys for controls: the name, or 'name#n' (z-order) when Excel
        holds several controls with the same name on one sheet."""
        zorder = self._vml_zorder()
        by_name: dict[str, list[Control]] = {}
        for c in self.controls():
            by_name.setdefault(c.name, []).append(c)
        keys = {}
        for name, items in by_name.items():
            if len(items) == 1:
                keys[name] = items[0]
            else:
                items.sort(key=lambda c: zorder.get(c.shape_id, c.shape_id))
                for i, c in enumerate(items, 1):
                    keys[f"{name}#{i}"] = c
        return keys

    def preferred_key(self, ctl: Control) -> str:
        """Most readable key that identifies ctl uniquely: name, link:<cell> or id:<shapeId>."""
        if sum(c.name == ctl.name for c in self.controls()) == 1:
            return ctl.name
        if ctl.link and ctl.kind == "CheckBox" and sum(c.link == ctl.link for c in self.controls()) == 1:
            return "link:" + ctl.link.replace("$", "")
        return f"id:{ctl.shape_id}"

    def find_control(self, key: str) -> Control:
        """key = name | name#n | link:<cell> | id:<shapeId>"""
        if key.startswith("id:"):
            hits = [c for c in self.controls() if str(c.shape_id) == key[3:]]
            if hits:
                return hits[0]
            raise XlsmError(f"No control with shape id {key[3:]} on sheet {self.name!r}")
        if key.startswith("link:"):
            want = key[5:].replace("$", "").upper()
            hits = [c for c in self.controls() if c.link and c.link.replace("$", "").upper() == want]
            if len(hits) == 1:
                return hits[0]
            raise XlsmError(f"{len(hits)} controls linked to {want} on sheet {self.name!r}; need exactly one")
        keys = self.control_keys()
        if key in keys:
            return keys[key]
        dups = [k for k in keys if k.startswith(key + "#")]
        if dups:
            raise XlsmError(f"Several controls named {key!r} on sheet {self.name!r}; use one of {dups}")
        raise XlsmError(f"No control named {key!r} on sheet {self.name!r}")

    def set_control(self, name: str, value) -> None:
        ctl = self.find_control(name)
        if ctl.kind == "Radio":
            group = [c for c in self.controls() if c.kind == "Radio" and c.group == ctl.group]
            for c in group:
                self._write_check(c, "Checked" if (value and c is ctl) else None)
            link = ctl.extra.get("group_link")
            if link:
                cur = 0
                for c in group:
                    if c.is_on:
                        cur = c.group_index
                self._write_link(link, cur)
        elif ctl.kind == "CheckBox":
            state = "Mixed" if value == "mixed" else ("Checked" if value else None)
            self._write_check(ctl, state)
            if ctl.link:
                self._write_link(ctl.link, None if state == "Mixed" else bool(value))
        else:
            raise XlsmError(f"Control {name!r} on {self.name!r} is {ctl.kind}; only CheckBox/Radio can be set")

    def _write_link(self, link: str, value) -> None:
        sheet, ref = self, link
        if "!" in link:
            sname, ref = link.rsplit("!", 1)
            sheet = self.wb.sheet(sname.strip("'").replace("''", "'"))
        ref = ref.replace("$", "")
        if value is None:
            cell = sheet._find_cell(ref, create=True)
            for child in list(cell):
                if child.tag in (q("m:v"), q("m:is")):
                    cell.remove(child)
            cell.set("t", "e")
            etree.SubElement(cell, q("m:v")).text = "#N/A"
            self.wb.pkg.touch(sheet.part)
        else:
            sheet.set(ref, value, allow_formula=False)

    def _write_check(self, ctl: Control, state: str | None) -> None:
        pkg = self.wb.pkg
        pr = pkg.xml(ctl.prop_part)
        if state:
            pr.set("checked", state)
        else:
            pr.attrib.pop("checked", None)
        pkg.touch(ctl.prop_part)
        ctl.checked = state

        vp = self.vml_part()
        if not vp:
            return
        data = pkg.read(vp)
        m = re.search(rb'<v:shape\b[^>]*?"_x0000_s%d".*?</v:shape>' % ctl.shape_id, data, re.S)
        if not m:
            return
        shape = m.group(0)
        new = re.sub(rb"\s*<x:Checked>[^<]*</x:Checked>", b"", shape)
        if state:
            val = b"2" if state == "Mixed" else b"1"
            new = new.replace(b"</x:ClientData>", b"<x:Checked>" + val + b"</x:Checked></x:ClientData>", 1)
        if new != shape:
            pkg.write(vp, data[:m.start()] + new + data[m.end():])


class Workbook:
    def __init__(self, path: str):
        self.pkg = Package(path)
        self.part = next(r.part for r in self.pkg.rels("") if r.type.endswith("/officeDocument"))
        sst = next((r.part for r in self.pkg.rels(self.part) if r.type.endswith("/sharedStrings")), None)
        self.sst = SharedStrings(self.pkg, sst)
        self.values_changed = False
        self.formulas_removed = False
        self.warnings: list[str] = []
        self._load_sheets()

    @property
    def root(self) -> etree._Element:
        return self.pkg.xml(self.part)

    def _load_sheets(self) -> None:
        rels = {r.id: r for r in self.pkg.rels(self.part)}
        self.sheets: list[Sheet] = []
        for s in self.root.find("m:sheets", NS):
            rel = rels[s.get(q("r:id"))]
            self.sheets.append(Sheet(self, s.get("name"), s.get("sheetId"), rel.id, rel.part))

    def sheet_names(self) -> list[str]:
        return [s.name for s in self.sheets]

    def sheet(self, name: str) -> Sheet:
        for s in self.sheets:
            if s.name == name:
                return s
        # Fallback for typing differences only: full/half width and spaces.
        # Must be unique, so "M01-2" can never resolve to "M01-2 (2)".
        def loose(n):
            return re.sub(r"\s+", "", unicodedata.normalize("NFKC", n))
        hits = [s for s in self.sheets if loose(s.name) == loose(name)]
        if len(hits) == 1:
            self.warnings.append(f"Sheet {name!r} matched {hits[0].name!r} after full-width/space normalization")
            return hits[0]
        if hits:
            raise XlsmError(f"Sheet {name!r} is ambiguous: {[s.name for s in hits]}")
        raise XlsmError(f"Sheet {name!r} not found. Existing sheets: {self.sheet_names()}")

    def cell_value(self, c: etree._Element):
        t = c.get("t")
        v = c.find("m:v", NS)
        if t == "inlineStr":
            is_ = c.find("m:is", NS)
            return si_text(is_) if is_ is not None else None
        if v is None or v.text is None:
            return None
        if t == "s":
            return self.sst.get(int(v.text))
        if t == "b":
            return v.text == "1"
        if t in ("str", "e"):
            return v.text
        try:
            n = float(v.text)
        except ValueError:
            return v.text
        return int(n) if n.is_integer() and re.fullmatch(r"-?\d+", v.text) else n

    # -- sheet visibility -------------------------------------------------
    def set_state(self, name: str, state: str) -> None:
        if state not in ("visible", "hidden", "veryHidden"):
            raise XlsmError(f"Invalid sheet state {state!r}")
        sh = self.sheet(name)
        if state == sh.state:
            return
        if state != "visible":
            visible = [s for s in self.sheets if s.state == "visible" and s is not sh]
            if not visible:
                raise XlsmError("Cannot hide the last visible sheet")
        if state == "visible":
            sh.el.attrib.pop("state", None)
        else:
            sh.el.set("state", state)
        self.pkg.touch(self.part)
        self._fix_active_tab()

    def _fix_active_tab(self) -> None:
        """Make sure the active / first tab is a visible sheet."""
        bv = self.root.find("m:bookViews/m:workbookView", NS)
        if bv is None:
            return
        n = len(self.sheets)
        visible = [i for i, s in enumerate(self.sheets) if s.state == "visible"]
        active = min(int(bv.get("activeTab", "0")), n - 1)
        if active not in visible:
            active = visible[0]
        first = min(int(bv.get("firstSheet", "0")), n - 1)
        if first not in visible or first > active:
            first = min(v for v in visible if v <= active)
        bv.set("activeTab", str(active))
        if first:
            bv.set("firstSheet", str(first))
        else:
            bv.attrib.pop("firstSheet", None)
        for i, s in enumerate(self.sheets):
            sv = s.root.find("m:sheetViews/m:sheetView", NS)
            if sv is None:
                continue
            want = i == active
            if (sv.get("tabSelected") == "1") != want:
                if want:
                    sv.set("tabSelected", "1")
                else:
                    sv.attrib.pop("tabSelected", None)
                self.pkg.touch(s.part)
        self.pkg.touch(self.part)

    # -- sheet deletion ---------------------------------------------------
    def delete_sheet(self, name: str) -> None:
        sh = self.sheet(name)
        idx = self.sheets.index(sh)
        if not [s for s in self.sheets if s is not sh and s.state == "visible"]:
            raise XlsmError("Cannot delete the last visible sheet")
        before = self.pkg.reachable()

        # 1. workbook.xml <sheet> + relationship
        sheets_el = self.root.find("m:sheets", NS)
        sheets_el.remove(sh.el)
        self.pkg.remove_rel(self.part, sh.rid)
        self.sheets.remove(sh)

        # 2. defined names: drop sheet-scoped, re-index the rest, #REF! the references
        dn = self.root.find("m:definedNames", NS)
        for d in list(dn) if dn is not None else []:
            lid = d.get("localSheetId")
            if lid is not None:
                lid = int(lid)
                if lid == idx:
                    dn.remove(d)
                    continue
                if lid > idx:
                    d.set("localSheetId", str(lid - 1))
            if d.text and refers_to_sheet(d.text, sh.name):
                d.text = replace_sheet_refs(d.text, sh.name)
        if dn is not None and len(dn) == 0:
            self.root.remove(dn)

        # 3. formulas on other sheets that point to the deleted sheet
        for other in self.sheets:
            changed = False
            for tag in ("m:f", "m:formula", "m:formula1", "m:formula2", "xm:f"):
                for f in other.root.iter(q(tag)):
                    if f.text and refers_to_sheet(f.text, sh.name):
                        f.text = replace_sheet_refs(f.text, sh.name)
                        changed = True
            if changed:
                self.pkg.touch(other.part)
                self.warnings.append(f"Formulas on {other.name!r} referenced {sh.name!r}; they now return #REF!")

        # 4. drop parts that only the deleted sheet used (drawings, VML, ctrlProps, ...)
        self.pkg.touch(self.part)
        removed = before - self.pkg.reachable()
        removed |= {sh.part, rels_path(sh.part)}
        calc = next((r for r in self.pkg.rels(self.part) if r.type.endswith("/calcChain")), None)
        if calc:  # calcChain indexes sheets by position; Excel rebuilds it
            self.pkg.remove_rel(self.part, calc.id)
            removed.add(calc.part)
        for p in removed:
            if self.pkg.exists(p):
                self.pkg.remove(p)
        self.pkg.remove_content_type_overrides(removed)

        self._fix_active_tab()
        self._fix_app_props(sh.name)
        if sh.code_name:
            self.warnings.append(
                f"VBA module {sh.code_name!r} of deleted sheet {sh.name!r} stays in vbaProject.bin "
                "(binary not rewritten). Open once in Excel and save to let Excel clean it up.")

    def _fix_app_props(self, sheet_name: str) -> None:
        part = "docProps/app.xml"
        if not self.pkg.exists(part):
            return
        root = self.pkg.xml(part)
        hp = root.find(f"{{*}}HeadingPairs/{q('vt:vector')}")
        tp = root.find(f"{{*}}TitlesOfParts/{q('vt:vector')}")
        if hp is None or tp is None:
            return
        prefixes = ("'" + sheet_name.replace("'", "''") + "'!", sheet_name + "!")
        variants = list(hp)
        titles = list(tp)
        pos = 0
        for i in range(0, len(variants) - 1, 2):
            cnt_el = variants[i + 1].find(q("vt:i4"))
            cnt = int(cnt_el.text)
            block = titles[pos:pos + cnt]
            pos += cnt
            kept = 0
            for t in block:
                txt = t.text or ""
                if txt == sheet_name or txt.startswith(prefixes):
                    tp.remove(t)
                else:
                    kept += 1
            cnt_el.text = str(kept)
        tp.set("size", str(len(tp)))
        self.pkg.touch(part)

    # -- save -------------------------------------------------------------
    def save(self, path: str) -> None:
        if self.values_changed or self.formulas_removed:
            calc = self.root.find("m:calcPr", NS)
            if calc is None:
                calc = etree.Element(q("m:calcPr"))
                anchor = None
                for tag in ("m:definedNames", "m:sheets"):
                    anchor = self.root.find(tag, NS)
                    if anchor is not None:
                        break
                anchor.addnext(calc)
            calc.set("fullCalcOnLoad", "1")
            self.pkg.touch(self.part)
        if self.formulas_removed:
            calc = next((r for r in self.pkg.rels(self.part) if r.type.endswith("/calcChain")), None)
            if calc:
                self.pkg.remove_rel(self.part, calc.id)
                self.pkg.remove(calc.part)
                self.pkg.remove_content_type_overrides({calc.part})
        self.pkg.save(path)

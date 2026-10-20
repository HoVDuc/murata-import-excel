"""Python API: edit a template in code instead of through a JSON spec.

    from xlsm_safe import Template

    t = Template("data/INSP_X102_ELECTRICAL-WORK_A_J_VER1200.xlsm")
    cover = t["表紙"]
    cover["Z11"] = 94715
    cover.update({"F9": "タカラスタンダード新潟工場向け自動倉庫", "AB41": "落合"})
    t["E01 結線確認"].check("link:AN11")
    t["M01-1 ラック"].select("option01")       # radio button
    t.hide("M01-2 (2)")
    t.delete("A08(走行ビームラック)")           # applied on save
    report = t.save("out.xlsm")                # verified before it lands on disk

Nothing touches the disk before save(). save() writes to a temporary file,
runs the structural verification and only then moves it to the output path,
so a failed import never leaves a broken file behind.
"""

from __future__ import annotations

import os
import tempfile
from dataclasses import dataclass, field

from .spec import CONTROL_KINDS, _control_value
from .verify import verify as verify_files
from .workbook import Workbook, XlsmError


@dataclass
class ControlInfo:
    key: str  # what to pass to check()/select()/set_control()
    name: str
    type: str
    link: str | None
    checked: bool | str | None
    group: str | None = None


@dataclass
class ImportReport:
    output: str
    cells: int = 0
    controls: int = 0
    states: int = 0
    deleted: int = 0
    warnings: list[str] = field(default_factory=list)
    changed_parts: list[str] = field(default_factory=list)
    verify: dict | None = None


class SheetEditor:
    """One sheet of a Template. Get one with template["sheet name"]."""

    def __init__(self, template: "Template", sheet):
        self._t = template
        self._sheet = sheet

    @property
    def name(self) -> str:
        return self._sheet.name

    @property
    def state(self) -> str:
        return self._sheet.state

    # -- cells ------------------------------------------------------------
    def __getitem__(self, ref: str):
        return self._sheet.get(ref)[0]

    def __setitem__(self, ref: str, value) -> None:
        self.set(ref, value)

    def formula(self, ref: str) -> str | None:
        return self._sheet.get(ref)[1]

    def set(self, ref: str, value, *, overwrite_formula: bool = False) -> None:
        """Write str / int / float / bool / None. Formula cells are refused unless overwrite_formula."""
        self._t._check_editable(self.name)
        self._sheet.set(ref, value, allow_formula=overwrite_formula)
        self._t._report.cells += 1

    def update(self, values: dict, *, overwrite_formula: bool = False) -> None:
        """Write several cells; all-or-nothing validation is done by Template.apply_spec."""
        for ref, value in values.items():
            self.set(ref, value, overwrite_formula=overwrite_formula)

    # -- form controls ----------------------------------------------------
    def controls(self) -> list[ControlInfo]:
        """CheckBox/Radio controls with the key to address each one."""
        return [
            ControlInfo(self._sheet.preferred_key(c), c.name, c.kind, c.link or c.extra.get("group_link"),
                        _control_value(c), c.group)
            for c in self._sheet.controls() if c.kind in CONTROL_KINDS
        ]

    def set_control(self, key: str, value) -> None:
        """key = name | name#n | link:<cell> | id:<shapeId>.
        CheckBox: True / False / "mixed". Radio: True selects it within its group."""
        self._t._check_editable(self.name)
        self._sheet.set_control(key, value)
        self._t._report.controls += 1

    def check(self, key: str, value: bool | str = True) -> None:
        self.set_control(key, value)

    def uncheck(self, key: str) -> None:
        self.set_control(key, False)

    def select(self, key: str) -> None:
        """Select a radio button (the others in its group are cleared)."""
        if self._sheet.find_control(key).kind != "Radio":
            raise XlsmError(f"{key!r} on {self.name!r} is not a radio button")
        self.set_control(key, True)

    # -- visibility -------------------------------------------------------
    def hide(self, very: bool = False) -> None:
        self._t.hide(self.name, very=very)

    def show(self) -> None:
        self._t.show(self.name)

    def __repr__(self) -> str:
        return f"<SheetEditor {self.name!r} state={self.state}>"


class Template:
    """An .xlsm template opened for editing. The source file is never modified."""

    def __init__(self, path: str | os.PathLike):
        self.path = os.fspath(path)
        self._wb = Workbook(self.path)
        self._to_delete: list[str] = []
        self._saved = False
        self._report = ImportReport(output="")

    # -- sheets -----------------------------------------------------------
    @property
    def sheet_names(self) -> list[str]:
        return [n for n in self._wb.sheet_names() if n not in self._to_delete]

    def __getitem__(self, name: str) -> SheetEditor:
        sheet = self._wb.sheet(name)
        if sheet.name in self._to_delete:
            raise XlsmError(f"Sheet {sheet.name!r} is scheduled for deletion")
        return SheetEditor(self, sheet)

    def __contains__(self, name: str) -> bool:
        try:
            self[name]
            return True
        except XlsmError:
            return False

    def __iter__(self):
        return (self[n] for n in self.sheet_names)

    def hide(self, name: str, very: bool = False) -> None:
        sheet = self._wb.sheet(name)
        self._check_editable(sheet.name)
        self._wb.set_state(sheet.name, "veryHidden" if very else "hidden")
        self._report.states += 1

    def show(self, name: str) -> None:
        sheet = self._wb.sheet(name)
        self._check_editable(sheet.name)
        self._wb.set_state(sheet.name, "visible")
        self._report.states += 1

    def delete(self, name: str) -> None:
        """Schedule a sheet for deletion (done on save, after all edits).
        Prefer hide(): deleting leaves the sheet's VBA module behind in vbaProject.bin."""
        sheet = self._wb.sheet(name)
        self._check_editable(sheet.name)
        if sheet.name not in self._to_delete:
            self._to_delete.append(sheet.name)

    @property
    def warnings(self) -> list[str]:
        return list(self._wb.warnings)

    # -- bulk -------------------------------------------------------------
    def apply_spec(self, spec: dict) -> None:
        """Apply a JSON-style spec (see xlsm_safe.spec). Every entry is tried and all
        problems are reported together; on error do not save this Template."""
        errors: list[str] = []

        def attempt(fn, *args):
            try:
                fn(*args)
            except (XlsmError, ValueError) as e:
                errors.append(str(e))

        for name in spec.get("delete_sheets", []):
            attempt(self.delete, name)
        for name, entry in spec.get("sheets", {}).items():
            try:
                sheet = self[name]
            except XlsmError as e:
                errors.append(str(e))
                continue
            # Controls first: their linked cells may be overridden by explicit cell values.
            for key, value in sorted(entry.get("controls", {}).items(), key=lambda kv: bool(kv[1])):
                attempt(sheet.set_control, key, value)
            for ref, value in entry.get("cells", {}).items():
                attempt(sheet.set, ref, value)
            for ref, value in entry.get("formula_overwrites", {}).items():
                attempt(lambda r, v: sheet.set(r, v, overwrite_formula=True), ref, value)
            if "state" in entry:
                attempt(self._set_state, sheet.name, entry["state"])
        if errors:
            raise XlsmError("Import aborted, nothing was written:\n  - " + "\n  - ".join(errors))

    def _set_state(self, name: str, state: str) -> None:
        self._wb.set_state(name, state)
        self._report.states += 1

    # -- save -------------------------------------------------------------
    def _check_editable(self, sheet_name: str) -> None:
        if self._saved:
            raise XlsmError("This Template was already saved; open the template again for a new file")
        if sheet_name in self._to_delete:
            raise XlsmError(f"Sheet {sheet_name!r} is scheduled for deletion")

    def save(self, output: str | os.PathLike, *, verify: bool = True) -> ImportReport:
        """Write the result to `output` (never the template itself).

        With verify=True (default) the file is checked against the template first;
        if the check fails nothing is written and XlsmError is raised.
        """
        output = os.fspath(output)
        if self._saved:
            raise XlsmError("This Template was already saved; open the template again for a new file")
        if os.path.abspath(output) == os.path.abspath(self.path) or (
                os.path.exists(output) and os.path.samefile(output, self.path)):
            raise XlsmError("Output must be a new file; the template is never modified in place")

        for name in self._to_delete:
            self._wb.delete_sheet(name)
            self._report.deleted += 1
        self._saved = True

        out_dir = os.path.dirname(os.path.abspath(output))
        fd, tmp = tempfile.mkstemp(suffix=".xlsm", prefix=".xlsm-safe-", dir=out_dir)
        os.close(fd)
        try:
            self._wb.save(tmp)
            result = verify_files(self.path, tmp) if verify else None
            if result is not None and not result["ok"]:
                raise XlsmError("Verification failed, nothing was written:\n  - " + "\n  - ".join(result["problems"]))
            os.replace(tmp, output)
        finally:
            if os.path.exists(tmp):
                os.remove(tmp)

        report = self._report
        report.output = output
        report.warnings = self.warnings
        report.changed_parts = self._wb.pkg.changed_parts()
        report.verify = result
        return report

    def __repr__(self) -> str:
        return f"<Template {self.path!r} sheets={len(self.sheet_names)}>"

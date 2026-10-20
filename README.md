# xlsm-safe

Import data into `.xlsm` templates (検査成績書) **without breaking the file structure**:
option buttons, checkboxes, AutoShapes/images, VBA macros, formulas and Japanese text are preserved.

## Installation

Requires Python 3.8 or later.

```bash
python -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -e .                   # add ".[dev]" to run the tests: pip install -e ".[dev]"
```

The package is installed in editable mode (`-e`), so changes to the code in `xlsm_safe/` take effect
immediately without reinstalling. Installing provides the `xlsm-safe` command and lets you
`import xlsm_safe` from Python.

## Usage

```bash
# 1. List sheets and controls (name, type, linked cell, radio group)
xlsm-safe inspect data/INSP_2205_PALLET-RACK_A_J_VER1700.xlsm

# 2. (optional) Derive a spec from a filled-in example to see which cells need to be filled
xlsm-safe extract data/INSP_X102_...xlsm "data/例)..._X102_...xlsm" -o spec.json

# 3. Import: template + spec -> new file (the template is never modified)
xlsm-safe apply data/INSP_X102_...xlsm spec.json -o out.xlsm

# 4. Check the structure (apply already runs this step); use --expect to compare content with a reference file
xlsm-safe verify data/INSP_X102_...xlsm out.xlsm --expect "data/例)...xlsm"
```

### Spec format (JSON)

```json
{
  "sheets": {
    "表紙":         {"cells": {"Z11": 94715, "F9": "タカラスタンダード新潟工場向け自動倉庫"}},
    "M01-1 ラック": {"cells": {"S6": 4805}, "controls": {"option01": true}},
    "M04-2":        {"controls": {"link:BE15": true, "id:293891": false}},
    "M01-2 (2)":    {"state": "hidden"}
  }
}
```

- `cells`: string / number / `true`/`false` / `null` (clears the value). Write dates as Excel serial
  numbers (e.g. `46255`) so the cell's existing date format still applies. For a merged range,
  write to its top-left cell.

See [examples/electrical-work_spec.json](examples/electrical-work_spec.json) for a full example.

## Python API

You can use the tool directly from Python instead of the CLI:

```python
from xlsm_safe import Template, XlsmError

t = Template("data/INSP_240G_PC-CRANE_A_J_VER2300.xlsm")   # the template is only read

cover = t["表紙"]                      # get a sheet by its exact name
cover["Z12"] = 94715                   # write one cell
cover.update({"F9": "タカラスタンダード新潟工場向け自動倉庫", "AB43": "立野"})

try:
    report = t.save("out.xlsm")        # write a temp file → verify → only then rename to out.xlsm
except XlsmError as e:
    print("Nothing written:", e)
else:
    print(report.cells, report.controls, report.deleted, report.warnings)
```

| API | Description |
|---|---|
| `Template(path)` | Opens a template. Nothing is written to disk until `save()` is called. |
| `t["sheet name"]` | Returns a `SheetEditor`. A missing or ambiguous sheet name raises `XlsmError`. |
| `t.sheet_names`, `"name" in t`, `for s in t` | List / check sheets. |
| `sheet["A1"]`, `sheet["A1"] = v`, `sheet.update({...})` | Read / write cells. Formula cells are refused. |
| `t.apply_spec(dict)` | Applies a JSON spec, like the CLI. Collects all errors and reports them together. |
| `t.save(output, verify=True)` | Returns an `ImportReport`. If verification fails it raises `XlsmError` and **no file is created**. Overwriting the template is not allowed. Each `Template` can be saved only once. |

The lower-level functions `apply(template, spec, output)`, `extract(...)`, `verify(...)` and
`inspect(...)` are still available.

## Verification

For all 3 template / example pairs in `data/`: `extract` → `apply` to the blank template → `verify --expect`.
Result: **0 differences** in cell values, in the state of 589 checkboxes/radio buttons, and in
hidden/deleted sheets (including the 4 sheets deleted in PC-CRANE). The radio-button linked-cell values
computed by the tool match the values saved by Excel 100%.
Re-run with `pytest` (requires `pip install -e ".[dev]"`).

This process also revealed two issues in the example files:

- Some **formulas in the example files differ from the template** (e.g. `INDIRECT(... AQ5)` → `AQ4` in `M03-2`).
  The tool does not import these cells and prints a warning during `extract`.
- Control coordinates in the example files are shifted by a few pixels compared with the template
  (219075 → 213360 EMU). Those files were saved on a machine with a different DPI/zoom, which is exactly
  what causes *shifted option buttons*. Files produced by the tool always keep the template's original coordinates.

## Known limitations

- **Deleting sheets**: when Excel itself deletes a sheet, it also removes that sheet's VBA module
  (`Sheet17`, …) from `vbaProject.bin`. The tool does **not** modify this binary file, so the orphaned
  module remains and the tool prints a warning. This has not been verified in real Excel (the build
  environment has no Excel). If you need to be completely safe, use `"state": "hidden"`
  (as the PALLET-RACK example file does), or open the file in Excel and save it once.
- **ActiveX controls** (e.g. the OptionButton on sheet 記入) are only listed; their state cannot be changed,
  because it is stored in a separate binary.
- **Adding/duplicating sheets** (`M01-2 (5)`…) is not supported. The templates already contain copies (2)…(4);
  hide or delete the ones you don't use.
- Rich text (multiple formats within one cell) is written as plain text using the cell's existing format.

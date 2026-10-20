"""Command line: inspect / extract / apply / verify."""

from __future__ import annotations

import argparse
import json
import sys

from .spec import apply, extract
from .verify import inspect, verify
from .workbook import XlsmError


def _dump(obj, path: str | None = None) -> None:
    text = json.dumps(obj, ensure_ascii=False, indent=2)
    if path:
        with open(path, "w", encoding="utf-8") as f:
            f.write(text + "\n")
    else:
        print(text)


def main(argv: list[str] | None = None) -> int:
    p = argparse.ArgumentParser(prog="xlsm-safe", description="Import data into .xlsm templates without breaking them")
    sub = p.add_subparsers(dest="cmd", required=True)

    s = sub.add_parser("inspect", help="list sheets and form controls")
    s.add_argument("file")

    s = sub.add_parser("extract", help="derive an import spec from a filled example")
    s.add_argument("template")
    s.add_argument("filled")
    s.add_argument("-o", "--output")

    s = sub.add_parser("apply", help="write a spec into a copy of the template")
    s.add_argument("template")
    s.add_argument("spec")
    s.add_argument("-o", "--output", required=True)
    s.add_argument("--no-verify", action="store_true")

    s = sub.add_parser("verify", help="check that output kept the template structure")
    s.add_argument("template")
    s.add_argument("output")
    s.add_argument("--expect", help="filled file whose values/controls the output must match")

    a = p.parse_args(argv)
    try:
        if a.cmd == "inspect":
            _dump(inspect(a.file))
        elif a.cmd == "extract":
            spec, warnings = extract(a.template, a.filled)
            _dump(spec, a.output)
            for w in warnings:
                print("warning:", w, file=sys.stderr)
        elif a.cmd == "apply":
            with open(a.spec, encoding="utf-8") as f:
                spec = json.load(f)
            if a.output == a.template:
                raise XlsmError("Output must be a new file; the template is never modified in place")
            # apply() verifies before the file lands on disk and raises on failure.
            _dump(apply(a.template, spec, a.output, verify=not a.no_verify))
        elif a.cmd == "verify":
            r = verify(a.template, a.output, a.expect)
            _dump(r)
            return 0 if r["ok"] else 2
    except XlsmError as e:
        print(f"error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())

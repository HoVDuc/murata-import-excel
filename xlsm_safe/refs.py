"""Cell reference helpers and formula reference rewriting."""

from __future__ import annotations

import re

_REF = re.compile(r"^\$?([A-Za-z]{1,3})\$?(\d+)$")


def col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n: int) -> str:
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def split_ref(ref: str) -> tuple[int, int]:
    """'AB12' -> (col=28, row=12)."""
    m = _REF.match(ref.strip())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return col_to_num(m.group(1)), int(m.group(2))


def norm_ref(ref: str) -> str:
    c, r = split_ref(ref)
    return f"{num_to_col(c)}{r}"


def parse_range(rng: str) -> tuple[int, int, int, int]:
    a, _, b = rng.partition(":")
    c1, r1 = split_ref(a)
    c2, r2 = split_ref(b or a)
    return min(c1, c2), min(r1, r2), max(c1, c2), max(r1, r2)


_PLAIN_SHEET = re.compile(r"^[A-Za-z_\u0080-￿][\w.\u0080-￿]*$")
_CELL_OR_RANGE = r"(?:\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?|\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}|\$?\d+:\$?\d+)"


def sheet_ref_pattern(sheet: str) -> re.Pattern:
    """Matches references like 'Sheet (2)'!$A$1 or Sheet!A1:B2 to the given sheet.

    References into external workbooks ([1]Sheet!A1) are not matched.
    """
    quoted = "'" + sheet.replace("'", "''") + "'"
    alts = [re.escape(quoted)]
    if _PLAIN_SHEET.match(sheet):
        alts.append(re.escape(sheet))
    return re.compile(r"(?<![\w\]'])(?:%s)!%s" % ("|".join(alts), _CELL_OR_RANGE))


def refers_to_sheet(formula: str, sheet: str) -> bool:
    return bool(sheet_ref_pattern(sheet).search(formula))


def replace_sheet_refs(formula: str, sheet: str) -> str:
    """Replace references to a deleted sheet with #REF! (what Excel does)."""
    return sheet_ref_pattern(sheet).sub("#REF!", formula)

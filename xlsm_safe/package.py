"""Low-level OPC (zip) package access that rewrites only the parts we touch.

Every part that is not explicitly modified is written back byte-for-byte,
so VBA projects, VML drawings, form controls, ActiveX and printer settings
are never re-serialized.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass

from lxml import etree

REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

# Parsing must not drop whitespace-only text (e.g. a cell containing " ").
_PARSER = etree.XMLParser(remove_blank_text=False, resolve_entities=False, huge_tree=True)


@dataclass
class Rel:
    id: str
    type: str
    target: str
    external: bool
    part: str | None  # resolved absolute part name for internal targets


def rels_path(part: str) -> str:
    d, b = posixpath.split(part)
    return posixpath.join(d, "_rels", b + ".rels")


class Package:
    def __init__(self, path: str):
        self.path = path
        with zipfile.ZipFile(path) as z:
            self.infos = z.infolist()
            self.data = {i.filename: z.read(i.filename) for i in self.infos}
        self._trees: dict[str, etree._Element] = {}
        self._dirty: set[str] = set()
        self._removed: set[str] = set()

    # -- raw access -------------------------------------------------------
    def exists(self, name: str) -> bool:
        return name in self.data and name not in self._removed

    def read(self, name: str) -> bytes:
        if name in self._trees and name in self._dirty:
            return self._serialize(name)
        return self.data[name]

    def write(self, name: str, data: bytes) -> None:
        self._trees.pop(name, None)
        self.data[name] = data
        self._dirty.add(name)
        self._removed.discard(name)

    def remove(self, name: str) -> None:
        self._removed.add(name)
        self._trees.pop(name, None)
        self._dirty.discard(name)

    def names(self) -> list[str]:
        return [i.filename for i in self.infos if i.filename not in self._removed]

    # -- XML access -------------------------------------------------------
    def xml(self, name: str) -> etree._Element:
        """Parsed root of an XML part (cached). Call touch() after editing."""
        if name not in self._trees:
            self._trees[name] = etree.fromstring(self.data[name], _PARSER)
        return self._trees[name]

    def touch(self, name: str) -> None:
        self._dirty.add(name)

    def _serialize(self, name: str) -> bytes:
        return etree.tostring(self._trees[name], xml_declaration=True, encoding="UTF-8", standalone=True)

    # -- relationships ----------------------------------------------------
    def rels(self, part: str) -> list[Rel]:
        rp = rels_path(part)
        if not self.exists(rp):
            return []
        base = posixpath.dirname(part)
        out = []
        for r in self.xml(rp):
            ext = r.get("TargetMode") == "External"
            target = r.get("Target")
            resolved = None
            if not ext:
                resolved = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(base, target))
            out.append(Rel(r.get("Id"), r.get("Type"), target, ext, resolved))
        return out

    def rel_by_id(self, part: str, rid: str) -> Rel | None:
        return next((r for r in self.rels(part) if r.id == rid), None)

    def remove_rel(self, part: str, rid: str) -> None:
        rp = rels_path(part)
        root = self.xml(rp)
        for r in list(root):
            if r.get("Id") == rid:
                root.remove(r)
        self.touch(rp)

    def reachable(self) -> set[str]:
        """All parts reachable from the package root via relationships."""
        seen: set[str] = set()
        stack = [r.part for r in self.rels("") if r.part]
        while stack:
            p = stack.pop()
            if p in seen or not self.exists(p):
                continue
            seen.add(p)
            rp = rels_path(p)
            if self.exists(rp):
                seen.add(rp)
                stack.extend(r.part for r in self.rels(p) if r.part)
        seen.add("_rels/.rels")
        seen.add("[Content_Types].xml")
        return seen

    def remove_content_type_overrides(self, parts: set[str]) -> None:
        name = "[Content_Types].xml"
        root = self.xml(name)
        for o in list(root):
            if o.tag == f"{{{CT_NS}}}Override" and o.get("PartName", "").lstrip("/") in parts:
                root.remove(o)
        self.touch(name)

    # -- save -------------------------------------------------------------
    def changed_parts(self) -> list[str]:
        return sorted(self._dirty - self._removed)

    def save(self, path: str) -> None:
        with zipfile.ZipFile(path, "w") as z:
            for info in self.infos:
                name = info.filename
                if name in self._removed:
                    continue
                data = self._serialize(name) if (name in self._dirty and name in self._trees) else self.data[name]
                zi = zipfile.ZipInfo(name, date_time=info.date_time)
                zi.compress_type = info.compress_type
                zi.external_attr = info.external_attr
                zi.create_system = info.create_system
                z.writestr(zi, data)

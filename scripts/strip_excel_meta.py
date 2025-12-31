#!/usr/bin/env python
"""
Strip personal/organization metadata from Excel files before commit.
Handles xlsx/xlsm/xltx/xltm (OOXML). For legacy xls, it emits a warning and leaves
as-is (hook will still fail if you prefer strict behavior; see STRICT_XLS flag).
"""

from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path
from typing import Iterable
import xml.etree.ElementTree as ET

OOXML_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
LEGACY_EXTS = {".xls"}
ALL_EXTS = OOXML_EXTS | LEGACY_EXTS

CORE_TEMPLATE = b"""<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n<cp:coreProperties xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\" xmlns:dc=\"http://purl.org/dc/elements/1.1/\" xmlns:dcterms=\"http://purl.org/dc/terms/\" xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">\n  <dc:title/>\n  <dc:subject/>\n  <dc:creator/>\n  <cp:keywords/>\n  <dc:description/>\n  <cp:lastModifiedBy/>\n  <cp:revision>0</cp:revision>\n  <dcterms:created xsi:type=\"dcterms:W3CDTF\">1900-01-01T00:00:00Z</dcterms:created>\n  <dcterms:modified xsi:type=\"dcterms:W3CDTF\">1900-01-01T00:00:00Z</dcterms:modified>\n  <cp:category/>\n</cp:coreProperties>\n"""

STRICT_XLS = False  # set True to fail the hook when .xls is encountered


def _clean_content_types(data: bytes) -> bytes:
    try:
        ET.register_namespace("", "http://schemas.openxmlformats.org/package/2006/content-types")
        root = ET.fromstring(data)
    except Exception:
        return data
    removed = False
    for child in list(root):
        part = child.attrib.get("PartName", "").lower()
        if part == "/docprops/custom.xml":
            root.remove(child)
            removed = True
    if not removed:
        return data
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _clean_app(data: bytes) -> bytes:
    try:
        ns = {"ap": "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"}
        ET.register_namespace("", ns["ap"])
        root = ET.fromstring(data)
    except Exception:
        return data
    changed = False
    for tag in ("Company", "Manager"):
        for node in root.findall(f".//ap:{tag}", ns):
            if node.text:
                node.text = ""
                changed = True
    if not changed:
        return data
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def strip_ooxml(path: Path) -> bool:
    changed = False
    with tempfile.TemporaryDirectory() as td:
        tmp_path = Path(td) / path.name
        with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(tmp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                data = zin.read(name)
                if name == "docProps/core.xml":
                    data = CORE_TEMPLATE
                    changed = True
                elif name == "docProps/custom.xml":
                    changed = True
                    continue  # drop entirely
                elif name == "[Content_Types].xml":
                    new_data = _clean_content_types(data)
                    if new_data != data:
                        changed = True
                        data = new_data
                elif name == "docProps/app.xml":
                    new_data = _clean_app(data)
                    if new_data != data:
                        changed = True
                        data = new_data
                zout.writestr(info, data)
        if changed:
            shutil.move(str(tmp_path), str(path))
    return changed


def strip_file(path: Path) -> bool:
    suffix = path.suffix.lower()
    if suffix in OOXML_EXTS:
        return strip_ooxml(path)
    if suffix in LEGACY_EXTS:
        msg = f"[strip-excel-meta] legacy .xls not scrubbed (convert to xlsx or scrub manually): {path}"
        if STRICT_XLS:
            raise SystemExit(msg)
        print(msg)
        return False
    return False


def iter_targets(paths: Iterable[str]) -> Iterable[Path]:
    for raw in paths:
        p = Path(raw)
        if p.suffix.lower() in ALL_EXTS and p.is_file():
            yield p


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Strip Excel metadata (xlsx/xlsm/xltx/xltm)")
    parser.add_argument("paths", nargs="+", help="Paths to process")
    args = parser.parse_args(argv)

    targets = list(iter_targets(args.paths))
    if not targets:
        return 0

    any_changed = False
    for path in targets:
        try:
            changed = strip_file(path)
            any_changed = any_changed or changed
        except SystemExit:
            raise
        except Exception as exc:
            print(f"[strip-excel-meta] failed on {path}: {exc}", file=sys.stderr)
            return 1
    return 0 if not any_changed else 0


if __name__ == "__main__":
    sys.exit(main())

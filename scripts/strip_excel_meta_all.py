#!/usr/bin/env python
"""
Recursively strip personal/organization metadata from all Excel files under a given directory.
Uses strip_excel_meta.strip_file for each target.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from strip_excel_meta import strip_file, ALL_EXTS


def iter_targets(root: Path):
    for p in root.rglob("*"):
        if p.is_file() and p.suffix.lower() in ALL_EXTS:
            yield p


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Strip Excel metadata recursively")
    parser.add_argument("root", nargs="?", default=".", help="Root directory to scan (default: .)")
    args = parser.parse_args(argv)

    root = Path(args.root).resolve()
    if not root.exists():
        print(f"Root does not exist: {root}", file=sys.stderr)
        return 1

    targets = list(iter_targets(root))
    if not targets:
        print("No Excel files found.")
        return 0

    changed = 0
    for p in targets:
        try:
            if strip_file(p):
                changed += 1
                print(f"Stripped: {p}")
            else:
                print(f"Checked (no change): {p}")
        except SystemExit:
            raise
        except Exception as exc:
            print(f"[ERROR] {p}: {exc}", file=sys.stderr)
            return 1

    print(f"Done. Processed {len(targets)} file(s); changed {changed}.")
    return 0


if __name__ == "__main__":
    sys.exit(main())

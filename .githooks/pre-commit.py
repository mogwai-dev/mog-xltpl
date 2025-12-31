#!/usr/bin/env python
import subprocess
import sys
from pathlib import Path


def main() -> int:
    repo_root = subprocess.check_output(["git", "rev-parse", "--show-toplevel"], text=True).strip()
    exts = {".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"}

    diff_bytes = subprocess.check_output([
        "git",
        "diff",
        "--cached",
        "--name-only",
        "-z",
        "--diff-filter=ACM",
    ])
    names = [n for n in diff_bytes.decode("utf-8").split("\0") if n]

    targets = []
    for name in names:
        p = Path(name)
        if p.suffix.lower() in exts and p.is_file():
            targets.append(name)

    if not targets:
        return 0

    script = str(Path(repo_root) / "scripts" / "strip_excel_meta.py")
    cmd = [sys.executable, script, *targets]
    res = subprocess.run(cmd)
    if res.returncode != 0:
        return res.returncode

    # Re-stage cleaned files
    add_res = subprocess.run(["git", "add", "--", *targets])
    return add_res.returncode


if __name__ == "__main__":
    sys.exit(main())

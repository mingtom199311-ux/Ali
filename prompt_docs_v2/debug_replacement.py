from __future__ import annotations

import re
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"


def show_context(text: str, pos: int, radius: int = 160) -> str:
    start = max(0, pos - radius)
    end = min(len(text), pos + radius)
    return repr(text[start:end])


def main() -> int:
    found = False
    print("=== Source file scan ===")
    for path in sorted(ROOT.glob("*.md")) + sorted(ROOT.glob("*.py")):
        text = path.read_text(encoding="utf-8")
        for marker, label in [("\ufffd", "U+FFFD"), ("\x00", "NUL")]:
            for m in re.finditer(marker, text):
                found = True
                print(f"SOURCE {path.name} {label} at {m.start()}: {show_context(text, m.start())}")

    print("=== DOCX XML scan ===")
    docs = sorted(OUTPUT.glob("*.docx"))
    if not docs:
        print("No DOCX files found")
        return 1
    for docx in docs:
        print(f"DOCX {docx.name}")
        with zipfile.ZipFile(docx, "r") as zf:
            for name in zf.namelist():
                if not (name.endswith(".xml") or name.endswith(".rels")):
                    continue
                raw = zf.read(name)
                try:
                    text = raw.decode("utf-8")
                except UnicodeDecodeError as exc:
                    found = True
                    print(f"  UTF-8 decode error in {name}: {exc}")
                    continue
                for marker, label in [("\ufffd", "U+FFFD"), ("\x00", "NUL")]:
                    for m in re.finditer(marker, text):
                        found = True
                        print(f"  {name} {label} at {m.start()}: {show_context(text, m.start())}")
    if not found:
        print("No replacement or NUL characters found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

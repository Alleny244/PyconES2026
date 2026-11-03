#!/usr/bin/env python3
"""Render the personal-prep markdown files to PDF.

    python tools/build_notes.py              # all of them
    python tools/build_notes.py QA.md        # just one

Requires: pip install markdown weasyprint
Outputs sit next to the source (QA.md -> QA.pdf) and are gitignored.
"""
import pathlib
import sys

import markdown
from weasyprint import CSS, HTML

ROOT = pathlib.Path(__file__).resolve().parent.parent
DEFAULTS = ["STUDY_NOTES.md", "SCRIPT.md", "QA.md", "STUDY_PLAN.md"]

CSS_TEXT = """
@page { size: A4; margin: 18mm 16mm 16mm 16mm;
  @bottom-center { content: counter(page); font-family: serif; font-size: 9pt; color: #777; } }
body { font-family: "DejaVu Serif", Georgia, serif; font-size: 10.5pt; line-height: 1.5; color: #1a1a1a; }
h1 { font-family: "DejaVu Sans", sans-serif; font-size: 19pt; color: #0b3d5c;
     border-bottom: 2px solid #0b3d5c; padding-bottom: 4px; margin-top: 0;
     page-break-before: always; }
h1:first-of-type { page-break-before: avoid; }
h2 { font-family: "DejaVu Sans", sans-serif; font-size: 13pt; color: #11557c;
     margin-top: 18px; page-break-after: avoid; }
h3 { font-family: "DejaVu Sans", sans-serif; font-size: 11pt; page-break-after: avoid; }
p, li { orphans: 2; widows: 2; }
code { font-family: "DejaVu Sans Mono", monospace; font-size: 9pt; background: #f2f4f6;
       padding: 0 2px; border-radius: 2px; }
pre { background: #f2f4f6; border-left: 3px solid #11557c; padding: 7px 9px; font-size: 9pt;
      page-break-inside: avoid; }
pre code { background: none; }
table { border-collapse: collapse; width: 100%; font-size: 9.5pt; margin: 8px 0;
        page-break-inside: avoid; }
th, td { border: 1px solid #ccd; padding: 4px 7px; text-align: left; vertical-align: top; }
th { background: #eef3f7; font-family: "DejaVu Sans", sans-serif; }
hr { border: none; border-top: 1px solid #dde; margin: 16px 0; }
strong { color: #000; }
blockquote { border-left: 3px solid #bcd; margin-left: 0; padding-left: 10px; color: #444; }
"""


def render(md_path: pathlib.Path) -> None:
    body = markdown.markdown(
        md_path.read_text(encoding="utf-8"),
        extensions=["tables", "fenced_code", "sane_lists"],
    )
    out = md_path.with_suffix(".pdf")
    HTML(string=f"<html><head><meta charset='utf-8'></head><body>{body}</body></html>").write_pdf(
        out, stylesheets=[CSS(string=CSS_TEXT)]
    )
    print(f"wrote {out.relative_to(ROOT)}")


def main() -> None:
    names = sys.argv[1:] or DEFAULTS
    for name in names:
        path = (ROOT / name).resolve()
        if path.exists():
            render(path)
        else:
            print(f"skipped {name} (not found)")


if __name__ == "__main__":
    main()

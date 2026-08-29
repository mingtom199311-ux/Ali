from __future__ import annotations

import sys
from pathlib import Path

import fitz

ROOT = Path(__file__).resolve().parent
OUTPUT = ROOT / "output"
PREVIEW = ROOT / "preview"
PREVIEW.mkdir(parents=True, exist_ok=True)

REQUIRED = [
    "内容数据库重建",
    "答案边界复核",
    "阅读理解专项规则",
    "反幻觉与可追溯机制",
    "中文／日文字体双向审计",
    "共同终审清单",
    "最终执行命令",
]


def validate_pdf(pdf_path: Path) -> list[str]:
    errors: list[str] = []
    doc = fitz.open(pdf_path)
    if doc.page_count < 8:
        errors.append(f"page count too low: {doc.page_count}")
    all_text = "\n".join(page.get_text("text") for page in doc)
    for phrase in REQUIRED:
        if phrase not in all_text:
            errors.append(f"rendered PDF missing phrase: {phrase}")
    for bad in ("TOC \\o", "\ufffd"):
        if bad in all_text:
            errors.append(f"bad rendered text found: {bad}")
    for idx, page in enumerate(doc):
        text = page.get_text("text").strip()
        if len(text) < 6:
            errors.append(f"blank or near-blank page: {idx + 1}")
        rect = page.rect
        for block in page.get_text("blocks"):
            x0, y0, x1, y1 = block[:4]
            if x0 < -2 or y0 < -2 or x1 > rect.width + 2 or y1 > rect.height + 2:
                errors.append(f"text block outside page on page {idx + 1}")
                break

    sample_pages = sorted(set([0, 1, doc.page_count // 2, doc.page_count - 1]))
    stem = pdf_path.stem
    for idx in sample_pages:
        page = doc[idx]
        pix = page.get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
        pix.save(PREVIEW / f"{stem}_page_{idx + 1:03d}.png")
    doc.close()
    return errors


def main() -> int:
    pdfs = sorted(OUTPUT.glob("*.pdf"))
    lines = ["高考日语新题型配套解析最终Prompt｜渲染检查报告", ""]
    all_errors: list[str] = []
    if len(pdfs) != 2:
        all_errors.append(f"expected 2 PDFs, found {len(pdfs)}")
    for pdf in pdfs:
        errors = validate_pdf(pdf)
        lines.append(f"文件：{pdf.name}")
        with fitz.open(pdf) as d:
            lines.append(f"页数：{d.page_count}")
        if errors:
            lines.append("状态：FAIL")
            lines.extend(f"- {e}" for e in errors)
            all_errors.extend(f"{pdf.name}: {e}" for e in errors)
        else:
            lines.append("状态：PASS")
        lines.append("")
    report = OUTPUT / "render_validation_report.txt"
    report.write_text("\n".join(lines), encoding="utf-8")
    print(report.read_text(encoding="utf-8"))
    if all_errors:
        print("\n".join(all_errors), file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

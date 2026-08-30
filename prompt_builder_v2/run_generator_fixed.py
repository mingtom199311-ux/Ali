#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import importlib.util
import zipfile
from pathlib import Path

from lxml import etree

HERE = Path(__file__).resolve().parent
GENERATOR_PATH = HERE / "generate_prompts.py"

spec = importlib.util.spec_from_file_location("prompt_generator_base", GENERATOR_PATH)
if spec is None or spec.loader is None:
    raise RuntimeError(f"Unable to load generator: {GENERATOR_PATH}")

g = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def fixed_structural_validate(path: Path, version_name: str) -> dict:
    required = [
        "【最终可复制Prompt开始】",
        "【最终可复制Prompt结束】",
        "内部内容数据库",
        "答案边界专项审核",
        "唯一答案四重验证",
        "阅读理解专项规则",
        "反幻觉与实体一致性",
        "字符清洗、中日文字体、静态目录与版式",
        "共同终审清零清单",
        "正式任务输入区",
        "独立答案",
        version_name,
    ]
    forbidden_text = ["{{VERSION", "�", "\u200b"]

    with zipfile.ZipFile(path, "r") as zf:
        bad = zf.testzip()
        if bad:
            raise AssertionError(f"Corrupt member {bad} in {path.name}")
        xml_bytes = zf.read("word/document.xml")
        root = etree.fromstring(xml_bytes)

        # Only real OOXML Ruby elements are forbidden in the Prompt instruction document.
        # Literal strings such as "w:ruby" inside the instructional text are allowed.
        active_ruby = root.xpath(".//w:ruby", namespaces=NS)
        if active_ruby:
            raise AssertionError(
                f"Prompt instruction document contains {len(active_ruby)} active Ruby nodes"
            )

        # Visible explanatory text may mention TOC field syntax, but there must be no active TOC field.
        instr_texts = root.xpath(".//w:instrText/text()", namespaces=NS)
        active_toc = [x for x in instr_texts if "TOC" in x.upper()]
        if active_toc:
            raise AssertionError(f"Active TOC field found in {path.name}: {active_toc}")

        document_xml = xml_bytes.decode("utf-8")
        if g.CN_FONT not in document_xml or g.JP_FONT not in document_xml:
            raise AssertionError("Expected Chinese/Japanese font declarations missing")

    text = g.extract_doc_text(path)
    for item in required:
        if item not in text:
            raise AssertionError(f"Missing required marker: {item}")
    for item in forbidden_text:
        if item in text:
            raise AssertionError(f"Forbidden text found: {repr(item)}")
    if len(text) < 20000:
        raise AssertionError(f"Prompt content unexpectedly short: {len(text)}")
    if path.stat().st_size < 30000:
        raise AssertionError(f"DOCX unexpectedly small: {path.stat().st_size}")

    return {
        "filename": path.name,
        "bytes": path.stat().st_size,
        "characters": len(text),
        "required_markers": "pass",
        "zip_integrity": "pass",
        "font_declarations": "pass",
        "active_ruby_nodes_in_prompt_document": 0,
        "active_toc_fields": 0,
    }


g.structural_validate = fixed_structural_validate

if __name__ == "__main__":
    g.main()

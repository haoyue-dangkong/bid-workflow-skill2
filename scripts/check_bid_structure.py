#!/usr/bin/env python3
"""按标准 TXT 核对文件夹、DOCX、标题文本和标题格式。"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET

from parse_bid_tree import parse_file, relative_entries


W_NS = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"


def docx_audit(path: Path) -> tuple[str, list[str]]:
    with ZipFile(path) as archive:
        xml = archive.read("word/document.xml")
    root = ET.fromstring(xml)
    paragraphs = list(root.iter(f"{W_NS}p"))
    nonempty = []
    for paragraph in paragraphs:
        text = "".join(node.text or "" for node in paragraph.iter(f"{W_NS}t"))
        if text:
            nonempty.append((paragraph, text))

    errors: list[str] = []
    if len(paragraphs) != 1 or len(nonempty) != 1:
        errors.append(f"应只有一个标题段落，实际段落 {len(paragraphs)} 个、非空段落 {len(nonempty)} 个")
    if not nonempty:
        return "", errors

    paragraph, title = nonempty[0]
    jc = paragraph.find(f"{W_NS}pPr/{W_NS}jc")
    if jc is None or jc.get(f"{W_NS}val") != "center":
        errors.append("标题未居中")

    runs = list(paragraph.iter(f"{W_NS}r"))
    if len(runs) != 1:
        errors.append(f"标题应只有一个文字块，实际 {len(runs)} 个")
    if runs:
        run_properties = runs[0].find(f"{W_NS}rPr")
        if run_properties is None:
            errors.append("标题缺少字体格式")
        else:
            fonts = run_properties.find(f"{W_NS}rFonts")
            if fonts is None or fonts.get(f"{W_NS}eastAsia") != "宋体":
                errors.append("标题不是宋体")
            size = run_properties.find(f"{W_NS}sz")
            if size is None or size.get(f"{W_NS}val") != "24":
                errors.append("标题不是 12 磅")
            color = run_properties.find(f"{W_NS}color")
            if color is None or color.get(f"{W_NS}val", "").upper() != "000000":
                errors.append("标题不是黑色")
            if run_properties.find(f"{W_NS}b") is not None:
                errors.append("标题不应加粗")
    return title, errors


def main() -> int:
    parser = argparse.ArgumentParser(description="核对投标目录生成结果")
    parser.add_argument("input", type=Path, help="UTF-8 投标目录 TXT")
    parser.add_argument("output_root", type=Path, help="已创建的输出根目录")
    parser.add_argument("--json", action="store_true")
    args = parser.parse_args()

    try:
        nodes, warnings = parse_file(args.input)
        expected_dirs, expected_files = relative_entries(nodes)
    except (OSError, ValueError) as exc:
        print(f"解析失败：{exc}", file=sys.stderr)
        return 2

    root = args.output_root.resolve()
    if not root.exists() or not root.is_dir():
        print(f"输出根目录不存在：{root}", file=sys.stderr)
        return 2

    actual_dirs = {item.relative_to(root) for item in root.rglob("*") if item.is_dir()}
    actual_files = {item.relative_to(root) for item in root.rglob("*") if item.is_file()}
    expected_dir_set = set(expected_dirs)
    expected_file_set = {path for path, _ in expected_files}

    title_errors: list[str] = []
    for path, expected_title in expected_files:
        target = root / path
        if not target.exists():
            continue
        try:
            actual_title, format_errors = docx_audit(target)
        except Exception as exc:  # noqa: BLE001 - 将损坏文件作为核对结果报告
            title_errors.append(f"{path}: DOCX 无法读取（{exc}）")
            continue
        if actual_title != expected_title:
            title_errors.append(f"{path}: 标题不一致，期望 {expected_title!r}，实际 {actual_title!r}")
        title_errors.extend(f"{path}: {error}" for error in format_errors)

    report = {
        "output_root": str(root),
        "expected_folder_count": len(expected_dirs),
        "expected_file_count": len(expected_files),
        "missing_folders": sorted(str(path) for path in expected_dir_set - actual_dirs),
        "missing_files": sorted(str(path) for path in expected_file_set - actual_files),
        "unexpected_folders": sorted(str(path) for path in actual_dirs - expected_dir_set),
        "unexpected_files": sorted(str(path) for path in actual_files - expected_file_set),
        "title_errors": title_errors,
        "parse_warnings": warnings,
    }
    passed = not any(
        report[key]
        for key in ("missing_folders", "missing_files", "unexpected_folders", "unexpected_files", "title_errors")
    )
    report["passed"] = passed

    if args.json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        print("结构核对：通过" if passed else "结构核对：未通过")
        print(f"预期文件夹：{len(expected_dirs)} 个；预期 DOCX：{len(expected_files)} 个")
        for key, label in (
            ("missing_folders", "缺少文件夹"),
            ("missing_files", "缺少文件"),
            ("unexpected_folders", "多余文件夹"),
            ("unexpected_files", "多余文件"),
            ("title_errors", "标题或 DOCX 错误"),
        ):
            if report[key]:
                print(f"{label}：")
                for item in report[key]:
                    print(f"- {item}")
    return 0 if passed else 1


if __name__ == "__main__":
    raise SystemExit(main())

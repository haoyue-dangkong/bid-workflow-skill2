#!/usr/bin/env python3
"""在显式批准参数存在时，创建目录和只含标题的 DOCX。"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from xml.sax.saxutils import escape
from zipfile import ZIP_DEFLATED, ZipFile

from parse_bid_tree import parse_file, relative_entries, render_preview


INVALID_COMPONENTS = set('<>:"/\\|?*')
WINDOWS_RESERVED = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{number}" for number in range(1, 10)),
    *(f"LPT{number}" for number in range(1, 10)),
}


def validate_component(component: str) -> None:
    if not component or component in {".", ".."}:
        raise ValueError("路径名称为空或为保留名称")
    if component.endswith((" ", ".")):
        raise ValueError(f"路径名称不能以空格或句点结尾：{component!r}")
    if any(ord(char) < 32 for char in component):
        raise ValueError(f"路径名称含控制字符：{component!r}")
    if any(char in INVALID_COMPONENTS for char in component):
        raise ValueError(f"路径名称含 Windows 不允许的字符：{component!r}")
    stem = component.split(".", 1)[0].upper()
    if stem in WINDOWS_RESERVED:
        raise ValueError(f"路径名称是 Windows 保留名称：{component!r}")


def validate_paths(folders: list[Path], files: list[tuple[Path, str]]) -> None:
    seen: dict[str, tuple[str, Path]] = {}
    entries = [("文件夹", path) for path in folders]
    entries.extend(("Word文件", path) for path, _ in files)
    for kind, path in entries:
        for part in path.parts:
            validate_component(part)
        key = str(path).replace("/", "\\").casefold()
        if key in seen:
            previous_kind, previous_path = seen[key]
            raise ValueError(
                f"目录中存在重复或大小写冲突路径：{previous_kind} {previous_path}；{kind} {path}"
            )
        seen[key] = (kind, path)


def make_docx(path: Path, title: str) -> None:
    title_xml = escape(title)
    document_xml = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:jc w:val="center"/></w:pPr>
      <w:r>
        <w:rPr><w:rFonts w:ascii="SimSun" w:eastAsia="宋体" w:hAnsi="SimSun"/><w:sz w:val="24"/><w:color w:val="000000"/></w:rPr>
        <w:t xml:space="preserve">{title_xml}</w:t>
      </w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>'''
    content_types = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>'''
    root_rels = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>'''
    path.parent.mkdir(parents=True, exist_ok=True)
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", content_types)
        archive.writestr("_rels/.rels", root_rels)
        archive.writestr("word/document.xml", document_xml)


def main() -> int:
    parser = argparse.ArgumentParser(description="按审核通过的 TXT 创建投标目录和 DOCX")
    parser.add_argument("input", type=Path, help="UTF-8 投标目录 TXT")
    parser.add_argument("output_root", type=Path, help="独立输出根目录")
    parser.add_argument("--approved", action="store_true", help="用户已明确审核通过；没有此参数只预览")
    parser.add_argument("--indent-width", type=int, default=2, choices=(2, 4))
    args = parser.parse_args()

    try:
        nodes, warnings = parse_file(args.input, indent_width=args.indent_width)
        folders, files = relative_entries(nodes)
        validate_paths(folders, files)
    except (OSError, ValueError) as exc:
        print(f"解析失败：{exc}", file=sys.stderr)
        return 2

    print("目录预览：")
    print(render_preview(nodes))
    if warnings:
        print("\n说明：")
        for warning in warnings:
            print(f"- {warning}")

    if not args.approved:
        print("\n未检测到 --approved，预览阶段未创建任何实际文件。")
        return 0

    root = args.output_root.resolve()
    conflicts: list[str] = []
    for folder in folders:
        target = root / folder
        if target.exists() and not target.is_dir():
            conflicts.append(f"应为文件夹但已存在文件：{target}")
    for file_path, _ in files:
        target = root / file_path
        if target.exists():
            conflicts.append(f"文件已存在，拒绝覆盖：{target}")
    if conflicts:
        print("\n检测到冲突，未创建任何新文件：", file=sys.stderr)
        for conflict in conflicts:
            print(f"- {conflict}", file=sys.stderr)
        return 3

    root.mkdir(parents=True, exist_ok=True)
    created_folders = 0
    reused_folders = 0
    for folder in folders:
        target = root / folder
        if target.exists():
            reused_folders += 1
        else:
            target.mkdir(parents=True)
            created_folders += 1

    for file_path, title in files:
        make_docx(root / file_path, title)

    print(
        f"\n创建完成：新建文件夹 {created_folders} 个，"
        f"复用文件夹 {reused_folders} 个，DOCX {len(files)} 个。"
    )
    print(f"输出根目录：{root}")
    print("请继续使用 check_bid_structure.py 进行创建后核对。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

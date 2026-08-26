#!/usr/bin/env python3
"""解析通用的 [文件夹]/[Word文件] 投标目录 TXT。"""

from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Iterable


NODE_RE = re.compile(r"^(?P<indent>[ \t]*)\[(?P<kind>文件夹|Word文件)\]\s*(?P<name>.+?)\s*$")


@dataclass(frozen=True)
class Node:
    line: int
    level: int
    kind: str
    name: str


def _indent_level(text: str, indent_width: int) -> int:
    spaces = 0
    for char in text:
        spaces += indent_width if char == "\t" else 1
    if spaces % indent_width:
        raise ValueError(f"缩进不是 {indent_width} 个空格的整数倍")
    return spaces // indent_width


def normalize_docx_name(name: str) -> str:
    lower = name.lower()
    if lower.endswith(".docx"):
        return name
    if lower.endswith(".doc"):
        return name[:-4] + ".docx"
    return name + ".docx"


def title_from_word_name(name: str) -> str:
    lower = name.lower()
    if lower.endswith(".docx"):
        return name[:-5]
    if lower.endswith(".doc"):
        return name[:-4]
    return name


def parse_lines(lines: Iterable[str], indent_width: int = 2) -> tuple[list[Node], list[str]]:
    nodes: list[Node] = []
    warnings: list[str] = []
    stack: list[Node] = []
    in_notes = False

    for line_no, raw in enumerate(lines, start=1):
        text = raw.rstrip("\r\n")
        if not text.strip():
            continue
        if text.strip() == "依据和备注":
            in_notes = True
            continue
        if in_notes:
            continue

        match = NODE_RE.match(text)
        if not match:
            if nodes:
                warnings.append(f"第 {line_no} 行不是结构节点，已按说明文字忽略：{text.strip()}")
            continue

        try:
            level = _indent_level(match.group("indent"), indent_width)
        except ValueError as exc:
            raise ValueError(f"第 {line_no} 行：{exc}") from exc

        name = match.group("name").strip()
        if not name:
            raise ValueError(f"第 {line_no} 行：节点名称为空")
        if not nodes and level != 0:
            raise ValueError(f"第 {line_no} 行：第一个节点必须从第 0 级开始")
        if nodes and level > nodes[-1].level + 1:
            raise ValueError(f"第 {line_no} 行：层级从 {nodes[-1].level} 跳到了 {level}")

        while stack and stack[-1].level >= level:
            stack.pop()
        if stack and stack[-1].kind == "Word文件":
            raise ValueError(f"第 {line_no} 行：Word 文件不能包含子节点")
        if level > len(stack):
            raise ValueError(f"第 {line_no} 行：找不到第 {level} 级父节点")

        node = Node(line_no, level, match.group("kind"), name)
        nodes.append(node)
        stack.append(node)

    if not nodes:
        raise ValueError("没有找到任何目录节点")
    return nodes, warnings


def parse_file(path: Path, indent_width: int = 2) -> tuple[list[Node], list[str]]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except UnicodeDecodeError as exc:
        raise ValueError("TXT 不是 UTF-8 编码，或含有无法读取的字符") from exc
    return parse_lines(text.splitlines(), indent_width=indent_width)


def relative_entries(nodes: list[Node]) -> tuple[list[Path], list[tuple[Path, str]]]:
    """返回预期文件夹路径，以及（DOCX 路径、无扩展名标题）。"""
    folders: list[Path] = []
    files: list[tuple[Path, str]] = []
    stack: list[tuple[int, str, str]] = []

    for node in nodes:
        while stack and stack[-1][0] >= node.level:
            stack.pop()
        parent_parts = [item[1] for item in stack if item[2] == "文件夹"]
        if node.kind == "文件夹":
            rel = Path(*parent_parts, node.name)
            folders.append(rel)
            stack.append((node.level, node.name, node.kind))
        else:
            filename = normalize_docx_name(node.name)
            rel = Path(*parent_parts, filename)
            files.append((rel, title_from_word_name(node.name)))
            stack.append((node.level, node.name, node.kind))
    return folders, files


def render_preview(nodes: list[Node]) -> str:
    lines: list[str] = []
    for node in nodes:
        name = normalize_docx_name(node.name) if node.kind == "Word文件" else node.name
        lines.append(f"{'  ' * node.level}[{node.kind}] {name}")
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser(description="解析投标目录 TXT 并输出预览")
    parser.add_argument("input", type=Path, help="UTF-8 投标目录 TXT")
    parser.add_argument("--indent-width", type=int, default=2, choices=(2, 4))
    parser.add_argument("--json", action="store_true", help="输出机器可读 JSON")
    args = parser.parse_args()

    try:
        nodes, warnings = parse_file(args.input, indent_width=args.indent_width)
        folders, files = relative_entries(nodes)
    except (OSError, ValueError) as exc:
        print(f"解析失败：{exc}", file=sys.stderr)
        return 2

    if args.json:
        payload = {
            "input": str(args.input),
            "nodes": [asdict(node) for node in nodes],
            "folders": [str(path) for path in folders],
            "files": [{"path": str(path), "title": title} for path, title in files],
            "warnings": warnings,
        }
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0

    print("目录预览：")
    print(render_preview(nodes))
    print(f"\n文件夹：{len(folders)} 个；Word 文件：{len(files)} 个")
    if warnings:
        print("\n被忽略的说明行：")
        for warning in warnings:
            print(f"- {warning}")
    print("\n预览阶段未创建任何实际文件。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

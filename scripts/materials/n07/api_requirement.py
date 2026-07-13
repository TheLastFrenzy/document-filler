import re
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from materials.shared.ledger import ApiInterface, ParameterGroup


STOP_LABELS = {
    "共享任务开发代码检查",
    "共享接口命名规范性",
    "测试结果",
    "测试结论",
}


def _normalize_heading(value):
    text = re.sub(r"^\s*\d+\s*[、.．)]\s*", "", str(value or "").strip())
    return re.sub(r"\s+", "", text)


def read_docx_blocks(path):
    document = Document(path)
    blocks = []
    for child in document.element.body.iterchildren():
        if child.tag.endswith("}p"):
            paragraph = Paragraph(child, document)
            blocks.append({"type": "paragraph", "text": paragraph.text.strip()})
        elif child.tag.endswith("}tbl"):
            table = Table(child, document)
            matrix = [[cell.text.strip() for cell in row.cells] for row in table.rows]
            if matrix:
                blocks.append({"type": "table", "headers": matrix[0], "rows": matrix[1:]})
    return blocks


def _find_paragraph(blocks, text, start=0):
    target = _normalize_heading(text)
    for index in range(start, len(blocks)):
        block = blocks[index]
        if block["type"] == "paragraph" and _normalize_heading(block["text"]) == target:
            return index
    raise ValueError(f"自测报告中未找到{text}章节")


def _parameter_groups(blocks):
    groups = []
    label = ""
    for block in blocks:
        if block["type"] == "paragraph":
            text = block["text"].strip()
            normalized = text.rstrip("：:")
            if text.startswith("测试结果") or normalized in STOP_LABELS:
                break
            if text:
                label = normalized
            continue
        groups.append(
            ParameterGroup(
                label=label,
                headers=list(block["headers"]),
                rows=[list(row) for row in block["rows"]],
            )
        )
        label = ""
    return groups


def parse_api_report(report_path: Path, interfaces: list[ApiInterface]) -> list[ApiInterface]:
    blocks = read_docx_blocks(report_path)
    test_start = _find_paragraph(blocks, "测试内容")
    positions = {}
    for interface in interfaces:
        target = _normalize_heading(interface.chinese_name)
        for index in range(test_start + 1, len(blocks)):
            block = blocks[index]
            if block["type"] == "paragraph" and _normalize_heading(block["text"]) == target:
                positions[interface.chinese_name] = index
                break
    missing = [item.chinese_name for item in interfaces if item.chinese_name not in positions]
    if missing:
        raise ValueError(f"自测报告中未匹配到接口: {', '.join(missing)}")

    for index, interface in enumerate(interfaces):
        start = positions[interface.chinese_name] + 1
        end = positions[interfaces[index + 1].chinese_name] if index + 1 < len(interfaces) else len(blocks)
        section = blocks[start:end]
        input_index = _find_paragraph(section, "输入参数")
        output_index = _find_paragraph(section, "输出参数", input_index + 1)
        interface.input_groups = _parameter_groups(section[input_index + 1 : output_index])
        interface.output_groups = _parameter_groups(section[output_index + 1 :])
        if not interface.input_groups or not interface.output_groups:
            raise ValueError(f"接口{interface.chinese_name}缺少输入或输出参数表")
    return interfaces

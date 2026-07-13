import re
import tempfile
from pathlib import Path

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

from materials.shared.embedded_docx import extract_embedded_docx_by_work_order
from materials.shared.ledger import ApiInterface, ParameterGroup, read_api_work_orders
from materials.shared.word_sections import (
    clone_element,
    element_text,
    elements_between,
    find_heading,
    paragraph_element,
    replace_after,
    replace_between,
    table_element,
    update_toc_via_com,
)


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


def _parameter_descriptions(groups, limit=3):
    values = []
    for group in groups:
        for row in group.rows:
            value = row[1].strip() if len(row) > 1 else row[0].strip()
            if value and value not in values:
                values.append(value)
            if len(values) >= limit:
                return values
    return values


def _join_chinese(values):
    if not values:
        return "相关参数"
    if len(values) == 1:
        return values[0]
    return "、".join(values)


def build_interface_purpose(order, interface):
    inputs = _join_chinese(_parameter_descriptions(interface.input_groups))
    outputs = _join_chinese(_parameter_descriptions(interface.output_groups))
    name = interface.chinese_name
    if "令牌" in name:
        text = f"该接口接收{inputs}，返回{outputs}，供同一业务中的身份认证申请和核验接口调用。"
    elif "申请" in name:
        text = f"该接口接收{inputs}等认证申请信息，登记本次身份认证请求并返回{outputs}，供后续身份核验请求引用。"
    elif "请求" in name or "认证" in name:
        text = f"该接口接收{inputs}等核验信息，调用出入境身份认证服务并返回{outputs}，用于离境退税“掌上办”平台校验境外人员身份。"
    elif "查询" in name:
        text = f"该接口根据{inputs}查询业务数据并返回{outputs}，用于{order.title}相关信息核验。"
    else:
        text = f"该接口接收{inputs}并返回{outputs}，用于处理{order.title}对应的数据共享请求。"
    for banned in ("赋能", "彰显", "至关重要", "确保", "重要支撑"):
        text = text.replace(banned, "")
    return re.sub(r"\s+", "", text)


def _style_ids(document):
    return {
        name: document.styles[name].style_id
        for name in ("Heading 2", "Heading 3", "Normal", "Body Text")
    }


def _build_business_scene(document, orders, styles, business_heading, demand_heading):
    existing = elements_between(business_heading._p, demand_heading._p)
    first_paragraph = next(
        (element for element in existing if element.tag.endswith("}p") and element_text(element)),
        None,
    )
    titles = "、".join(order.title for order in orders)
    total = sum(order.program_count for order in orders)
    scene = f"在本服务周期内，围绕{titles}开展API接口开发，共提供{total}个API接口服务。"
    elements = []
    if first_paragraph is not None:
        elements.append(clone_element(first_paragraph))
    elements.append(paragraph_element(scene, styles["Normal"], 480))
    replace_between(business_heading._p, demand_heading._p, elements)


def _build_demand_section(orders, styles):
    request_count = len({order.demand_no for order in orders if order.demand_no})
    total = sum(order.program_count for order in orders)
    summary = (
        f"服务周期内，共有{request_count}张需求单，{len(orders)}张工单涉及{total}个API接口服务。"
        "具体需求单、工单和产出如下表："
    )
    headers = ["序号", "对应需求单编号", "对应工单编号", "工单内容", "涉及共享任务数量（个）"]
    rows = [
        [str(index), order.demand_no, order.work_order_no, order.title, str(order.program_count)]
        for index, order in enumerate(orders, start=1)
    ]
    rows.append(["总计", "总计", "", "", str(total)])
    elements = [
        paragraph_element(summary, styles["Normal"], 480),
        table_element(headers, rows, [650, 2100, 2100, 3300, 850]),
    ]
    for order in orders:
        elements.append(
            paragraph_element(f"{order.work_order_no}_{order.title}", styles["Heading 3"])
        )
        elements.append(
            paragraph_element(f"需求口径：{order.description}", styles["Normal"], 480)
        )
    return elements


def _build_api_section(orders, styles):
    elements = []
    for order in orders:
        for interface in order.interfaces:
            interface.purpose = build_interface_purpose(order, interface)
            elements.append(paragraph_element(interface.chinese_name, styles["Heading 3"]))
            elements.append(paragraph_element(interface.purpose, styles["Normal"], 480))
            elements.append(paragraph_element("计划供数方式：API接口对外服务", styles["Normal"]))
            elements.append(paragraph_element("计划更新频率：无", styles["Normal"]))
            elements.append(paragraph_element("接口输入参数：", styles["Normal"]))
            for group in interface.input_groups:
                if group.label:
                    elements.append(paragraph_element(f"{group.label}：", styles["Body Text"]))
                elements.append(table_element(group.headers, group.rows))
            elements.append(paragraph_element("接口输出参数：", styles["Normal"]))
            for group in interface.output_groups:
                if group.label:
                    elements.append(paragraph_element(f"{group.label}：", styles["Body Text"]))
                elements.append(table_element(group.headers, group.rows))
    return elements


def build_api_requirement_document(excel_path, service_dir, template_path, output_path):
    orders = read_api_work_orders(excel_path, service_dir)
    with tempfile.TemporaryDirectory(prefix="document_filler_api_requirement_") as temp_dir:
        reports = extract_embedded_docx_by_work_order(
            excel_path,
            orders,
            "自测报告附件",
            Path(temp_dir) / "reports",
        )
        for order in orders:
            order.self_report_path = reports[order.work_order_no]
            parse_api_report(order.self_report_path, order.interfaces)

    document = Document(template_path)
    styles = _style_ids(document)
    business_heading = find_heading(document, "Heading 2", "业务场景")
    demand_intro_heading = find_heading(document, "Heading 2", "需求说明")
    demand_heading = find_heading(document, "Heading 2", "需求清单")
    shared_heading = find_heading(document, "Heading 2", "共享开放方案")
    api_heading = find_heading(document, "Heading 2", "API接口清单")

    _build_business_scene(document, orders, styles, business_heading, demand_intro_heading)
    replace_between(demand_heading._p, shared_heading._p, _build_demand_section(orders, styles))
    replace_after(api_heading._p, _build_api_section(orders, styles))

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    update_toc_via_com(output)
    return str(output)

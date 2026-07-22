import re
import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

from materials.shared.embedded_docx import extract_embedded_docx_by_work_order
from materials.shared.ledger import ApiInterface, ParameterGroup, read_api_work_orders
from materials.shared.word_sections import (
    center_table_header_rows,
    clone_element,
    clone_paragraph_with_text,
    clone_table_with_data,
    element_text,
    elements_between,
    find_heading,
    replace_after,
    replace_between,
    update_toc_via_com,
)


STOP_LABELS = {
    "共享任务开发代码检查",
    "共享接口命名规范性",
    "测试结果",
    "测试结论",
}

CHINESE_DESCRIPTION_HEADERS = (
    "名称",
    "参数说明",
    "字段中文名",
    "字段中文名称",
    "中文名称",
    "数据项名称",
    "数据项说明",
    "字段注释",
    "字段说明",
    "说明",
    "备注",
)
PARAMETER_NAME_HEADERS = (
    "参数名",
    "参数项",
    "参数",
    "数据项",
    "字段英文名",
    "英文名称",
    "字段名",
    "字段名称",
)
LOW_PRIORITY_NAME_HEADERS = (
    "字段名称",
    "参数名称",
)
NON_DESCRIPTION_HEADERS = {
    "序号",
    "参数",
    "参数项",
    "参数名",
    "字段英文名",
    "英文名称",
    "类型",
    "字段类型",
    "是否必选",
    "是否必填",
    "默认",
    "不可为空",
    "唯一",
    "主键/外键",
    "测试数据1",
    "结果数据",
}
NORMALIZED_PARAMETER_HEADERS = ["序号", "参数名", "名称"]


@dataclass(frozen=True)
class ApiRequirementPrototypes:
    business_before: tuple
    business_scene: object
    business_after: tuple
    demand_summary: object
    demand_table: object
    work_order_heading: object
    work_order_body: object
    interface_heading: object
    interface_purpose: object
    supply_method: object
    update_frequency: object
    input_label: object
    output_label: object
    input_parameter_table: object
    output_parameter_table: object


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
                grid = table._tbl.tblGrid
                widths = (
                    [int(column.get("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}w")) for column in grid.gridCol_lst]
                    if grid is not None
                    else []
                )
                blocks.append(
                    {"type": "table", "headers": matrix[0], "rows": matrix[1:], "column_widths": widths}
                )
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
                column_widths=list(block.get("column_widths", [])),
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


def _header_key(value):
    return re.sub(r"\s+", "", str(value or "").strip().rstrip("：:"))


def _has_chinese(value):
    return bool(re.search(r"[\u4e00-\u9fff]", str(value or "")))


def _clean_parameter_description(value):
    return re.sub(r"\s+", "", str(value or "").strip())


def _preferred_parameter_description(headers, row):
    cells = [_clean_parameter_description(cell) for cell in row]
    indexed = [
        (_header_key(headers[index]) if index < len(headers) else "", cell)
        for index, cell in enumerate(cells)
        if cell
    ]

    preferred = [
        cell
        for header, cell in indexed
        if header in CHINESE_DESCRIPTION_HEADERS
    ]
    low_priority_names = [
        cell
        for header, cell in indexed
        if header in LOW_PRIORITY_NAME_HEADERS
    ]
    fallback_chinese = [
        cell
        for header, cell in indexed
        if header not in NON_DESCRIPTION_HEADERS and _has_chinese(cell)
    ]

    for candidates in (preferred, fallback_chinese, low_priority_names):
        for cell in candidates:
            if _has_chinese(cell):
                return cell
    for candidates in (preferred, low_priority_names):
        for cell in candidates:
            return cell
    if len(cells) > 1 and cells[1]:
        return cells[1]
    return cells[0] if cells else ""


def _cell(row, index):
    if index is None or index >= len(row):
        return ""
    return str(row[index] or "").strip()


def _header_index(headers, candidates):
    normalized = [_header_key(header) for header in headers]
    candidate_set = {_header_key(candidate) for candidate in candidates}
    for index, value in enumerate(normalized):
        if value in candidate_set:
            return index
    return None


def _fallback_parameter_index(headers):
    sequence_index = _header_index(headers, ["序号"])
    if sequence_index == 0 and len(headers) > 1:
        return 1
    return 0


def _parameter_name(headers, row):
    index = _header_index(headers, PARAMETER_NAME_HEADERS)
    value = _cell(row, index)
    if value:
        return value
    fallback_index = _fallback_parameter_index(headers)
    return _cell(row, fallback_index)


def normalized_parameter_rows(groups):
    rows = []
    for group in groups:
        for row in group.rows:
            parameter_name = _parameter_name(group.headers, row)
            description = _preferred_parameter_description(group.headers, row)
            if not parameter_name and not description:
                continue
            rows.append([str(len(rows) + 1), parameter_name, description])
    return rows


def _parameter_descriptions(groups, limit=3):
    values = []
    for group in groups:
        for row in group.rows:
            value = _preferred_parameter_description(group.headers, row)
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


def _sentence_fragment(value):
    text = re.sub(r"\s+", "", str(value or "").strip())
    return text.strip("。；;，,、 ")


def _summary_sentence_fragment(value):
    text = re.sub(r"\s+", "", str(value or "").strip())
    first = re.split(r"[。；;]", text, maxsplit=1)[0]
    if len(first) > 80:
        first = first[:80]
    return first.strip("。；;，,、 ")


def _work_order_background(order):
    title = _sentence_fragment(order.title)
    description = _summary_sentence_fragment(order.description)
    if title and description and description not in title:
        return f"围绕{title}，{description}，"
    if title:
        return f"围绕{title}，"
    if description:
        return f"结合{description}，"
    return ""


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
    text = f"{_work_order_background(order)}{text}"
    for banned in ("赋能", "彰显", "至关重要", "确保", "重要支撑"):
        text = text.replace(banned, "")
    return re.sub(r"\s+", "", text)


def _paragraph_style_id(element):
    properties = element.find(qn("w:pPr"))
    style = properties.find(qn("w:pStyle")) if properties is not None else None
    return style.get(qn("w:val")) if style is not None else None


def _nonempty_paragraphs(elements):
    return [
        element
        for element in elements
        if element.tag == qn("w:p") and element_text(element)
    ]


def _require_prototype(value, name):
    if value is None:
        raise ValueError(f"模板中未找到{name}格式原型")
    return value


def _capture_template_prototypes(
    document,
    business_heading,
    demand_intro_heading,
    demand_heading,
    shared_heading,
    api_heading,
):
    heading_3_style = document.styles["Heading 3"].style_id

    business_elements = elements_between(business_heading._p, demand_intro_heading._p)
    business_paragraphs = _nonempty_paragraphs(business_elements)
    if not business_paragraphs:
        raise ValueError("模板业务场景章节缺少动态段落格式原型")
    dynamic_scene = business_paragraphs[1] if len(business_paragraphs) > 1 else business_paragraphs[0]
    dynamic_scene_index = business_elements.index(dynamic_scene)
    business_before = tuple(business_elements[:dynamic_scene_index])
    business_after = tuple(business_elements[dynamic_scene_index + 1 :])

    demand_elements = elements_between(demand_heading._p, shared_heading._p)
    demand_paragraphs = _nonempty_paragraphs(demand_elements)
    demand_summary = demand_paragraphs[0] if demand_paragraphs else None
    demand_table = next((item for item in demand_elements if item.tag == qn("w:tbl")), None)
    work_order_heading = next(
        (
            item
            for item in demand_elements
            if item.tag == qn("w:p") and _paragraph_style_id(item) == heading_3_style
        ),
        None,
    )
    work_order_body = None
    if work_order_heading is not None:
        start = demand_elements.index(work_order_heading) + 1
        work_order_body = next(
            (
                item
                for item in demand_elements[start:]
                if item.tag == qn("w:p") and element_text(item)
            ),
            None,
        )

    api_elements = []
    current = api_heading._p.getnext()
    while current is not None and current.tag != qn("w:sectPr"):
        api_elements.append(current)
        current = current.getnext()
    interface_heading = next(
        (
            item
            for item in api_elements
            if item.tag == qn("w:p") and _paragraph_style_id(item) == heading_3_style
        ),
        None,
    )
    interface_purpose = None
    if interface_heading is not None:
        start = api_elements.index(interface_heading) + 1
        interface_purpose = next(
            (
                item
                for item in api_elements[start:]
                if item.tag == qn("w:p") and element_text(item)
            ),
            None,
        )

    def paragraph_named(text):
        return next(
            (
                item
                for item in api_elements
                if item.tag == qn("w:p") and element_text(item) == text
            ),
            None,
        )

    input_label = paragraph_named("接口输入参数：")
    output_label = paragraph_named("接口输出参数：")

    def first_table_after(paragraph):
        if paragraph is None:
            return None
        start = api_elements.index(paragraph) + 1
        return next(
            (item for item in api_elements[start:] if item.tag == qn("w:tbl")),
            None,
        )

    return ApiRequirementPrototypes(
        business_before=business_before,
        business_scene=dynamic_scene,
        business_after=business_after,
        demand_summary=_require_prototype(demand_summary, "需求清单说明段落"),
        demand_table=_require_prototype(demand_table, "需求清单表格"),
        work_order_heading=_require_prototype(work_order_heading, "工单标题"),
        work_order_body=_require_prototype(work_order_body, "工单需求口径"),
        interface_heading=_require_prototype(interface_heading, "API接口标题"),
        interface_purpose=_require_prototype(interface_purpose, "API接口说明"),
        supply_method=_require_prototype(
            paragraph_named("计划供数方式：API接口对外服务"), "计划供数方式"
        ),
        update_frequency=_require_prototype(
            paragraph_named("计划更新频率：无"), "计划更新频率"
        ),
        input_label=_require_prototype(input_label, "接口输入参数标签"),
        output_label=_require_prototype(output_label, "接口输出参数标签"),
        input_parameter_table=_require_prototype(
            first_table_after(input_label), "API输入参数表格"
        ),
        output_parameter_table=_require_prototype(
            first_table_after(output_label), "API输出参数表格"
        ),
    )


def _build_business_scene(orders, prototypes):
    titles = "、".join(order.title for order in orders)
    total = sum(order.program_count for order in orders)
    scene = f"在本服务周期内，围绕{titles}开展API接口开发，共提供{total}个API接口服务。"
    return [
        *(clone_element(element) for element in prototypes.business_before),
        clone_paragraph_with_text(prototypes.business_scene, scene),
        *(clone_element(element) for element in prototypes.business_after),
    ]


def _build_demand_section(orders, prototypes):
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
    elements = [
        clone_paragraph_with_text(prototypes.demand_summary, summary),
        clone_table_with_data(
            prototypes.demand_table,
            headers,
            rows,
            footer=["总计", "", "", "", str(total)],
        ),
    ]
    for order in orders:
        elements.append(
            clone_paragraph_with_text(
                prototypes.work_order_heading,
                f"{order.work_order_no}_{order.title}",
            )
        )
        elements.append(
            clone_paragraph_with_text(
                prototypes.work_order_body,
                f"需求口径：{order.description}",
            )
        )
    return elements


def _build_api_section(orders, prototypes):
    elements = []
    for order in orders:
        for interface in order.interfaces:
            interface.purpose = build_interface_purpose(order, interface)
            elements.append(
                clone_paragraph_with_text(prototypes.interface_heading, interface.chinese_name)
            )
            elements.append(
                clone_paragraph_with_text(prototypes.interface_purpose, interface.purpose)
            )
            elements.append(
                clone_paragraph_with_text(
                    prototypes.supply_method, "计划供数方式：API接口对外服务"
                )
            )
            elements.append(
                clone_paragraph_with_text(prototypes.update_frequency, "计划更新频率：无")
            )
            elements.append(clone_paragraph_with_text(prototypes.input_label, "接口输入参数："))
            elements.append(
                clone_table_with_data(
                    prototypes.input_parameter_table,
                    NORMALIZED_PARAMETER_HEADERS,
                    normalized_parameter_rows(interface.input_groups),
                )
            )
            elements.append(clone_paragraph_with_text(prototypes.output_label, "接口输出参数："))
            elements.append(
                clone_table_with_data(
                    prototypes.output_parameter_table,
                    NORMALIZED_PARAMETER_HEADERS,
                    normalized_parameter_rows(interface.output_groups),
                )
            )
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
    business_heading = find_heading(document, "Heading 2", "业务场景")
    demand_intro_heading = find_heading(document, "Heading 2", "需求说明")
    demand_heading = find_heading(document, "Heading 2", "需求清单")
    shared_heading = find_heading(document, "Heading 2", "共享开放方案")
    api_heading = find_heading(document, "Heading 2", "API接口清单")
    prototypes = _capture_template_prototypes(
        document,
        business_heading,
        demand_intro_heading,
        demand_heading,
        shared_heading,
        api_heading,
    )

    replace_between(
        business_heading._p,
        demand_intro_heading._p,
        _build_business_scene(orders, prototypes),
    )
    replace_between(
        demand_heading._p,
        shared_heading._p,
        _build_demand_section(orders, prototypes),
    )
    replace_after(api_heading._p, _build_api_section(orders, prototypes))

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    center_table_header_rows(document)
    document.save(output)
    update_toc_via_com(output)
    return str(output)

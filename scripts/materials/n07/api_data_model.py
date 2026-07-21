import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from materials.n07.api_requirement import (
    NORMALIZED_PARAMETER_HEADERS,
    normalized_parameter_rows,
    parse_api_report,
)
from materials.shared.embedded_docx import extract_embedded_docx_by_work_order
from materials.shared.ledger import read_api_work_orders
from materials.shared.word_sections import (
    clone_paragraph_with_text,
    clone_table_with_data,
    element_text,
    elements_between,
    find_heading,
    replace_between,
    update_toc_via_com,
)


@dataclass(frozen=True)
class ApiDataModelPrototypes:
    work_order_heading: object
    interface_heading: object
    input_heading: object
    output_heading: object
    input_parameter_table: object
    output_parameter_table: object


def _paragraph_style_id(element):
    properties = element.find(qn("w:pPr"))
    style = properties.find(qn("w:pStyle")) if properties is not None else None
    return style.get(qn("w:val")) if style is not None else None


def _require_prototype(value, name):
    if value is None:
        raise ValueError(f"模板中未找到{name}格式原型")
    return value


def _first_paragraph_with_style(elements, style_id, text=None, start=0):
    for index in range(start, len(elements)):
        element = elements[index]
        if element.tag != qn("w:p") or _paragraph_style_id(element) != style_id:
            continue
        if text is None or element_text(element) == text:
            return element
    return None


def _first_table_after(elements, paragraph):
    if paragraph is None:
        return None
    start = elements.index(paragraph) + 1
    return next((item for item in elements[start:] if item.tag == qn("w:tbl")), None)


def _capture_template_prototypes(document, section_heading, next_heading):
    elements = elements_between(section_heading._p, next_heading._p)
    heading_3_style = document.styles["Heading 3"].style_id
    heading_4_style = document.styles["Heading 4"].style_id
    heading_5_style = document.styles["Heading 5"].style_id

    work_order_heading = _first_paragraph_with_style(elements, heading_3_style)
    interface_heading = _first_paragraph_with_style(elements, heading_4_style)
    input_heading = _first_paragraph_with_style(elements, heading_5_style, "接口入参配置表")
    output_heading = _first_paragraph_with_style(elements, heading_5_style, "接口出参配置表")

    return ApiDataModelPrototypes(
        work_order_heading=_require_prototype(work_order_heading, "工单标题"),
        interface_heading=_require_prototype(interface_heading, "接口中文名标题"),
        input_heading=_require_prototype(input_heading, "接口入参配置表标题"),
        output_heading=_require_prototype(output_heading, "接口出参配置表标题"),
        input_parameter_table=_require_prototype(
            _first_table_after(elements, input_heading), "接口入参配置表"
        ),
        output_parameter_table=_require_prototype(
            _first_table_after(elements, output_heading), "接口出参配置表"
        ),
    )


def _build_interface_section(orders, prototypes):
    elements = []
    for order in orders:
        elements.append(clone_paragraph_with_text(prototypes.work_order_heading, order.title))
        for interface in order.interfaces:
            elements.append(
                clone_paragraph_with_text(prototypes.interface_heading, interface.chinese_name)
            )
            elements.append(
                clone_paragraph_with_text(prototypes.input_heading, "接口入参配置表")
            )
            elements.append(
                clone_table_with_data(
                    prototypes.input_parameter_table,
                    NORMALIZED_PARAMETER_HEADERS,
                    normalized_parameter_rows(interface.input_groups),
                )
            )
            elements.append(
                clone_paragraph_with_text(prototypes.output_heading, "接口出参配置表")
            )
            elements.append(
                clone_table_with_data(
                    prototypes.output_parameter_table,
                    NORMALIZED_PARAMETER_HEADERS,
                    normalized_parameter_rows(interface.output_groups),
                )
            )
    return elements


def build_api_data_model_document(excel_path, service_dir, template_path, output_path):
    orders = read_api_work_orders(excel_path, service_dir)
    with tempfile.TemporaryDirectory(prefix="document_filler_api_data_model_") as temp_dir:
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
    section_heading = find_heading(document, "Heading 2", "接口出入参配置表")
    next_heading = find_heading(document, "Heading 2", "接口组件配置表")
    prototypes = _capture_template_prototypes(document, section_heading, next_heading)

    replace_between(section_heading._p, next_heading._p, _build_interface_section(orders, prototypes))

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    update_toc_via_com(output)
    return str(output)

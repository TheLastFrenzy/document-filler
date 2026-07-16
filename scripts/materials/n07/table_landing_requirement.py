from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from materials.n07.table_landing import read_table_landing_work_orders
from materials.shared.word_sections import (
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


@dataclass(frozen=True)
class TableLandingRequirementPrototypes:
    business_summary: object
    demand_source_table: object
    work_order_heading: object
    demand_body: object
    task_count: object
    detail_table: object
    supply_method: object
    update_frequency: object
    update_time: object


def _paragraph_style_id(element):
    properties = element.find(qn("w:pPr"))
    style = properties.find(qn("w:pStyle")) if properties is not None else None
    return style.get(qn("w:val")) if style is not None else None


def _require_prototype(value, name):
    if value is None:
        raise ValueError(f"模板中未找到{name}格式原型")
    return value


def _first_nonempty_paragraph(elements, start=0):
    for index in range(start, len(elements)):
        element = elements[index]
        if element.tag == qn("w:p") and element_text(element):
            return element
    return None


def _paragraph_with_prefix(elements, prefix):
    for element in elements:
        if element.tag == qn("w:p") and element_text(element).startswith(prefix):
            return element
    return None


def _first_table(elements, column_count=None):
    for element in elements:
        if element.tag != qn("w:tbl"):
            continue
        if column_count is None:
            return element
        grid = element.find(qn("w:tblGrid"))
        if grid is not None and len(grid.findall(qn("w:gridCol"))) == column_count:
            return element
        first_row = element.find(qn("w:tr"))
        if first_row is not None and len(first_row.findall(qn("w:tc"))) == column_count:
            return element
    return None


def _capture_template_prototypes(document, business_heading, demand_heading):
    business_elements = elements_between(business_heading._p, demand_heading._p)
    demand_source_table = _first_table(business_elements, 5)
    business_summary = _first_nonempty_paragraph(business_elements)

    demand_elements = []
    current = demand_heading._p.getnext()
    while current is not None and current.tag != qn("w:sectPr"):
        demand_elements.append(current)
        current = current.getnext()

    heading_2_style = document.styles["Heading 2"].style_id
    work_order_heading = next(
        (
            item
            for item in demand_elements
            if item.tag == qn("w:p") and _paragraph_style_id(item) == heading_2_style and element_text(item)
        ),
        None,
    )
    start = demand_elements.index(work_order_heading) + 1 if work_order_heading is not None else 0
    after_heading = demand_elements[start:]
    demand_body = _paragraph_with_prefix(after_heading, "需求口径")
    if demand_body is None:
        demand_body = _first_nonempty_paragraph(after_heading)
    task_count = _paragraph_with_prefix(after_heading, "涉及共享任务数量")
    detail_table = _first_table(after_heading, 6)
    supply_method = _paragraph_with_prefix(after_heading, "计划供数方式")
    update_frequency = _paragraph_with_prefix(after_heading, "计划更新频率")
    if update_frequency is None:
        update_frequency = _paragraph_with_prefix(after_heading, "更新频率")
    update_time = _paragraph_with_prefix(after_heading, "更新时间")
    if update_time is None:
        update_time = update_frequency

    return TableLandingRequirementPrototypes(
        business_summary=_require_prototype(business_summary, "业务场景说明段落"),
        demand_source_table=_require_prototype(demand_source_table, "业务场景需求来源表格"),
        work_order_heading=_require_prototype(work_order_heading, "工单标题"),
        demand_body=_require_prototype(demand_body, "需求口径段落"),
        task_count=_require_prototype(task_count, "涉及共享任务数量段落"),
        detail_table=_require_prototype(detail_table, "库表落地明细表格"),
        supply_method=_require_prototype(supply_method, "计划供数方式段落"),
        update_frequency=_require_prototype(update_frequency, "计划更新频率段落"),
        update_time=_require_prototype(update_time, "更新时间段落"),
    )


def _build_business_scene(orders, prototypes):
    request_count = len({order.demand_no for order in orders if order.demand_no})
    total = sum(order.program_count for order in orders)
    summary = (
        f"服务周期内，共有{request_count}张需求单，{len(orders)}张工单涉及{total}个库表落地共享任务。"
        "具体需求单、工单和产出如下表："
    )
    headers = ["序号", "对应需求单编号", "对应工单编号", "工单内容", "涉及共享任务数量（个）"]
    rows = [
        [str(index), order.demand_no, order.work_order_no, order.title, str(order.program_count)]
        for index, order in enumerate(orders, start=1)
    ]
    return [
        clone_paragraph_with_text(prototypes.business_summary, summary),
        clone_table_with_data(
            prototypes.demand_source_table,
            headers,
            rows,
            footer=["总计", "", "", "", str(total)],
        ),
    ]


def _source_table_for_task(order, index):
    if index < len(order.source_tables):
        return order.source_tables[index]
    return order.source_tables[-1] if order.source_tables else ""


def _build_detail_table(order, prototypes):
    headers = ["序号", "落地库", "落地表", "对象用户", "库表来源", "业务场景"]
    rows = [
        [
            str(index + 1),
            task.landing_database,
            task.landing_table,
            order.target_user,
            _source_table_for_task(order, index),
            task.business_scene,
        ]
        for index, task in enumerate(order.tasks)
    ]
    return clone_table_with_data(prototypes.detail_table, headers, rows)


def _build_requirement_sections(orders, prototypes):
    elements = []
    for order in orders:
        elements.append(
            clone_paragraph_with_text(
                prototypes.work_order_heading,
                f"{order.work_order_no}_{order.title}",
            )
        )
        elements.append(
            clone_paragraph_with_text(prototypes.demand_body, f"需求口径：{order.description}")
        )
        elements.append(
            clone_paragraph_with_text(prototypes.task_count, f"涉及共享任务数量：{order.program_count}")
        )
        elements.append(_build_detail_table(order, prototypes))
        elements.append(
            clone_paragraph_with_text(prototypes.supply_method, "计划供数方式：库表下发")
        )
        elements.append(
            clone_paragraph_with_text(prototypes.update_frequency, f"计划更新频率：{order.update_cycle}")
        )
        elements.append(
            clone_paragraph_with_text(prototypes.update_time, f"更新时间：{order.update_requirement}")
        )
    return elements


def build_table_landing_requirement_document(excel_path, service_dir, template_path, output_path):
    orders = read_table_landing_work_orders(excel_path, service_dir)
    document = Document(template_path)
    business_heading = find_heading(document, "Heading 2", "业务场景")
    demand_heading = find_heading(document, "Heading 2", "需求说明")
    prototypes = _capture_template_prototypes(document, business_heading, demand_heading)

    replace_between(
        business_heading._p,
        demand_heading._p,
        _build_business_scene(orders, prototypes),
    )
    replace_after(demand_heading._p, _build_requirement_sections(orders, prototypes))

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(output)
    update_toc_via_com(output)
    return str(output)

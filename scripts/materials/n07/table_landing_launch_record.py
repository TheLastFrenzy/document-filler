import tempfile
from dataclasses import dataclass
from pathlib import Path

from docx import Document
from docx.oxml.ns import qn

from materials.n07.api_launch_record import _save_document
from materials.n07.api_test_report import _convert_legacy_doc_template
from materials.n07.table_landing import read_table_landing_work_orders
from materials.shared.word_sections import (
    clone_element,
    clone_table_with_data,
    elements_between,
    find_heading,
    replace_between,
)


@dataclass(frozen=True)
class TableLandingLaunchRecordPrototypes:
    list_table: object
    list_elements: tuple[object, ...]


def _require_prototype(value, name):
    if value is None:
        raise ValueError(f"模板中未找到{name}格式原型")
    return value


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


def _capture_template_prototypes(list_heading, record_heading):
    list_elements = tuple(elements_between(list_heading._p, record_heading._p))
    return TableLandingLaunchRecordPrototypes(
        list_table=_require_prototype(_first_table(list_elements, 5), "库表落地上线清单表格"),
        list_elements=list_elements,
    )


def _source_table_for_task(order, index):
    if index < len(order.source_tables):
        return order.source_tables[index]
    return order.source_tables[-1] if order.source_tables else ""


def _all_order_tasks(orders):
    for order in orders:
        for index, task in enumerate(order.tasks):
            yield order, index, task


def _build_launch_table(orders, prototypes):
    rows = [
        [
            str(index),
            _source_table_for_task(order, task_index),
            task.launch_time,
            order.update_cycle,
            task.business_scene,
        ]
        for index, (order, task_index, task) in enumerate(_all_order_tasks(orders), start=1)
    ]
    return clone_table_with_data(
        prototypes.list_table,
        ["序号", "调度名", "上线时间", "更新频率", "场景名称"],
        rows,
    )


def _build_list_elements(orders, prototypes):
    elements = []
    replaced = False
    for element in prototypes.list_elements:
        if not replaced and element is prototypes.list_table:
            elements.append(_build_launch_table(orders, prototypes))
            replaced = True
            continue
        elements.append(clone_element(element))
    if not replaced:
        elements.append(_build_launch_table(orders, prototypes))
    return elements


def build_table_landing_launch_record_document(excel_path, service_dir, template_path, output_path):
    orders = read_table_landing_work_orders(excel_path, service_dir)
    with tempfile.TemporaryDirectory(prefix="document_filler_table_landing_launch_record_") as temp_dir:
        converted_template = _convert_legacy_doc_template(template_path, Path(temp_dir) / "template")
        document = Document(converted_template)
        list_heading = find_heading(document, "Heading 1", "库表落地上线清单")
        record_heading = find_heading(document, "Heading 1", "库表落地上线记录")
        prototypes = _capture_template_prototypes(list_heading, record_heading)

        replace_between(list_heading._p, record_heading._p, _build_list_elements(orders, prototypes))

        return _save_document(document, output_path, Path(temp_dir) / "output")

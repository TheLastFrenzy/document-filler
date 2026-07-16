from dataclasses import dataclass, field
from pathlib import Path
import re
import tempfile
import zipfile

import openpyxl

from materials.shared.embedded_docx import (
    package_stream_from_ole,
    read_ole_anchors,
)
from materials.shared.ledger import merged_value_getter


REQUIRED_TABLE_LANDING_HEADERS = (
    "服务目录",
    "需求单号",
    "工单号",
    "工单标题",
    "程序数",
    "工单描述",
    "结果表清单",
    "数据统计分析执行周期",
    "数据更新要求",
    "自测报告附件",
    "下发前置机中文名",
)

REQUIRED_TASK_HEADERS = (
    "落地库名",
    "落地表名",
    "表中文名称",
)


@dataclass
class TableLandingTask:
    landing_database: str
    landing_table: str
    business_scene: str
    landing_data_count: str = ""
    dispatch_description: str = ""
    target_volume_image: bytes | None = None


@dataclass
class TableLandingWorkOrder:
    demand_no: str
    work_order_no: str
    title: str
    description: str
    program_count: int
    source_rows: tuple[int, ...]
    source_tables: list[str]
    target_user: str
    database_type: str
    update_cycle: str
    update_requirement: str
    attachment_path: Path | None = None
    tasks: list[TableLandingTask] = field(default_factory=list)


def _parse_program_count(value):
    try:
        return int(float(str(value).strip()))
    except (TypeError, ValueError):
        raise ValueError(f"程序数不是有效整数: {value}") from None


def _read_order_rows(excel_path, service_dir):
    workbook = openpyxl.load_workbook(excel_path, data_only=True)
    sheet = workbook.active
    headers = [str(sheet.cell(1, column).value or "").strip() for column in range(1, sheet.max_column + 1)]
    missing = [header for header in REQUIRED_TABLE_LANDING_HEADERS if header not in headers]
    if missing:
        workbook.close()
        raise ValueError(f"台账缺少必要列: {', '.join(missing)}")

    columns = {header: headers.index(header) + 1 for header in REQUIRED_TABLE_LANDING_HEADERS}
    database_type_column = headers.index("下发前置机数据库类型") + 1 if "下发前置机数据库类型" in headers else None
    get = merged_value_getter(sheet)
    grouped = {}
    order_keys = []
    for row in range(2, sheet.max_row + 1):
        if get(row, columns["服务目录"]) != service_dir:
            continue
        work_order_no = get(row, columns["工单号"])
        if not work_order_no:
            workbook.close()
            raise ValueError(f"第{row}行缺少工单号")
        if work_order_no not in grouped:
            grouped[work_order_no] = {
                "demand_no": get(row, columns["需求单号"]),
                "title": get(row, columns["工单标题"]),
                "description": get(row, columns["工单描述"]),
                "program_count": _parse_program_count(get(row, columns["程序数"])),
                "source_rows": [],
                "source_tables": [],
                "target_user": get(row, columns["下发前置机中文名"]),
                "database_type": get(row, database_type_column) if database_type_column else "",
                "update_cycle": get(row, columns["数据统计分析执行周期"]),
                "update_requirement": get(row, columns["数据更新要求"]),
            }
            order_keys.append(work_order_no)
        group = grouped[work_order_no]
        group["source_rows"].append(row)
        source_table = get(row, columns["结果表清单"])
        if source_table:
            group["source_tables"].append(source_table)
    workbook.close()

    if not grouped:
        raise ValueError(f"未找到匹配数据: 服务目录={service_dir}")

    return [
        TableLandingWorkOrder(
            demand_no=grouped[key]["demand_no"],
            work_order_no=key,
            title=grouped[key]["title"],
            description=grouped[key]["description"],
            program_count=grouped[key]["program_count"],
            source_rows=tuple(grouped[key]["source_rows"]),
            source_tables=grouped[key]["source_tables"],
            target_user=grouped[key]["target_user"],
            database_type=grouped[key]["database_type"],
            update_cycle=grouped[key]["update_cycle"],
            update_requirement=grouped[key]["update_requirement"],
        )
        for key in order_keys
    ]


def _header_column(excel_path, header):
    workbook = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheet = workbook.active
    headers = [str(sheet.cell(1, column).value or "").strip() for column in range(1, sheet.max_column + 1)]
    workbook.close()
    if header not in headers:
        raise ValueError(f"台账缺少必要列: {header}")
    return headers.index(header) + 1


def _workbook_payload(data):
    if data.startswith(b"PK\x03\x04"):
        return data
    return package_stream_from_ole(data)


def extract_embedded_workbooks_by_work_order(excel_path, work_orders, attachment_header, work_dir):
    attachment_column = _header_column(excel_path, attachment_header)
    output_dir = Path(work_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    workbooks = {}
    for anchor in read_ole_anchors(excel_path):
        if anchor.column != attachment_column:
            continue
        order = next((item for item in work_orders if anchor.row in item.source_rows), None)
        if order is None:
            continue
        try:
            payload = _workbook_payload(anchor.payload)
        except Exception as exc:
            raise ValueError(f"工单{order.work_order_no}的自测报告附件无法解析为Excel: {exc}") from exc
        path = output_dir / f"{order.work_order_no}_self_report.xlsx"
        path.write_bytes(payload)
        workbooks[order.work_order_no] = path

    missing = [item.work_order_no for item in work_orders if item.work_order_no not in workbooks]
    if missing:
        raise ValueError(f"以下工单缺少可解析的自测报告附件: {', '.join(missing)}")
    return workbooks


def _extract_images_via_cellimages(excel_path):
    result = {}
    with zipfile.ZipFile(excel_path, "r") as archive:
        if "xl/cellimages.xml" not in archive.namelist():
            return result
        cell_images = archive.read("xl/cellimages.xml").decode("utf-8")
        name_to_rid = {}
        for block in re.findall(r"<etc:cellImage>(.*?)</etc:cellImage>", cell_images, re.DOTALL):
            name_match = re.search(r'name="([^"]+)"', block)
            rid_match = re.search(r'r:embed="(rId\d+)"', block)
            if name_match and rid_match:
                name_to_rid[name_match.group(1)] = rid_match.group(1)
        relationships_path = "xl/_rels/cellimages.xml.rels"
        if relationships_path not in archive.namelist():
            return result
        relationships = archive.read(relationships_path).decode("utf-8")
        rid_to_file = {}
        for match in re.finditer(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', relationships):
            target = match.group(2)
            if target.startswith("../"):
                target = "xl/" + target[3:]
            elif target.startswith("xl/"):
                target = target
            else:
                target = "xl/" + target.lstrip("/")
            rid_to_file[match.group(1)] = target
        for name, rid in name_to_rid.items():
            target = rid_to_file.get(rid)
            if target:
                result[name] = archive.read(target)
    return result


def _target_volume_images_by_row(workbook_path):
    images_by_name = _extract_images_via_cellimages(workbook_path)
    if not images_by_name:
        return {}
    workbook = openpyxl.load_workbook(workbook_path, read_only=True, data_only=False)
    if len(workbook.worksheets) < 2:
        workbook.close()
        return {}
    sheet = workbook.worksheets[1]
    headers = [str(sheet.cell(1, column).value or "").strip() for column in range(1, sheet.max_column + 1)]
    if "目标表数据量" not in headers:
        workbook.close()
        return {}
    column = headers.index("目标表数据量") + 1
    matched = {}
    for row in range(2, sheet.max_row + 1):
        value = str(sheet.cell(row, column).value or "")
        match = re.search(r'DISPIMG\("([^"]+)"', value)
        if match and match.group(1) in images_by_name:
            matched[row] = images_by_name[match.group(1)]
    workbook.close()
    return matched


def parse_landing_tasks(workbook_path):
    workbook = openpyxl.load_workbook(workbook_path, data_only=True, read_only=True)
    sheet = workbook.worksheets[0]
    headers = [str(sheet.cell(1, column).value or "").strip() for column in range(1, sheet.max_column + 1)]
    missing = [header for header in REQUIRED_TASK_HEADERS if header not in headers]
    if missing:
        workbook.close()
        raise ValueError(f"附件缺少必要列: {', '.join(missing)}")
    columns = {header: headers.index(header) + 1 for header in REQUIRED_TASK_HEADERS}
    landing_data_column = headers.index("落地数据量") + 1 if "落地数据量" in headers else None
    dispatch_column = (
        headers.index("下发任务简单说明（50字以内）") + 1
        if "下发任务简单说明（50字以内）" in headers
        else None
    )
    target_images = _target_volume_images_by_row(workbook_path)

    tasks = []
    for row in range(2, sheet.max_row + 1):
        landing_database = str(sheet.cell(row, columns["落地库名"]).value or "").strip()
        landing_table = str(sheet.cell(row, columns["落地表名"]).value or "").strip()
        business_scene = str(sheet.cell(row, columns["表中文名称"]).value or "").strip()
        landing_data_count = str(sheet.cell(row, landing_data_column).value or "").strip() if landing_data_column else ""
        dispatch_description = str(sheet.cell(row, dispatch_column).value or "").strip() if dispatch_column else ""
        if not any((landing_database, landing_table, business_scene)):
            continue
        tasks.append(
            TableLandingTask(
                landing_database=landing_database,
                landing_table=landing_table,
                business_scene=business_scene,
                landing_data_count=landing_data_count,
                dispatch_description=dispatch_description,
                target_volume_image=target_images.get(row),
            )
        )
    workbook.close()
    if not tasks:
        raise ValueError(f"附件未解析到库表落地明细: {workbook_path}")
    return tasks


def read_table_landing_work_orders(excel_path, service_dir):
    orders = _read_order_rows(excel_path, service_dir)
    with tempfile.TemporaryDirectory(prefix="document_filler_table_landing_") as temp_dir:
        workbooks = extract_embedded_workbooks_by_work_order(
            excel_path,
            orders,
            "自测报告附件",
            Path(temp_dir) / "workbooks",
        )
        for order in orders:
            order.attachment_path = workbooks[order.work_order_no]
            order.tasks = parse_landing_tasks(order.attachment_path)
    return orders

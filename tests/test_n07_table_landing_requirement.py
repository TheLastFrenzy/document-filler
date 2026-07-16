import importlib.util
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

import openpyxl
from docx import Document


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
FILL_SCRIPT = SCRIPTS / "fill_document.py"


def load_fill_document_module():
    spec = importlib.util.spec_from_file_location("fill_document_table_landing_test", FILL_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def make_attachment_bytes(rows):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "上线记录库表落地"
    sheet.append(["工单号", "落地库名", "落地表名", "表中文名称"])
    for row in rows:
        sheet.append(row)
    with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as temp_file:
        path = Path(temp_file.name)
    try:
        workbook.save(path)
        return path.read_bytes()
    finally:
        path.unlink(missing_ok=True)


def inject_workbook_embeddings(path, embeddings):
    with zipfile.ZipFile(path, "r") as source:
        entries = {name: source.read(name) for name in source.namelist()}

    sheet_xml = entries["xl/worksheets/sheet1.xml"].decode("utf-8")
    if "xmlns:r=" not in sheet_xml:
        sheet_xml = sheet_xml.replace(
            "<worksheet ",
            '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
            1,
        )
    ole_objects = (
        '<oleObjects>'
        '<oleObject progId="Excel.Sheet.12" shapeId="1025" r:id="rId2"/>'
        '<oleObject progId="Excel.Sheet.12" shapeId="1026" r:id="rId3"/>'
        "</oleObjects>"
        '<legacyDrawing r:id="rId1"/>'
    )
    sheet_xml = sheet_xml.replace("</worksheet>", ole_objects + "</worksheet>")
    entries["xl/worksheets/sheet1.xml"] = sheet_xml.encode("utf-8")
    entries["xl/worksheets/_rels/sheet1.xml.rels"] = b"""<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/Workbook1.xlsx"/>
 <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/Workbook2.xlsx"/>
</Relationships>"""
    entries["xl/drawings/vmlDrawing1.vml"] = b"""<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <v:shape id="_x0000_s1025" o:ole="t">
  <x:ClientData><x:Anchor>9, 0, 1, 0, 9, 0, 1, 0</x:Anchor></x:ClientData>
 </v:shape>
 <v:shape id="_x0000_s1026" o:ole="t">
  <x:ClientData><x:Anchor>9, 0, 2, 0, 9, 0, 3, 0</x:Anchor></x:ClientData>
 </v:shape>
</xml>"""
    entries["xl/embeddings/Workbook1.xlsx"] = embeddings[0]
    entries["xl/embeddings/Workbook2.xlsx"] = embeddings[1]

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as target:
        for name, data in entries.items():
            target.writestr(name, data)


def make_table_landing_ledger(path):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.append(
        [
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
        ]
    )
    sheet.append(
        [
            "N07-库表落地方式",
            "REQ-1",
            "WO-1",
            "工单一",
            1,
            "工单一描述",
            "PRE_SRC_ONE",
            "每日",
            "每日凌晨",
            None,
            "前置机A",
        ]
    )
    sheet.append(
        [
            "N07-库表落地方式",
            "REQ-2",
            "WO-2",
            "工单二",
            2,
            "工单二描述",
            "PRE_SRC_TWO_A",
            "每月",
            "每月15日",
            None,
            "前置机B",
        ]
    )
    sheet.append([None, None, None, None, None, None, "PRE_SRC_TWO_B", None, None, None, None])
    for column in "ABCDEFHIJK":
        sheet.merge_cells(f"{column}3:{column}4")
    workbook.save(path)

    inject_workbook_embeddings(
        path,
        [
            make_attachment_bytes([["WO-1", "DB_ONE", "TABLE_ONE", "业务一"]]),
            make_attachment_bytes(
                [
                    ["WO-2", "DB_TWO", "TABLE_TWO", "业务二"],
                    ["WO-2", "DB_THREE", "TABLE_THREE", "业务三"],
                ]
            ),
        ],
    )
    return path


def make_requirement_template(path):
    document = Document()
    document.add_heading("需求文档", level=1)
    document.add_heading("业务场景", level=2)
    document.add_paragraph("上海市大数据中心通过库表落地方式开展公共数据共享。")
    document.add_paragraph("旧需求来源说明。")
    source_table = document.add_table(rows=3, cols=5)
    for index, value in enumerate(["序号", "对应需求单编号", "对应工单编号", "工单内容", "涉及共享任务数量（个）"]):
        source_table.rows[0].cells[index].text = value
    for index, value in enumerate(["1", "REQ-OLD", "WO-OLD", "旧工单", "1"]):
        source_table.rows[1].cells[index].text = value
    source_table.rows[2].cells[0].text = "总计"
    source_table.rows[2].cells[4].text = "1"
    document.add_heading("需求说明", level=2)
    document.add_heading("WO-OLD_旧工单", level=2)
    document.add_paragraph("需求口径：旧内容")
    document.add_paragraph("涉及共享任务数量：1")
    detail_table = document.add_table(rows=2, cols=6)
    for index, value in enumerate(["序号", "落地库", "落地表", "对象用户", "库表来源", "业务场景"]):
        detail_table.rows[0].cells[index].text = value
    for index, value in enumerate(["1", "OLD_DB", "OLD_TABLE", "旧前置机", "OLD_SRC", "旧场景"]):
        detail_table.rows[1].cells[index].text = value
    document.add_paragraph("计划供数方式：库表下发")
    document.add_paragraph("计划更新频率：旧频率")
    document.add_paragraph("更新时间：旧时间")
    document.save(path)
    return path


class N07TableLandingRequirementTest(unittest.TestCase):
    def test_registration_uses_public_requirement_filename(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            output = module.resolve_output_path(temp_dir, "01-库表落地方式_需求文档")

        self.assertEqual(Path(output).name, "01-需求文档.docx")

    def test_read_table_landing_work_orders_groups_merged_rows_and_reads_workbook_tasks(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.n07.table_landing import read_table_landing_work_orders

        with tempfile.TemporaryDirectory() as temp_dir:
            ledger = make_table_landing_ledger(Path(temp_dir) / "ledger.xlsx")
            orders = read_table_landing_work_orders(ledger, "N07-库表落地方式")

        self.assertEqual([order.work_order_no for order in orders], ["WO-1", "WO-2"])
        self.assertEqual(orders[0].source_rows, (2,))
        self.assertEqual(orders[1].source_rows, (3, 4))
        self.assertEqual(orders[1].program_count, 2)
        self.assertEqual(orders[1].source_tables, ["PRE_SRC_TWO_A", "PRE_SRC_TWO_B"])
        self.assertEqual(
            [(task.landing_database, task.landing_table, task.business_scene) for task in orders[1].tasks],
            [("DB_TWO", "TABLE_TWO", "业务二"), ("DB_THREE", "TABLE_THREE", "业务三")],
        )

    def test_build_table_landing_requirement_document_replaces_business_and_detail_sections(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.n07.table_landing_requirement import build_table_landing_requirement_document

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = make_table_landing_ledger(temp / "ledger.xlsx")
            template = make_requirement_template(temp / "template.docx")
            output = temp / "out.docx"

            result = build_table_landing_requirement_document(
                excel_path=str(ledger),
                service_dir="N07-库表落地方式",
                template_path=str(template),
                output_path=str(output),
            )

            document = Document(result)

        paragraphs = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
        self.assertIn("WO-1_工单一", paragraphs)
        self.assertIn("WO-2_工单二", paragraphs)
        self.assertIn("需求口径：工单二描述", paragraphs)
        self.assertIn("计划更新频率：每月", paragraphs)
        self.assertIn("更新时间：每月15日", paragraphs)
        self.assertNotIn("WO-OLD_旧工单", paragraphs)

        source_rows = [[cell.text.strip() for cell in row.cells] for row in document.tables[0].rows]
        self.assertEqual(source_rows[1], ["1", "REQ-1", "WO-1", "工单一", "1"])
        self.assertEqual(source_rows[2], ["2", "REQ-2", "WO-2", "工单二", "2"])
        self.assertEqual(source_rows[-1][-1], "3")

        first_detail = [[cell.text.strip() for cell in row.cells] for row in document.tables[1].rows]
        second_detail = [[cell.text.strip() for cell in row.cells] for row in document.tables[2].rows]
        self.assertEqual(first_detail[1], ["1", "DB_ONE", "TABLE_ONE", "前置机A", "PRE_SRC_ONE", "业务一"])
        self.assertEqual(second_detail[1], ["1", "DB_TWO", "TABLE_TWO", "前置机B", "PRE_SRC_TWO_A", "业务二"])
        self.assertEqual(second_detail[2], ["2", "DB_THREE", "TABLE_THREE", "前置机B", "PRE_SRC_TWO_B", "业务三"])


if __name__ == "__main__":
    unittest.main()

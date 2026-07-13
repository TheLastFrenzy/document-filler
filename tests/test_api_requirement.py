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
    spec = importlib.util.spec_from_file_location("fill_document_api_test", FILL_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def make_api_ledger(path: Path):
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
            "自测报告附件",
        ]
    )
    sheet.append(
        [
            "N07-API接口开发",
            "REQ-1",
            "WO-1",
            "市公安局-出入境证件身份认证-共享接口",
            3,
            "升级离境退税掌上办平台的出入境记录校验能力。",
            "境外人员获取令牌接口 token_api",
            None,
        ]
    )
    sheet.append([None, None, None, None, None, None, "境外人员身份认证申请接口 apply_api", None])
    sheet.append([None, None, None, None, None, None, "境外人员身份认证请求接口 auth_api", None])
    for column in "ABCDEFH":
        sheet.merge_cells(f"{column}2:{column}4")
    workbook.save(path)
    return path


def add_parameter_table(document, headers, rows):
    table = document.add_table(rows=1, cols=len(headers))
    for index, header in enumerate(headers):
        table.rows[0].cells[index].text = header
    for values in rows:
        cells = table.add_row().cells
        for index, value in enumerate(values):
            cells[index].text = value


def make_self_report(path: Path):
    document = Document()
    document.add_heading("测试目的", level=1)
    document.add_paragraph("验证接口输入输出参数。")
    document.add_heading("测试内容", level=1)
    document.add_heading("身份认证申请接口", level=2)
    document.add_paragraph("共享接口命名规范性")
    document.add_paragraph("输入参数")
    document.add_paragraph("请求头：")
    add_parameter_table(
        document,
        ["参数", "参数说明", "是否必选"],
        [["access_token", "访问令牌", "是"]],
    )
    document.add_paragraph("请求体：")
    add_parameter_table(
        document,
        ["参数", "参数说明", "是否必选", "类型"],
        [["custNum", "业务站点号", "是", "String"]],
    )
    document.add_paragraph("测试结果：符合规范要求，测试通过。")
    document.add_paragraph("输出参数")
    add_parameter_table(
        document,
        ["参数", "参数说明", "备注"],
        [["bizSerialNum", "业务流水号", "申请成功后返回"]],
    )
    document.add_paragraph("测试结果：符合规范要求，测试通过。")
    document.add_paragraph("共享任务开发代码检查")
    document.save(path)
    return path


def make_ole_anchor_package(path: Path):
    worksheet = """<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <oleObjects><oleObject shapeId="1025" r:id="rId2"/></oleObjects>
 <legacyDrawing r:id="rId1"/>
</worksheet>"""
    relationships = """<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
 <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/>
 <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" Target="../embeddings/oleObject1.bin"/>
</Relationships>"""
    vml = """<?xml version="1.0" encoding="UTF-8"?>
<xml xmlns:v="urn:schemas-microsoft-com:vml"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel">
 <v:shape id="_x0000_s1025" o:ole="t">
  <x:ClientData><x:Anchor>7, 0, 1, 0, 7, 0, 3, 0</x:Anchor></x:ClientData>
 </v:shape>
</xml>"""
    with zipfile.ZipFile(path, "w") as archive:
        archive.writestr("xl/worksheets/sheet1.xml", worksheet)
        archive.writestr("xl/worksheets/_rels/sheet1.xml.rels", relationships)
        archive.writestr("xl/drawings/vmlDrawing1.vml", vml)
        archive.writestr("xl/embeddings/oleObject1.bin", b"embedded-ole-payload")
    return path


class ApiRequirementTest(unittest.TestCase):
    def test_api_requirement_registration_uses_public_output_filename(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            output = module.resolve_output_path(temp_dir, "01-API接口开发_需求文档")

        self.assertEqual(Path(output).name, "01-需求文档.docx")

    def test_read_api_work_orders_groups_merged_rows_and_counts_programs_once(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.shared.ledger import read_api_work_orders

        with tempfile.TemporaryDirectory() as temp_dir:
            ledger = make_api_ledger(Path(temp_dir) / "ledger.xlsx")
            orders = read_api_work_orders(ledger, "N07-API接口开发")

        self.assertEqual(len(orders), 1)
        self.assertEqual(orders[0].program_count, 3)
        self.assertEqual(orders[0].source_rows, (2, 3, 4))
        self.assertEqual(
            [item.chinese_name for item in orders[0].interfaces],
            ["境外人员获取令牌接口", "境外人员身份认证申请接口", "境外人员身份认证请求接口"],
        )

    def test_parse_self_report_preserves_input_groups_and_output_tables(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.n07.api_requirement import parse_api_report
        from materials.shared.ledger import ApiInterface

        with tempfile.TemporaryDirectory() as temp_dir:
            report = make_self_report(Path(temp_dir) / "self-test.docx")
            interfaces = [ApiInterface("身份认证申请接口", "api_apply", 2)]
            parsed = parse_api_report(report, interfaces)

        item = parsed[0]
        self.assertEqual([group.label for group in item.input_groups], ["请求头", "请求体"])
        self.assertEqual(item.input_groups[0].headers, ["参数", "参数说明", "是否必选"])
        self.assertEqual(item.input_groups[1].rows[0], ["custNum", "业务站点号", "是", "String"])
        self.assertEqual(item.output_groups[0].rows[0][0], "bizSerialNum")

    def test_read_ole_anchors_resolves_vml_position_and_embedding_payload(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.shared.embedded_docx import read_ole_anchors

        with tempfile.TemporaryDirectory() as temp_dir:
            package = make_ole_anchor_package(Path(temp_dir) / "ledger.xlsx")
            anchors = read_ole_anchors(package)

        self.assertEqual(len(anchors), 1)
        self.assertEqual((anchors[0].row, anchors[0].column), (2, 8))
        self.assertEqual(anchors[0].payload, b"embedded-ole-payload")


if __name__ == "__main__":
    unittest.main()

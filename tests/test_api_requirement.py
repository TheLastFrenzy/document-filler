import importlib.util
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest import mock

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

    document.add_heading("境外人员获取令牌接口", level=2)
    document.add_paragraph("输入参数")
    add_parameter_table(
        document,
        ["序号", "参数项", "名称"],
        [["1", "userId", "用户ID"], ["2", "password", "密码"]],
    )
    document.add_paragraph("输出参数")
    add_parameter_table(
        document,
        ["参数", "参数说明", "备注"],
        [["access_token", "Token", "后续接口调用凭证"]],
    )
    document.add_paragraph("测试结果：符合规范要求，测试通过。")

    document.add_heading("境外人员身份认证申请接口", level=2)
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

    document.add_heading("境外人员身份认证请求接口", level=2)
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
        [["bizSerialNum", "业务流水号", "是", "String"], ["idAuthData", "身份验证数据", "是", "Object"]],
    )
    document.add_paragraph("输出参数")
    add_parameter_table(
        document,
        ["参数", "参数说明", "备注"],
        [["success", "是否成功", ""], ["result", "核验结果", "四位结果数组"]],
    )
    document.add_paragraph("测试结果：符合规范要求，测试通过。")
    document.save(path)
    return path


def make_api_template(path: Path):
    document = Document()
    document.add_heading("需求文档", level=1)
    revision = document.add_table(rows=2, cols=4)
    for index, value in enumerate(["版本", "更新人员", "更新内容", "时间"]):
        revision.rows[0].cells[index].text = value
    document.add_heading("业务场景", level=2)
    document.add_paragraph("上海市大数据中心通过API接口开展公共数据共享开放。")
    document.add_paragraph("旧业务场景内容。")
    document.add_heading("需求说明", level=2)
    document.add_heading("需求清单", level=2)
    document.add_paragraph("服务周期内，共有3张需求单。")
    demand = document.add_table(rows=2, cols=5)
    for index, value in enumerate(["序号", "对应需求单编号", "对应工单编号", "工单内容", "涉及共享任务数量（个）"]):
        demand.rows[0].cells[index].text = value
    document.add_heading("OLD-WO_旧工单", level=3)
    document.add_paragraph("需求口径：旧内容。")
    document.add_heading("共享开放方案", level=2)
    document.add_paragraph("STATIC_SHARED_SOLUTION_MARKER")
    document.add_heading("API接口清单", level=2)
    document.add_heading("企业电子票据证件号码校验接口", level=3)
    document.add_paragraph("旧接口说明。")
    document.add_paragraph("计划供数方式：API接口对外服务")
    document.add_paragraph("计划更新频率：无")
    document.add_paragraph("接口输入参数：")
    add_parameter_table(document, ["序号", "参数名", "名称"], [["1", "old", "旧参数"]])
    document.add_paragraph("接口输出参数：")
    add_parameter_table(document, ["序号", "参数项", "名称"], [["1", "oldResult", "旧结果"]])
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

    def test_api_requirement_nonexistent_output_directory_uses_public_filename(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            result_dir = Path(temp_dir) / "结果"
            output = module.resolve_output_path(result_dir, "01-API接口开发_需求文档")

        self.assertEqual(Path(output), result_dir / "01-需求文档.docx")

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
            interfaces = [ApiInterface("境外人员身份认证申请接口", "api_apply", 2)]
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

    def test_build_api_requirement_document_replaces_dynamic_sections_and_preserves_static_content(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.n07 import api_requirement

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = make_api_ledger(temp / "ledger.xlsx")
            report = make_self_report(temp / "self-test.docx")
            template = make_api_template(temp / "template.docx")
            output = temp / "01-需求文档.docx"
            with mock.patch.object(
                api_requirement,
                "extract_embedded_docx_by_work_order",
                return_value={"WO-1": report},
                create=True,
            ):
                with mock.patch.object(api_requirement, "update_toc_via_com", create=True):
                    api_requirement.build_api_requirement_document(
                        excel_path=ledger,
                        service_dir="N07-API接口开发",
                        template_path=template,
                        output_path=output,
                    )

            document = Document(output)
            text = "\n".join(paragraph.text for paragraph in document.paragraphs)

        self.assertIn("共提供3个API接口服务", text)
        self.assertIn("WO-1_市公安局-出入境证件身份认证-共享接口", text)
        self.assertIn("需求口径：升级离境退税掌上办平台的出入境记录校验能力。", text)
        self.assertIn("STATIC_SHARED_SOLUTION_MARKER", text)
        self.assertIn("境外人员获取令牌接口", text)
        self.assertIn("境外人员身份认证申请接口", text)
        self.assertIn("境外人员身份认证请求接口", text)
        self.assertIn("计划供数方式：API接口对外服务", text)
        self.assertNotIn("企业电子票据证件号码校验接口", text)
        self.assertNotIn("旧接口说明", text)

    def test_purpose_text_is_grounded_and_avoids_banned_ai_phrases(self):
        if str(SCRIPTS) not in sys.path:
            sys.path.insert(0, str(SCRIPTS))
        from materials.n07.api_requirement import build_interface_purpose
        from materials.shared.ledger import ApiInterface, ApiWorkOrder, ParameterGroup

        interface = ApiInterface(
            chinese_name="境外人员身份认证申请接口",
            english_name="apply_api",
            source_row=2,
            input_groups=[ParameterGroup("请求体", ["参数", "参数说明"], [["custNum", "业务站点号"]])],
            output_groups=[ParameterGroup("", ["参数", "参数说明"], [["bizSerialNum", "业务流水号"]])],
        )
        order = ApiWorkOrder(
            demand_no="REQ-1",
            work_order_no="WO-1",
            title="出入境证件身份认证",
            description="升级离境退税掌上办平台的出入境记录校验能力。",
            program_count=1,
            source_rows=(2,),
            interfaces=[interface],
        )

        purpose = build_interface_purpose(order, interface)

        self.assertIn("身份认证", purpose)
        self.assertIn("业务流水号", purpose)
        for banned in ("赋能", "至关重要", "确保", "重要支撑"):
            self.assertNotIn(banned, purpose)

    def test_fill_document_dispatches_registered_api_requirement_material(self):
        module = load_fill_document_module()
        builder = mock.Mock(return_value="out.docx")
        spec = mock.Mock(default_filename="01-需求文档.docx")
        with mock.patch.object(module, "get_material_spec", return_value=spec):
            with mock.patch.object(module, "load_material_builder", return_value=builder, create=True):
                result = module.fill_document(
                    excel_path="ledger.xlsx",
                    service_dir="N07-API接口开发",
                    material_type="01-API接口开发_需求文档",
                    template_path="template.docx",
                    output_path="out.docx",
                )

        self.assertEqual(result, "out.docx")
        builder.assert_called_once_with(
            excel_path="ledger.xlsx",
            service_dir="N07-API接口开发",
            template_path="template.docx",
            output_path="out.docx",
        )


if __name__ == "__main__":
    unittest.main()

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
from PIL import Image


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "fill_document.py"
BUILDER_SCRIPT = ROOT / "scripts" / "build_stats_result_usage_workbook.py"


def load_fill_document_module():
    spec = importlib.util.spec_from_file_location("fill_document", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def load_builder_module():
    spec = importlib.util.spec_from_file_location("build_stats_result_usage_workbook", BUILDER_SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def make_ledger(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["服务目录", "程序XML文本", "统计分析结果表清单"])
    ws.append(["N08-数据统计分析", "<mxGraphModel />", "测试结果表 TEST_RESULT"])
    wb.save(path)


def make_new_column_ledger(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["服务目录", "程序XML文本", "结果表清单"])
    ws.append(["N08-数据统计分析", "<mxGraphModel />", "测试结果表 TEST_RESULT"])
    wb.save(path)


def make_result_template(path: Path):
    wb = openpyxl.Workbook()
    wb.active.title = "说明"
    wb.create_sheet("1、数据源表list")
    wb.create_sheet("2、表融合关系")
    wb.create_sheet("3、数据统计分析结果表list")
    wb.create_sheet("4、数据统计分析结果表详情")
    wb.save(path)


def make_records_with_flowcharts(temp: Path):
    flowcharts = []
    for index in range(1, 3):
        image_path = temp / f"flowchart-{index}.png"
        Image.new("RGB", (16, 16), "white").save(image_path)
        flowcharts.append(image_path)
    return [
        {
            "result_cn": "结果表一",
            "result_en": "RESULT_ONE",
            "sources": ["SRC_ONE"],
            "logic_steps": ["读取源表一", "写入结果表一"],
            "flowchart": str(flowcharts[0]),
            "fields": [
                {
                    "字段中文名": "字段一",
                    "字段英文名": "FIELD_ONE",
                    "字段类型": "VARCHAR2(20)",
                    "默认": "",
                    "不可为空": "No",
                    "唯一": "No",
                    "字段注释": "字段说明",
                }
            ],
        },
        {
            "result_cn": "结果表二",
            "result_en": "RESULT_TWO",
            "sources": ["SRC_TWO"],
            "logic_steps": ["读取源表二", "写入结果表二"],
            "flowchart": str(flowcharts[1]),
            "fields": [
                {
                    "字段中文名": "字段二",
                    "字段英文名": "FIELD_TWO",
                    "字段类型": "NUMBER(10)",
                    "默认": "",
                    "不可为空": "No",
                    "唯一": "No",
                    "字段注释": "字段说明",
                }
            ],
        },
    ]


class StatsResultDispatchTest(unittest.TestCase):
    def test_business_logic_steps_avoid_unsuitable_xml_logic_terms(self):
        module = load_builder_module()
        record = {
            "result_cn": "热线工单统计结果表",
            "result_en": "FUSION_HOTLINE_ORDER",
            "sources": ["SRC_SECRET"],
            "nodes": [{"sql": "insert into FUSION_HOTLINE_ORDER select ID, AREA_NAME from SRC_SECRET"}],
            "fields": [{"字段中文名": "区域名称"}, {"字段中文名": "工单数量"}],
        }
        resource_info = {
            "SRC_SECRET": {"资源名称": "属地返还加密明细表"},
        }

        steps = module.build_business_logic_steps(record, resource_info)
        text = "\n".join(steps)

        for banned in ["清洗", "抽取", "加密", "质量检查"]:
            self.assertNotIn(banned, text)
        self.assertIn("数据范围确认", text)
        self.assertIn("字段口径", text)

    def test_dispatches_to_stats_result_workbook_builder(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = temp / "ledger.xlsx"
            template = temp / "template.xlsx"
            catalog = temp / "catalog.xlsx"
            output = temp / "out.xlsx"
            make_ledger(ledger)
            openpyxl.Workbook().save(template)
            openpyxl.Workbook().save(catalog)

            calls = []

            def fake_builder(excel_path, service_dir, template_path, output_path, catalog_path):
                calls.append((excel_path, service_dir, template_path, output_path, catalog_path))
                Path(output_path).write_bytes(b"generated")
                return output_path

            with mock.patch.object(module, "fill_stats_result_usage_workbook", fake_builder, create=True):
                result = module.fill_document(
                    excel_path=str(ledger),
                    service_dir="N08-数据统计分析",
                    material_type="04-数据统计分析_结果表及使用说明",
                    template_path=str(template),
                    output_path=str(output),
                    catalog_path=str(catalog),
                )

            self.assertEqual(result, str(output))
            self.assertEqual(calls, [(str(ledger), "N08-数据统计分析", str(template), str(output), str(catalog))])
            self.assertEqual(output.read_bytes(), b"generated")

    def test_result_workbook_builder_accepts_new_result_list_column_name(self):
        module = load_builder_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = temp / "ledger.xlsx"
            make_new_column_ledger(ledger)

            rows = module.load_ledger_rows(ledger, "N08-数据统计分析")

        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["统计分析结果表清单"], "测试结果表 TEST_RESULT")

    def test_draw_flowchart_png_uses_transparent_background_and_no_push_node(self):
        module = load_builder_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            image_path = temp / "flowchart.png"
            record = {
                "row": 1,
                "result_cn": "结果表一",
                "result_en": "RESULT_ONE",
                "sources": ["SRC_ONE", "SRC_TWO"],
                "resource_info": {
                    "SRC_ONE": {"资源名称": "源表一"},
                    "SRC_TWO": {"资源名称": "源表二"},
                },
            }

            module.draw_flowchart_png(record, image_path)

            with Image.open(image_path) as img:
                self.assertEqual(img.mode, "RGBA")
                self.assertEqual(img.getpixel((0, 0))[3], 0)
                region = img.crop((445, 800, 835, 910))
                opaque_pixels = sum(
                    count
                    for count, color in region.getcolors(maxcolors=region.width * region.height)
                    if color[3] > 0
                )

            self.assertLess(opaque_pixels, 4000)

    def test_draw_flowchart_png_does_not_route_merge_line_through_staggered_source_node(self):
        module = load_builder_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            image_path = temp / "flowchart.png"
            record = {
                "row": 1,
                "result_cn": "RESULT_CN",
                "result_en": "RESULT_EN",
                "sources": ["SRC_ONE", "SRC_TWO", "SRC_THREE", "SRC_FOUR", "SRC_FIVE"],
                "resource_info": {},
            }

            module.draw_flowchart_png(record, image_path)

            with Image.open(image_path) as img:
                lower_source_inner = img.crop((510, 395, 770, 475))
                line_pixels = sum(
                    count
                    for count, color in lower_source_inner.getcolors(
                        maxcolors=lower_source_inner.width * lower_source_inner.height
                    )
                    if color == (43, 43, 43, 255)
                )

            self.assertEqual(line_pixels, 0)

    def test_result_workbook_applies_songti_gray_fill_defaults_and_relation_spacing(self):
        module = load_builder_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            template = temp / "template.xlsx"
            output = temp / "out.xlsx"
            make_result_template(template)
            records = make_records_with_flowcharts(temp)
            resource_info = {
                "SRC_ONE": {"数据目录代码": "DIR-SRC-1", "资源名称": "源表一", "单位名称": "原始单位一", "资源类型": "原始类型一"},
                "SRC_TWO": {"数据目录代码": "DIR-SRC-2", "资源名称": "源表二", "单位名称": "原始单位二", "资源类型": "原始类型二"},
                "RESULT_ONE": {"数据目录代码": "DIR-RESULT-1", "资源名称": "结果表一", "单位名称": "结果单位一", "资源类型": "结果类型一"},
                "RESULT_TWO": {"资源名称": "结果表二"},
            }

            module.build_workbook(template, output, records, resource_info)

            wb = openpyxl.load_workbook(output)
            source_list = wb["1、数据源表list"]
            relation = wb["2、表融合关系"]
            result_list = wb["3、数据统计分析结果表list"]
            result_detail = wb["4、数据统计分析结果表详情"]

            self.assertEqual(source_list["D2"].value, "上海市大数据中心")
            self.assertEqual(source_list["E2"].value, "库表")
            self.assertEqual(source_list["D3"].value, "上海市大数据中心")
            self.assertEqual(source_list["E3"].value, "库表")
            self.assertEqual(
                [result_list.cell(1, col).value for col in range(1, 7)],
                ["序号", "资源编目（非必填）", "资源名称", "数据提供方", "资源类型", "数据统计分析表的资源信息（表名）"],
            )
            self.assertEqual([result_list.cell(2, col).value for col in range(1, 7)], [1, "DIR-RESULT-1", "结果表一", "结果单位一", "结果类型一", "RESULT_ONE"])
            self.assertEqual([result_list.cell(3, col).value for col in range(1, 7)], [2, None, "结果表二", "上海市大数据中心", "库表", "RESULT_TWO"])
            self.assertNotIn("A1:C1", [str(merged) for merged in relation.merged_cells.ranges])
            self.assertEqual(relation["A1"].value, "2.1 RESULT_ONE")
            self.assertIsNone(relation["B1"].value)
            self.assertIsNone(relation["C1"].value)
            self.assertIsNone(relation["A1"].fill.fill_type)
            self.assertIsNone(relation["B2"].fill.fill_type)
            self.assertIsNone(relation["B3"].fill.fill_type)
            self.assertIsNone(relation["A4"].value)
            self.assertEqual(relation.row_dimensions[4].height, 20)
            self.assertNotIn("A5:C5", [str(merged) for merged in relation.merged_cells.ranges])
            self.assertEqual(relation["A5"].value, "2.2 RESULT_TWO")
            self.assertIsNone(relation["A5"].fill.fill_type)
            self.assertIsNone(relation["B6"].fill.fill_type)
            self.assertIsNone(relation["B7"].fill.fill_type)
            detail_merges = [str(merged) for merged in result_detail.merged_cells.ranges]
            self.assertNotIn("A1:H1", detail_merges)
            self.assertNotIn("B2:H2", detail_merges)
            self.assertEqual(result_detail["A1"].value, "4.1 结果表一")
            self.assertIsNone(result_detail["A1"].fill.fill_type)
            self.assertIsNone(result_detail["B1"].value)
            self.assertEqual(result_detail["B2"].value, "RESULT_ONE")
            self.assertTrue(result_detail["B2"].font.bold)
            self.assertEqual(result_detail["B2"].alignment.horizontal, "center")
            self.assertIsNone(result_detail["C2"].value)
            self.assertNotIn("A6:H6", detail_merges)
            self.assertNotIn("B7:H7", detail_merges)
            self.assertEqual(result_detail["A6"].value, "4.2 结果表二")
            self.assertIsNone(result_detail["A6"].fill.fill_type)
            self.assertEqual(result_detail["B7"].value, "RESULT_TWO")

            for worksheet in wb.worksheets:
                for row in worksheet.iter_rows():
                    for cell in row:
                        if cell.value not in (None, ""):
                            self.assertEqual(cell.font.name, "宋体")

            for cell in [source_list["A1"], wb["3、数据统计分析结果表list"]["A1"], wb["4、数据统计分析结果表详情"]["B3"]]:
                self.assertEqual(cell.fill.fgColor.rgb, "00F2F2F2")
            wb.close()


if __name__ == "__main__":
    unittest.main()


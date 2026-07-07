import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "fill_document.py"


def load_fill_document_module():
    spec = importlib.util.spec_from_file_location("fill_document", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def make_split_ledger(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "服务目录",
        "需求单号",
        "工单号",
        "工单内容",
        "报表统计次数",
        "业务描述",
        "统计分析结果表清单",
        "程序XML文本",
        "数据统计分析执行周期",
        "数据更新要求",
        "数据量对后续运维的特殊要求",
    ]
    ws.append(headers)
    ws.append(
        [
            "N08-数据统计分析",
            "REQ-1",
            "WO-1",
            "工单一",
            "3",
            "业务描述一",
            "结果表一 RESULT_ONE",
            "<mxGraphModel />",
            "按日更新",
            "每日更新",
            "无",
        ]
    )
    ws.append([None, None, None, None, None, None, "结果表二 RESULT_TWO", "<mxGraphModel />", None, None, None])
    ws.append([None, None, None, None, None, None, "结果表三 RESULT_THREE", "<mxGraphModel />", None, None, None])
    for column in ["A", "B", "C", "D", "E", "F", "I", "J", "K"]:
        ws.merge_cells(f"{column}2:{column}4")
    wb.save(path)


def make_new_column_split_ledger(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    headers = [
        "服务目录",
        "需求单号",
        "工单号",
        "工单标题",
        "程序数",
        "工单描述",
        "结果表清单",
        "程序XML文本",
        "数据统计分析执行周期",
        "数据更新要求",
        "数据量对后续运维的特殊要求",
    ]
    ws.append(headers)
    ws.append(
        [
            "N08-数据统计分析",
            "REQ-1",
            "WO-1",
            "新版工单标题",
            "3",
            "新版工单描述",
            "结果表一 RESULT_ONE",
            "<mxGraphModel />",
            "按日更新",
            "每日更新",
            "无",
        ]
    )
    ws.append([None, None, None, None, None, None, "结果表二 RESULT_TWO", "<mxGraphModel />", None, None, None])
    ws.append([None, None, None, None, None, None, "结果表三 RESULT_THREE", "<mxGraphModel />", None, None, None])
    for column in ["A", "B", "C", "D", "E", "F", "I", "J", "K"]:
        ws.merge_cells(f"{column}2:{column}4")
    wb.save(path)


def make_relation_workbook(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2、表融合关系"
    ws["A1"] = "2.1 RESULT_ONE"
    ws["B2"] = "文字描述"
    ws["C2"] = "1\t数据来源准备：读取源表A。\n2\t结果输出：写入结果表一。"
    ws["A4"] = "2.2 RESULT_TWO"
    ws["B5"] = "文字描述"
    ws["C5"] = "1\t数据来源准备：读取源表B。\n2\t结果输出：写入结果表二。"
    wb.save(path)


class StatsRequirementSplitRowsTest(unittest.TestCase):
    def test_groups_split_result_rows_by_work_order(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            ledger = Path(temp_dir) / "ledger.xlsx"
            make_split_ledger(ledger)

            groups = module.read_stats_requirement_groups(str(ledger), "N08-数据统计分析")

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["工单号"], "WO-1")
        self.assertEqual(groups[0]["报表统计次数"], "3")
        self.assertEqual(
            groups[0]["results"],
            [
                ("结果表一", "RESULT_ONE"),
                ("结果表二", "RESULT_TWO"),
                ("结果表三", "RESULT_THREE"),
            ],
        )

    def test_groups_accept_new_ledger_column_names(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            ledger = Path(temp_dir) / "ledger.xlsx"
            make_new_column_split_ledger(ledger)

            groups = module.read_stats_requirement_groups(str(ledger), "N08-数据统计分析")

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0]["工单内容"], "新版工单标题")
        self.assertEqual(groups[0]["报表统计次数"], "3")
        self.assertEqual(groups[0]["业务描述"], "新版工单描述")
        self.assertEqual(groups[0]["业务说明"], "新版工单描述")
        self.assertEqual(groups[0]["统计分析结果表清单"], "结果表一 RESULT_ONE\n结果表二 RESULT_TWO\n结果表三 RESULT_THREE")

    def test_loads_relation_descriptions_by_result_table_name(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            relation = Path(temp_dir) / "relation.xlsx"
            make_relation_workbook(relation)

            descriptions = module.load_stats_relation_descriptions(str(relation))

        self.assertIn("RESULT_ONE", descriptions)
        self.assertIn("RESULT_TWO", descriptions)
        self.assertIn("读取源表A", descriptions["RESULT_ONE"])

    def test_single_result_description_is_split_into_multiple_logic_steps(self):
        module = load_fill_document_module()

        steps = module.build_stats_requirement_business_logic_steps(
            [("结果表一", "RESULT_ONE")],
            {
                "RESULT_ONE": (
                    "1\t数据来源准备：读取源表A作为基础数据。\n"
                    "2\t数据筛选清洗：过滤无效记录并统一统计口径。\n"
                    "3\t结果输出：写入结果表一供统计分析使用。"
                )
            },
        )

        self.assertEqual([step for step, _ in steps], ["1", "2", "3"])
        self.assertIn("读取源表A", steps[0][1])
        self.assertIn("过滤无效记录", steps[1][1])
        self.assertNotIn("清洗", steps[1][1])
        self.assertIn("数据范围确认", steps[1][1])
        self.assertIn("写入结果表一", steps[2][1])

    def test_dispatch_uses_split_groups_and_sibling_relation_workbook(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = temp / "ledger.xlsx"
            template = temp / "01-数据统计分析_需求文档.docx"
            relation = temp / "04-数据统计分析_结果表及使用说明.xlsx"
            output = temp / "out.docx"
            make_split_ledger(ledger)
            make_relation_workbook(relation)
            template.write_bytes(b"placeholder")

            calls = []

            def fake_fill(excel_path, data_rows, template_path, output_path, relation_path=None):
                calls.append((excel_path, data_rows, template_path, output_path, relation_path))
                Path(output_path).write_bytes(b"generated")
                return output_path

            with mock.patch.object(module, "fill_stats_requirement_doc", fake_fill):
                result = module.fill_document(
                    excel_path=str(ledger),
                    service_dir="N08-数据统计分析",
                    material_type="01-数据统计分析_需求文档",
                    template_path=str(template),
                    output_path=str(output),
                )

            self.assertEqual(result, str(output))
            self.assertEqual(len(calls), 1)
            self.assertEqual(len(calls[0][1]), 1)
            self.assertEqual(len(calls[0][1][0]["results"]), 3)
            self.assertEqual(calls[0][4], str(relation))

    def test_dispatch_prefers_generated_relation_workbook_next_to_output(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            template_dir = temp / "template"
            output_dir = temp / "output"
            template_dir.mkdir()
            output_dir.mkdir()
            ledger = temp / "ledger.xlsx"
            template = template_dir / "01-数据统计分析_需求文档.docx"
            stale_relation = template_dir / "04-数据统计分析_结果表及使用说明.xlsx"
            fresh_relation = output_dir / "04-数据统计分析_结果表及使用说明.xlsx"
            output = output_dir / "01-数据统计分析_需求文档.docx"
            make_split_ledger(ledger)
            make_relation_workbook(stale_relation)
            make_relation_workbook(fresh_relation)
            template.write_bytes(b"placeholder")

            calls = []

            def fake_fill(excel_path, data_rows, template_path, output_path, relation_path=None):
                calls.append((excel_path, data_rows, template_path, output_path, relation_path))
                Path(output_path).write_bytes(b"generated")
                return output_path

            with mock.patch.object(module, "fill_stats_requirement_doc", fake_fill):
                result = module.fill_document(
                    excel_path=str(ledger),
                    service_dir="N08-数据统计分析",
                    material_type="01-数据统计分析_需求文档",
                    template_path=str(template),
                    output_path=str(output),
                )

            self.assertEqual(result, str(output))
            self.assertEqual(calls[0][4], str(fresh_relation))


if __name__ == "__main__":
    unittest.main()

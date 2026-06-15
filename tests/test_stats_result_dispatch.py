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


def make_ledger(path: Path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["服务目录", "程序XML文本", "统计分析结果表清单"])
    ws.append(["N08-数据统计分析", "<mxGraphModel />", "测试结果表 TEST_RESULT"])
    wb.save(path)


class StatsResultDispatchTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()


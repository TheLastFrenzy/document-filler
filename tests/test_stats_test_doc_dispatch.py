import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
from pypdf import PdfReader


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
    ws.append(["服务目录", "需求单号", "工单号", "工单内容", "统计分析结果表清单", "03-数据统计分析_测试文档_工单自测报告附件"])
    ws.append(["N08-数据统计分析", "REQ-1", "WO-1", "工单一", "结果表一 RESULT_ONE", None])
    wb.save(path)


class StatsTestDocDispatchTest(unittest.TestCase):
    def test_dispatches_to_stats_test_pdf_generator(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = temp / "ledger.xlsx"
            template = temp / "03-数据统计分析_测试文档.pdf"
            output = temp / "out.pdf"
            make_ledger(ledger)
            template.write_bytes(b"%PDF-1.4\n% placeholder\n")

            calls = []

            def fake_generator(excel_path, service_dir, template_path, output_path):
                calls.append((excel_path, service_dir, template_path, output_path))
                Path(output_path).write_bytes(b"%PDF-1.4\n% generated\n")
                return output_path

            with mock.patch.object(module, "fill_stats_test_pdf", fake_generator, create=True):
                result = module.fill_document(
                    excel_path=str(ledger),
                    service_dir="N08-数据统计分析",
                    material_type="03-数据统计分析_测试文档",
                    template_path=str(template),
                    output_path=str(output),
                )

            self.assertEqual(result, str(output))
            self.assertEqual(calls, [(str(ledger), "N08-数据统计分析", str(template), str(output))])
            self.assertEqual(output.read_bytes(), b"%PDF-1.4\n% generated\n")

    def test_generates_pdf_with_original_conclusion_totals(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            ledger = temp / "ledger.xlsx"
            template = temp / "03-数据统计分析_测试文档.pdf"
            output = temp / "out.pdf"

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.append(["服务目录", "需求单号", "工单号", "工单内容", "统计分析结果表清单", "03-数据统计分析_测试文档_工单自测报告附件"])
            ws.append(["N08-数据统计分析", "REQ-1", "WO-1", "工单一", "结果表一 RESULT_ONE", None])
            ws.append([None, None, None, None, "结果表二 RESULT_TWO", None])
            for column in ["A", "B", "C", "D"]:
                ws.merge_cells(f"{column}2:{column}3")
            wb.save(ledger)
            template.write_bytes(b"%PDF-1.4\n% placeholder\n")

            result = module.fill_document(
                excel_path=str(ledger),
                service_dir="N08-数据统计分析",
                material_type="03-数据统计分析_测试文档",
                template_path=str(template),
                output_path=str(output),
            )

            self.assertEqual(result, str(output))
            self.assertTrue(output.exists())
            pdf = PdfReader(str(output))
            last_text = pdf.pages[-1].extract_text()
            self.assertIn("经测试验证", last_text)
            self.assertIn("数据测试", last_text)
            self.assertIn("程序测试", last_text)
            self.assertGreaterEqual(last_text.count("2"), 4)
            self.assertIn("100%", last_text)


if __name__ == "__main__":
    unittest.main()

import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import fitz
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


def make_template_pdf(path: Path):
    doc = fitz.open()
    for index in range(21):
        page = doc.new_page(width=595, height=842)
        page.insert_text((72, 72), f"TEMPLATE_PAGE_{index + 1}", fontsize=12)
    doc[0].insert_text((72, 110), "TEMPLATE_COVER_MARKER", fontsize=12)
    doc[1].insert_text((72, 110), "TEMPLATE_REVISION_MARKER", fontsize=12)
    doc[18].insert_text((72, 110), "TEMPLATE_STATIC_CHAPTER_MARKER", fontsize=12)
    doc.save(path)
    doc.close()


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
            make_template_pdf(template)

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
            self.assertIn("TEMPLATE_COVER_MARKER", pdf.pages[0].extract_text())
            self.assertIn("TEMPLATE_REVISION_MARKER", pdf.pages[1].extract_text())
            self.assertIn("TEMPLATE_STATIC_CHAPTER_MARKER", pdf.pages[4].extract_text())
            last_text = pdf.pages[-1].extract_text()
            self.assertGreaterEqual(last_text.count("2"), 4)
            self.assertIn("100%", last_text)
            link_doc = fitz.open(str(output))
            links = link_doc[2].get_links() + link_doc[3].get_links()
            goto_targets = [link["page"] for link in links if link.get("kind") == fitz.LINK_GOTO]
            self.assertGreaterEqual(len(goto_targets), 14)
            self.assertIn(7, goto_targets)
            self.assertIn(8, goto_targets)
            self.assertEqual(max(goto_targets), link_doc.page_count - 1)
            link_doc.close()


if __name__ == "__main__":
    unittest.main()

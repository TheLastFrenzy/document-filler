import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import openpyxl
from PIL import Image, ImageDraw
from docx.oxml.ns import qn


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "fill_document.py"
CATALOG_COL = "02-数据报表_设计文档-数据来源库表清单对应数据目录代码"
PNG_1X1 = (
    b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01"
    b"\x00\x00\x00\x01\x08\x02\x00\x00\x00\x90wS\xde\x00"
    b"\x00\x00\x0cIDATx\x9cc\xf8\xff\xff?\x00\x05\xfe\x02"
    b"\xfeA\xe2!Q\x00\x00\x00\x00IEND\xaeB`\x82"
)


def load_fill_document_module():
    spec = importlib.util.spec_from_file_location("fill_document", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


def content_bbox(path: Path):
    image = Image.open(path).convert("RGB")
    mask = image.point(lambda value: 255 if value < 245 else 0)
    return mask.getbbox(), image.size


class DataReportRegressionTest(unittest.TestCase):
    def test_requirement_text_adds_etc_for_partial_catalog_codes_and_sanitizes_delivery(self):
        module = load_fill_document_module()
        row = {
            "数据需求": "本次需求围绕002420412/000269、002420412/000114目录开展统计核验。",
            "交付要求": "验收时抽查929945317274169及【字段名称】并复核检查结果。",
            CATALOG_COL: "002420412/000269\n002420412/000114\nMB2F30661/000152",
            "工单内容": "数据目录字段统计",
            "业务说明": "根据附件梳理字段名称、数据类型、空值数等内容。",
        }

        normalized = module.normalize_data_report_text_fields(row)

        self.assertIn("002420412/000269、002420412/000114等目录", normalized["数据需求"])
        for banned in ["检查", "抽查", "复核"]:
            self.assertNotIn(banned, normalized["交付要求"])
        self.assertNotIn("929945317274169", normalized["交付要求"])
        self.assertFalse(any(mark in normalized["交付要求"] for mark in "【】[]"))

    def test_select_excel_preview_range_prefers_richer_sheet_and_caps_wide_ranges(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "book.xlsx"
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "empty"
            ws["A1"] = "only one cell"
            rich = wb.create_sheet("rich")
            for row in range(1, 21):
                for col in range(1, 16):
                    rich.cell(row, col).value = f"R{row}C{col}"
            wb.save(path)

            sheet_name, first_row, first_col, last_row, last_col = module.select_excel_preview_range(path)

        self.assertEqual(sheet_name, "rich")
        self.assertEqual((first_row, first_col), (1, 1))
        self.assertEqual(last_row, 20)
        self.assertEqual(last_col - first_col + 1, 9)

    def test_crop_and_normalize_image_removes_large_blank_margins_and_keeps_resolution(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "preview.png"
            image = Image.new("RGB", (1000, 800), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((350, 280, 650, 420), fill="black")
            image.save(path)

            self.assertTrue(module.crop_and_normalize_image(path, min_width=1200, min_height=500))
            bbox, size = content_bbox(path)

        self.assertIsNotNone(bbox)
        width, height = size
        left, top, right, bottom = bbox
        self.assertGreaterEqual(width, 1200)
        self.assertGreaterEqual(height, 500)
        max_margin = max(left / width, top / height, (width - right) / width, (height - bottom) / height)
        self.assertLess(max_margin, 0.18)

    def test_attachment_preview_uses_excel_range_rendering_before_pdf_fallback(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            source = temp / "attachment.xlsx"
            openpyxl.Workbook().save(source)
            calls = []

            def fake_excel_range_to_png(excel_path, image_path):
                calls.append((excel_path, image_path))
                image_path.write_bytes(PNG_1X1)
                return True

            with mock.patch.object(module, "excel_range_to_png", fake_excel_range_to_png, create=True):
                preview = module.generate_attachment_screenshot_bytes([source], temp, row_number=2)

        self.assertEqual(preview, PNG_1X1)
        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], source)

    def test_launch_record_helpers_apply_fixed_id_format_indent_and_word_wrap(self):
        module = load_fill_document_module()
        row = {"需求单号": "X_A_RCGZ_202601070119", "工单号": "G_A_RCGZ_202601070180"}

        self.assertEqual(
            module.build_launch_identifier_line(row),
            "需求编号：X_A_RCGZ_202601070119\t对应工单编号：G_A_RCGZ_202601070180",
        )

        paragraph = module.mp("统计报表：3次。", "Normal", 480, word_wrap=True)
        p_pr = paragraph.find(qn("w:pPr"))
        self.assertIsNotNone(p_pr.find(qn("w:wordWrap")))
        indent = p_pr.find(qn("w:ind"))
        self.assertEqual(indent.get(qn("w:firstLine")), "480")


if __name__ == "__main__":
    unittest.main()

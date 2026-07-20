import sys
import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPTS = ROOT / "scripts"
if str(SCRIPTS) not in sys.path:
    sys.path.insert(0, str(SCRIPTS))


class SharedOfficeWordTest(unittest.TestCase):
    def test_docx_template_is_returned_without_conversion(self):
        from materials.shared.office_word import convert_legacy_doc_template

        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "template.docx"
            template.write_bytes(b"docx")

            with mock.patch("materials.shared.office_word.subprocess.run") as run:
                result = convert_legacy_doc_template(template, Path(temp_dir) / "work")

        self.assertEqual(result, template)
        run.assert_not_called()

    def test_unsupported_template_extension_keeps_existing_error_contract(self):
        from materials.shared.office_word import convert_legacy_doc_template

        with tempfile.TemporaryDirectory() as temp_dir:
            template = Path(temp_dir) / "template.pdf"

            with self.assertRaisesRegex(
                ValueError,
                r"API接口测试报告模板仅支持 \.doc 或 \.docx",
            ):
                convert_legacy_doc_template(template, Path(temp_dir) / "work")

    def test_legacy_template_conversion_creates_work_directory_and_checks_output(self):
        from materials.shared.office_word import convert_legacy_doc_template

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            template = temp / "template.doc"
            template.write_bytes(b"legacy")
            work_dir = temp / "missing" / "work"

            def fake_run(*_args, **_kwargs):
                converted = work_dir / "template.docx"
                self.assertTrue(converted.parent.exists())
                converted.write_bytes(b"converted")
                return SimpleNamespace(stdout="CONVERTED", stderr="", returncode=0)

            with mock.patch("materials.shared.office_word.subprocess.run", side_effect=fake_run):
                result = convert_legacy_doc_template(template, work_dir)

        self.assertEqual(result, work_dir / "template.docx")

    def test_docx_save_uses_injected_toc_updater(self):
        from materials.shared.office_word import save_word_document

        class FakeDocument:
            def save(self, path):
                Path(path).write_bytes(b"docx")

        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "result" / "output.docx"
            toc_calls = []
            result = save_word_document(
                FakeDocument(),
                output,
                Path(temp_dir) / "work",
                toc_updater=lambda path: toc_calls.append(Path(path)),
            )

        self.assertEqual(result, str(output))
        self.assertEqual(toc_calls, [output])

    def test_legacy_save_uses_injected_converter(self):
        from materials.shared.office_word import save_word_document

        class FakeDocument:
            def save(self, path):
                Path(path).write_bytes(b"docx")

        with tempfile.TemporaryDirectory() as temp_dir:
            temp = Path(temp_dir)
            output = temp / "result" / "output.doc"
            work_dir = temp / "missing" / "work"
            converter_calls = []

            def converter(docx_path, output_path):
                converter_calls.append((Path(docx_path), Path(output_path)))
                Path(output_path).write_bytes(b"legacy")
                return Path(output_path)

            result = save_word_document(
                FakeDocument(),
                output,
                work_dir,
                legacy_converter=converter,
            )

        self.assertEqual(result, str(output))
        self.assertEqual(converter_calls, [(work_dir / "output.docx", output)])

    def test_existing_private_compatibility_entrypoints_remain_available(self):
        from materials.n07 import api_launch_record, api_test_report

        self.assertTrue(callable(api_launch_record._convert_docx_to_legacy_doc))
        self.assertTrue(callable(api_launch_record._save_document))
        self.assertTrue(callable(api_test_report._convert_legacy_doc_template))


if __name__ == "__main__":
    unittest.main()

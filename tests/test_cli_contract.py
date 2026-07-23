import importlib.util
import tempfile
import unittest
from pathlib import Path
from unittest import mock


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "fill_document.py"


def load_fill_document_module():
    spec = importlib.util.spec_from_file_location("fill_document", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class CliContractTest(unittest.TestCase):
    def test_registry_exposes_canonical_material_types_without_filename_aliases(self):
        module = load_fill_document_module()

        self.assertEqual(
            module.registered_material_types(),
            (
                "01-API接口开发_需求文档",
                "02-API接口开发_数据模型设计",
                "03-API接口开发_接口开发代码",
                "04-API接口开发_接口测试报告",
                "05-API接口开发_作业上线记录",
                "01-库表落地方式_需求文档",
                "02-库表落地方式_数据模型设计",
                "04-库表落地方式_测试报告",
                "05-库表落地方式_作业上线记录",
                "06-库表落地方式_共享记录",
                "03-N02材料_测试文档",
            ),
        )

    def test_registry_alias_resolves_to_the_canonical_spec(self):
        module = load_fill_document_module()

        canonical = module.get_material_spec("04-API接口开发_接口测试报告")
        alias = module.get_material_spec("04-接口测试报告（含《API接口列表》）")
        n02_alias = module.get_material_spec("03-测试文档")

        self.assertIs(alias, canonical)
        self.assertIsNotNone(n02_alias)
        self.assertEqual(n02_alias.default_filename, "03-测试文档.docx")

    def test_supported_material_types_combines_registered_and_legacy_dispatch(self):
        module = load_fill_document_module()

        supported = module.supported_material_types()

        self.assertEqual(len(supported), 18)
        self.assertIn("01-API接口开发_需求文档", supported)
        self.assertIn("04-数据统计分析_结果表及使用说明", supported)
        self.assertNotIn("03-接口开发代码", supported)

    def test_unsupported_material_error_is_generated_from_supported_types(self):
        module = load_fill_document_module()

        with self.assertRaises(ValueError) as context:
            module.fill_document(
                excel_path="ledger.xlsx",
                service_dir="N08-数据报表服务",
                material_type="不存在的材料",
                template_path="template.docx",
                output_path="output.docx",
            )

        message = str(context.exception)
        self.assertIn("不支持的材料类型: 不存在的材料", message)
        self.assertIn("01-API接口开发_需求文档", message)
        self.assertIn("04-数据统计分析_结果表及使用说明", message)

    def test_output_directory_resolves_to_material_filename_and_extension(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "outputs"
            output_dir.mkdir()
            expected_output = output_dir / "04-数据统计分析_结果表及使用说明.xlsx"
            calls = []

            def fake_builder(excel_path, service_dir, template_path, output_path, catalog_path):
                calls.append(output_path)
                Path(output_path).write_bytes(b"generated")
                return output_path

            with mock.patch.object(module, "fill_stats_result_usage_workbook", fake_builder):
                result = module.fill_document(
                    excel_path="ledger.xlsx",
                    service_dir="N08-数据统计分析",
                    material_type="04-数据统计分析_结果表及使用说明",
                    template_path="template.xlsx",
                    output_path=str(output_dir),
                    catalog_path="catalog.xlsx",
                )

            self.assertEqual(result, str(expected_output))
            self.assertEqual(calls, [str(expected_output)])
            self.assertEqual(expected_output.read_bytes(), b"generated")

    def test_output_directory_uses_docx_extension_for_stats_test_document(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "outputs"
            output_dir.mkdir()
            expected_output = output_dir / "03-数据统计分析_测试文档.docx"
            calls = []

            def fake_docx(excel_path, service_dir, template_path, output_path):
                calls.append(output_path)
                Path(output_path).write_bytes(b"generated docx")
                return output_path

            with mock.patch.object(module, "fill_stats_test_docx", fake_docx):
                result = module.fill_document(
                    excel_path="ledger.xlsx",
                    service_dir="N08-数据统计分析",
                    material_type="03-数据统计分析_测试文档",
                    template_path="template.docx",
                    output_path=str(output_dir),
                )

            self.assertEqual(result, str(expected_output))
            self.assertEqual(calls, [str(expected_output)])

    def test_output_directory_uses_docx_extension_for_n02_test_document(self):
        module = load_fill_document_module()
        with tempfile.TemporaryDirectory() as temp_dir:
            output_dir = Path(temp_dir) / "outputs"
            output_dir.mkdir()
            expected_output = output_dir / "03-测试文档.docx"
            calls = []

            def fake_builder(excel_path, service_dir, template_path, output_path):
                calls.append((excel_path, service_dir, template_path, output_path))
                Path(output_path).write_bytes(b"generated n02")
                return output_path

            with mock.patch.object(module, "load_material_builder", return_value=fake_builder):
                result = module.fill_document(
                    excel_path="ledger.xlsx",
                    service_dir="N02材料",
                    material_type="03-测试文档",
                    template_path="template.docx",
                    output_path=str(output_dir),
                )

            self.assertEqual(result, str(expected_output))
            self.assertEqual(calls, [("ledger.xlsx", "N02材料", "template.docx", str(expected_output))])
            self.assertEqual(expected_output.read_bytes(), b"generated n02")

    def test_missing_optional_dependency_reports_install_command_without_pip_install(self):
        module = load_fill_document_module()
        with mock.patch("importlib.util.find_spec", return_value=None), mock.patch("subprocess.check_call") as check_call:
            with self.assertRaisesRegex(
                ImportError,
                r"缺少 Python 依赖: pymupdf.*pip install -r requirements.txt",
            ):
                module.ensure_module("fitz", "pymupdf")

        check_call.assert_not_called()


if __name__ == "__main__":
    unittest.main()

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl


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


if __name__ == "__main__":
    unittest.main()

import importlib.util
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / "scripts" / "build_stats_result_usage_workbook.py"


def load_builder_module():
    spec = importlib.util.spec_from_file_location("build_stats_result_usage_workbook", SCRIPT)
    module = importlib.util.module_from_spec(spec)
    assert spec.loader is not None
    spec.loader.exec_module(module)
    return module


class StatsResultFieldParsingTest(unittest.TestCase):
    def test_parses_fields_after_oracle_dialect_line_comments(self):
        module = load_builder_module()
        sql = """
        CREATE TABLE IF NOT EXISTS SHIGONGANJU_FUSION.temp_FUSION_GAJ_WPINFO_VISIT_12345_DSJZX_V2(
          `wpid` varchar2(50) DEFAULT NULL COMMENT '工单编号',
          -- dialect: ORACLE
          `hf_state` varchar2(50) DEFAULT NULL COMMENT '回访方式',
          -- dialect: ORACLE
          `visit_time` varchar2(50) DEFAULT NULL COMMENT '回访时间',
          `dept_level2` varchar2(100) DEFAULT NULL COMMENT '主办单位'
        )
        """

        fields = module.parse_result_fields_from_sql(
            [sql],
            "FUSION_GAJ_WPINFO_VISIT_12345_DSJZX_V2",
        )

        self.assertEqual(
            [field["字段英文名"] for field in fields],
            ["WPID", "HF_STATE", "VISIT_TIME", "DEPT_LEVEL2"],
        )
        self.assertEqual(
            [field["字段中文名"] for field in fields],
            ["工单编号", "回访方式", "回访时间", "主办单位"],
        )


if __name__ == "__main__":
    unittest.main()

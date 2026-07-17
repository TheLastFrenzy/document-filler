from dataclasses import dataclass
from importlib import import_module


@dataclass(frozen=True)
class MaterialSpec:
    default_filename: str
    module: str
    function: str


SPECS = {
    "01-API接口开发_需求文档": MaterialSpec(
        default_filename="01-需求文档.docx",
        module="materials.n07.api_requirement",
        function="build_api_requirement_document",
    ),
    "02-API接口开发_数据模型设计": MaterialSpec(
        default_filename="02- 数据模型设计（API）.docx",
        module="materials.n07.api_data_model",
        function="build_api_data_model_document",
    ),
    "02-数据模型设计（API）": MaterialSpec(
        default_filename="02- 数据模型设计（API）.docx",
        module="materials.n07.api_data_model",
        function="build_api_data_model_document",
    ),
    "02- 数据模型设计（API）": MaterialSpec(
        default_filename="02- 数据模型设计（API）.docx",
        module="materials.n07.api_data_model",
        function="build_api_data_model_document",
    ),
    "03-API接口开发_接口开发代码": MaterialSpec(
        default_filename="03-接口开发代码.docx",
        module="materials.n07.api_code_doc",
        function="build_api_code_document",
    ),
    "03-接口开发代码": MaterialSpec(
        default_filename="03-接口开发代码.docx",
        module="materials.n07.api_code_doc",
        function="build_api_code_document",
    ),
    "04-API接口开发_接口测试报告": MaterialSpec(
        default_filename="04-接口测试报告（含《API接口列表》）.docx",
        module="materials.n07.api_test_report",
        function="build_api_test_report_document",
    ),
    "04-接口测试报告（含《API接口列表》）": MaterialSpec(
        default_filename="04-接口测试报告（含《API接口列表》）.docx",
        module="materials.n07.api_test_report",
        function="build_api_test_report_document",
    ),
    "05-API接口开发_作业上线记录": MaterialSpec(
        default_filename="05-作业上线记录（API）.doc",
        module="materials.n07.api_launch_record",
        function="build_api_launch_record_document",
    ),
    "05-作业上线记录（API）": MaterialSpec(
        default_filename="05-作业上线记录（API）.doc",
        module="materials.n07.api_launch_record",
        function="build_api_launch_record_document",
    ),
    "01-库表落地方式_需求文档": MaterialSpec(
        default_filename="01-需求文档.docx",
        module="materials.n07.table_landing_requirement",
        function="build_table_landing_requirement_document",
    ),
    "02-库表落地方式_数据模型设计": MaterialSpec(
        default_filename="02-数据模型设计.doc",
        module="materials.n07.table_landing_design",
        function="build_table_landing_design_document",
    ),
    "02-数据模型设计": MaterialSpec(
        default_filename="02-数据模型设计.doc",
        module="materials.n07.table_landing_design",
        function="build_table_landing_design_document",
    ),
    "04-库表落地方式_测试报告": MaterialSpec(
        default_filename="04- 测试报告（库表落地）.doc",
        module="materials.n07.table_landing_test_report",
        function="build_table_landing_test_report_document",
    ),
    "04- 测试报告（库表落地）": MaterialSpec(
        default_filename="04- 测试报告（库表落地）.doc",
        module="materials.n07.table_landing_test_report",
        function="build_table_landing_test_report_document",
    ),
    "05-库表落地方式_作业上线记录": MaterialSpec(
        default_filename="05-作业上线记录.doc",
        module="materials.n07.table_landing_launch_record",
        function="build_table_landing_launch_record_document",
    ),
    "05-作业上线记录": MaterialSpec(
        default_filename="05-作业上线记录.doc",
        module="materials.n07.table_landing_launch_record",
        function="build_table_landing_launch_record_document",
    ),
    "06-库表落地方式_共享记录": MaterialSpec(
        default_filename="06-共享记录.doc",
        module="materials.n07.table_landing_share_record",
        function="build_table_landing_share_record_document",
    ),
    "06-共享记录": MaterialSpec(
        default_filename="06-共享记录.doc",
        module="materials.n07.table_landing_share_record",
        function="build_table_landing_share_record_document",
    ),
}


def get_material_spec(material_type):
    return SPECS.get(material_type)


def load_material_builder(spec):
    return getattr(import_module(spec.module), spec.function)

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
}


def get_material_spec(material_type):
    return SPECS.get(material_type)


def load_material_builder(spec):
    return getattr(import_module(spec.module), spec.function)

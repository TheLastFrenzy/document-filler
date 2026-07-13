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
}


def get_material_spec(material_type):
    return SPECS.get(material_type)


def load_material_builder(spec):
    return getattr(import_module(spec.module), spec.function)

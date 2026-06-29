import re
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class SkillMetadataTest(unittest.TestCase):
    def test_skill_frontmatter_only_declares_name_and_description(self):
        content = (ROOT / "SKILL.md").read_text(encoding="utf-8")
        match = re.match(r"^---\n(.*?)\n---", content, re.DOTALL)

        self.assertIsNotNone(match)
        frontmatter = match.group(1)
        top_level_keys = [
            line.split(":", 1)[0]
            for line in frontmatter.splitlines()
            if line and not line.startswith(" ")
        ]

        self.assertEqual(top_level_keys, ["name", "description"])
        self.assertIn("Word、PDF 或 Excel", frontmatter)

    def test_openai_yaml_matches_current_material_scope(self):
        content = (ROOT / "agents" / "openai.yaml").read_text(encoding="utf-8")

        self.assertIn('display_name: "Document Filler"', content)
        self.assertIn("Word/PDF/Excel", content)
        self.assertIn("Use $document-filler", content)
        self.assertNotIn("支持需求/设计文档", content)

    def test_requirements_include_pdf_and_workbook_dependencies(self):
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8").splitlines()

        for package in ["openpyxl", "python-docx", "pymupdf", "reportlab", "pypdf"]:
            self.assertIn(package, requirements)


if __name__ == "__main__":
    unittest.main()

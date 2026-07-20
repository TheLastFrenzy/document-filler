# Document Filler Behavior-Preserving Optimization Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Consolidate duplicated Word infrastructure and material metadata while preserving every existing document-generation behavior.

**Architecture:** Add one public shared Office helper module and retain compatibility wrappers in existing material modules. Enrich registry query functions without moving N08 generation code or changing builder inputs and outputs.

**Tech Stack:** Python 3, unittest, python-docx, PowerShell Word COM, openpyxl.

---

### Task 1: Lock Shared Office Behavior

**Files:**
- Create: `tests/test_shared_office_word.py`
- Create: `scripts/materials/shared/office_word.py`
- Modify: `scripts/materials/n07/api_test_report.py`
- Modify: `scripts/materials/n07/api_launch_record.py`
- Modify: `scripts/materials/n07/table_landing_design.py`
- Modify: `scripts/materials/n07/table_landing_test_report.py`
- Modify: `scripts/materials/n07/table_landing_launch_record.py`
- Modify: `scripts/materials/n07/table_landing_share_record.py`

- [ ] Write tests that import the new public helper functions and assert DOCX passthrough, unsupported-extension errors, conversion output checks, callback-based TOC updates, and compatibility wrapper availability.
- [ ] Run `python -m unittest tests.test_shared_office_word` and verify it fails because `materials.shared.office_word` does not exist.
- [ ] Implement the shared functions by moving the existing PowerShell and save logic without changing commands, format codes, timeouts, or file checks.
- [ ] Replace cross-material private imports with public shared imports while retaining existing private wrapper names in API modules.
- [ ] Run `python -m unittest tests.test_shared_office_word tests.test_api_requirement tests.test_n07_table_landing_requirement` and expect all tests to pass.

### Task 2: Centralize Material Type Queries

**Files:**
- Modify: `tests/test_cli_contract.py`
- Modify: `scripts/materials/registry.py`
- Modify: `scripts/fill_document.py`

- [ ] Add failing tests for canonical registered types, alias resolution, and generated unsupported-material messages.
- [ ] Run `python -m unittest tests.test_cli_contract` and verify the new assertions fail because the registry query functions do not exist.
- [ ] Add registry query functions and replace the hard-coded CLI support string while preserving dispatch and filenames.
- [ ] Run `python -m unittest tests.test_cli_contract tests.test_api_requirement tests.test_n07_table_landing_requirement` and expect all tests to pass.

### Task 3: Remove the Unused Dependency

**Files:**
- Modify: `tests/test_skill_metadata.py`
- Modify: `requirements.txt`

- [ ] Change the dependency test to assert the runtime dependency set and explicitly reject unused `pypdf`.
- [ ] Run `python -m unittest tests.test_skill_metadata` and verify it fails while `pypdf` remains listed.
- [ ] Remove only `pypdf` from `requirements.txt`; keep `reportlab` because DOCX generation currently imports the PDF helper.
- [ ] Run `python -m unittest tests.test_skill_metadata` and expect it to pass.

### Task 4: Full Verification

**Files:**
- Verify all modified files.

- [ ] Run `python -m unittest discover -s tests` and require zero failures.
- [ ] Run `python -m compileall -q scripts tests` and require exit code 0.
- [ ] Run `python scripts/fill_document.py --help` and verify the existing CLI arguments remain present.
- [ ] Run `python -m pip check` and require no broken requirements.
- [ ] Run `git diff --check` and inspect `git diff` for any business-rule or generated-content changes.

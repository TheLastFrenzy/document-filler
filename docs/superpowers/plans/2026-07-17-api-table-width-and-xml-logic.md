# API Table Width And XML Logic Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make N07 API input tables full-width and generate concrete, XML-grounded N08 statistics logic descriptions.

**Architecture:** Keep the N07 layout correction local to the API test-report builder. Extend the existing dependency-free N08 SQL heuristics into structured node analysis, then render bounded descriptions from extracted facts so generation remains deterministic.

**Tech Stack:** Python 3, `python-docx`, `openpyxl`, `xml.etree.ElementTree`, regular expressions, `unittest`.

---

### Task 1: N07 API input-table width

**Files:**
- Modify: `tests/test_api_requirement.py`
- Modify: `scripts/materials/n07/api_test_report.py`

- [ ] Add a regression assertion that a deliberately narrow input prototype and wider output prototype produce equal generated widths.
- [ ] Run `python -m unittest tests.test_api_requirement.ApiRequirementTest.test_build_api_test_report_document_replaces_list_params_and_result_images -v` and confirm the width assertion fails.
- [ ] Add a local helper that copies target table width/indentation, sets fixed layout, scales `tblGrid`, and updates unmerged `tcW` values.
- [ ] Apply the helper only to cloned API test-report input tables.
- [ ] Re-run the focused test and confirm it passes.

### Task 2: N08 concrete XML logic descriptions

**Files:**
- Modify: `tests/test_stats_result_dispatch.py`
- Modify: `scripts/build_stats_result_usage_workbook.py`

- [ ] Add failing tests for natural-language `--` comments, concrete join field pairs, `INSERT ... SELECT` field mappings, latest-snapshot filters, and evidence-based database terminology.
- [ ] Add a negative test proving plain SQL does not acquire window-function, set-union, or record-linkage terminology.
- [ ] Run the focused tests and confirm failures come from missing concrete analysis.
- [ ] Implement top-level SQL list splitting, insert/select mapping extraction, join-condition extraction, comment filtering, operation detection, and bounded rendering.
- [ ] Replace generic calculation hints in `build_business_logic_steps` with the extracted node facts while retaining source/result context.
- [ ] Re-run focused N08 tests and confirm they pass.

### Task 3: Regression and regeneration

**Files:**
- Regenerate: `C:/Users/p/Desktop/回归测试20260717/结果/N07-API接口开发/04-接口测试报告（含《API接口列表》）.docx`
- Regenerate: `C:/Users/p/Desktop/回归测试20260717/结果/N08-数据统计分析/04-数据统计分析_结果表及使用说明.xlsx`
- Regenerate: `C:/Users/p/Desktop/回归测试20260717/结果/N08-数据统计分析/01-数据统计分析_需求文档.docx`
- Regenerate: `C:/Users/p/Desktop/回归测试20260717/结果/N08-数据统计分析/02-数据统计分析_设计文档.docx`

- [ ] Run the complete unit-test suite and compile checks.
- [ ] Create a temporary ZIP-level ledger header alias without saving through `openpyxl`; verify all OLE anchors and payload hashes are unchanged.
- [ ] Regenerate N07 API 04.
- [ ] Regenerate N08 statistics in order `04 -> 01 -> 02`.
- [ ] Verify Office ZIP integrity, equal N07 parameter-table widths, concrete mapping text, and unchanged N08 result/flowchart/media counts.
- [ ] Remove the temporary ledger and generated flowchart scratch directory.
- [ ] Confirm the result tree contains only service folders and deliverables.

Commits are intentionally omitted: the confirmed workspace is already dirty on `main`, and the user requested local optimization and regeneration rather than repository publication.

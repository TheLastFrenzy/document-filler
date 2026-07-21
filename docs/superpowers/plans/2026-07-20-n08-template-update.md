# N08 Template Update Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Update three N08 document generators to match the 20260721 templates and regenerate the corresponding deliverables without changing unrelated generation behavior.

**Architecture:** Keep the existing `fill_document.py` dispatch and shared ledger normalization. Add small heading/table-location helpers where needed, make each affected builder consume the new template anchors, and preserve existing catalog, relation-workbook, image, and TOC paths. The external desktop templates remain inputs; code and regression tests remain the repository deliverable.

**Tech Stack:** Python 3, `python-docx`, `openpyxl`, existing Word COM/Office helpers, `unittest`.

---

### Task 1: Commit the approved design and establish a clean baseline

**Files:**
- Create: `docs/superpowers/specs/2026-07-20-n08-template-update-design.md`
- Create: `docs/superpowers/plans/2026-07-20-n08-template-update.md`

- [x] **Step 1: Create and self-review design and plan.** Confirm the plan covers report-design source titles, removal of indicator tables, the two stats requirement subtables, renamed requirement labels, stats-design source columns, and the inner “数据结果表设计” heading.
- [x] **Step 2: Run baseline tests.**

Run: `python -m unittest discover -s tests`

Expected: `Ran 130 tests` and `OK` on `origin/main`.

- [ ] **Step 3: Commit the design documents.**

Run: `git add docs/superpowers/specs/2026-07-20-n08-template-update-design.md docs/superpowers/plans/2026-07-20-n08-template-update.md; git commit -m "docs: plan N08 template update"`

### Task 2: Add failing regression tests for the report design changes

**Files:**
- Modify: `tests/test_data_report_regressions.py`
- Modify: `scripts/fill_document.py`

- [ ] **Step 1: Write the failing assertions.** Add a fixture template/data row that calls `fill_design_doc_full` and assert the generated document has:

```python
headings = [p.text for p in output.paragraphs if p.style.name == "Heading 4"]
assert "目录中文名" in headings
assert "目录英文编码" not in headings
assert all("报表指标设计" not in p.text for p in output.paragraphs)
assert output.tables[1].cell(0, 0).text.startswith("目录英文编码 目录中文名")
```

- [ ] **Step 2: Run only the new test.**

Run: `python -m unittest tests.test_data_report_regressions.ReportDesignTemplateUpdateTest`

Expected: FAIL because the current builder uses the resource code as the Heading 4 title and appends the indicator table.

### Task 3: Implement report design template behavior

**Files:**
- Modify: `scripts/fill_document.py:fill_design_doc_full`
- Modify: `references/filling_rules.md`
- Modify: `SKILL.md`

- [ ] **Step 1: Change source-table titles.** In the loop over `codes`, compute `resource_title = resource_name or code` and `table_title = f"{resource_code} {resource_title}"` when both are available. Append `resource_title` as the Heading 4 and use `table_title` for `mk_ds_table`.
- [ ] **Step 2: Stop appending indicator content.** Keep attachment parsing because it still supports text normalization and result-form inference, but remove the `报表指标设计` heading and `mk_indicator_table(...)` append from this builder.
- [ ] **Step 3: Run the focused test and full affected regression module.**

Run: `python -m unittest tests.test_data_report_regressions.ReportDesignTemplateUpdateTest tests.test_data_report_regressions.DataReportRegressionTest`

Expected: PASS with the new assertions and all existing data-report regression tests green.

### Task 4: Add failing regression tests for stats requirement structure

**Files:**
- Modify: `tests/test_stats_requirement_split_rows.py`
- Modify: `scripts/fill_document.py`

- [ ] **Step 1: Assert the new two-table structure.** Generate from the existing split-row fixture and assert:

```python
assert "需求单、工单产出对应列表" in headings
assert "工单与任务的对应关系" in headings
assert source_table.rows[0].cells[3].text == "工单标题"
assert task_table.rows[0].cells == ["序号", "对应需求单编号", "对应工单编号", "任务中文名"]
assert task_table.rows[1].cells[3].text == "结果表一"
assert "数据加工要求" in headings
assert "数据量对后续运维的特殊要求" not in full_text
```

- [ ] **Step 2: Run the focused test before implementation.**

Run: `python -m unittest tests.test_stats_requirement_split_rows.StatsRequirementSplitRowsTest`

Expected: FAIL because the current builder fills only the first table, does not populate the task table, and keeps the old section labels.

### Task 5: Implement stats requirement headings and tables

**Files:**
- Modify: `scripts/fill_document.py:fill_stats_requirement_doc`
- Modify: `references/filling_rules.md`
- Modify: `SKILL.md`

- [ ] **Step 1: Locate both tables by the new Heading 3 anchors.** Add a helper that returns the first `w:tbl` after a heading paragraph and use it for `需求单、工单产出对应列表` and `工单与任务的对应关系`; raise `ValueError` naming the missing heading.
- [ ] **Step 2: Fill the source table and task table.** Keep the old five-column source-table rows and total. Build one task row per `(work_order, result_cn, result_en)` in stable ledger order with `[index, request_no, work_order_no, result_cn]`.
- [ ] **Step 3: Update the description and requirement labels.** Use `服务周期内，共有{ureq}张需求单，{ugd}张工单涉及{trep}次数据统计分析` and emit Heading 3 `数据加工要求`; change only the third body-label text to `对运维的工作要求` while reading the same source field.
- [ ] **Step 4: Run focused tests and verify old generated content is removed.**

Run: `python -m unittest tests.test_stats_requirement_split_rows.StatsRequirementSplitRowsTest tests.test_stats_test_doc_dispatch.StatsTestDocDispatchTest`

Expected: PASS and no old “数据加工周期” or “数据量对后续运维的特殊要求” in generated output.

### Task 6: Add failing regression tests for stats design structure

**Files:**
- Modify: `tests/test_stats_design_dispatch.py`
- Modify: `scripts/fill_document.py`

- [ ] **Step 1: Add assertions against a multi-program fixture.** Assert the source table headers are exactly `序号/需求单编号/工单编号/程序名称`, the row uses the Chinese result name, and each generated program section contains `数据结果表设计` but not `数据统计分析设计` as its inner Heading 3.
- [ ] **Step 2: Run the focused test before implementation.**

Run: `python -m unittest tests.test_stats_design_dispatch.StatsDesignDispatchTest`

Expected: FAIL because the current builder writes five source columns and emits the old inner heading.

### Task 7: Implement stats design source table and result heading

**Files:**
- Modify: `scripts/fill_document.py:fill_stats_design_doc`
- Modify: `references/filling_rules.md`
- Modify: `SKILL.md`

- [ ] **Step 1: Fill the new four-column source table.** Locate the table after `需求来源` and append `[序号, 需求单号, 工单号, result_cn]`; remove the old work-order title and English program columns.
- [ ] **Step 2: Emit `数据结果表设计`.** Keep the H1 truncation anchor compatible with both old and new template text, but append the per-program H3 as `数据结果表设计`.
- [ ] **Step 3: Run the focused regression tests.**

Run: `python -m unittest tests.test_stats_design_dispatch.StatsDesignDispatchTest tests.test_stats_requirement_split_rows.StatsRequirementSplitRowsTest`

Expected: PASS.

### Task 8: Regenerate and verify the three real deliverables

**Inputs:**
- `C:\Users\p\Desktop\新分支20260721\台账清单_20260720.xlsx`
- `C:\Users\p\Desktop\新分支20260721\数据目录数据.xlsx`
- Templates under `C:\Users\p\Desktop\新分支20260721\模版\N08-数据报表服务` and `...\N08-数据统计分析`

**Outputs:**
- `C:\Users\p\Desktop\新分支20260721\结果\N08-数据报表服务\02-数据报表_设计文档.docx`
- `C:\Users\p\Desktop\新分支20260721\结果\N08-数据统计分析\01-数据统计分析_需求文档.docx`
- `C:\Users\p\Desktop\新分支20260721\结果\N08-数据统计分析\02-数据统计分析_设计文档.docx`

- [ ] **Step 1: Generate report design.** Run `python scripts/fill_document.py --service-dir "N08-数据报表服务" --material-type "02-数据报表_设计文档" --excel "C:\Users\p\Desktop\新分支20260721\台账清单_20260720.xlsx" --template "C:\Users\p\Desktop\新分支20260721\模版\N08-数据报表服务\02-数据报表_设计文档.docx" --catalog "C:\Users\p\Desktop\新分支20260721\数据目录数据.xlsx" --output "C:\Users\p\Desktop\新分支20260721\结果\N08-数据报表服务"`.
- [ ] **Step 2: Generate stats requirement.** Run the equivalent command for `N08-数据统计分析 / 01-数据统计分析_需求文档` with its new template and output folder.
- [ ] **Step 3: Generate stats design.** Run the equivalent command for `N08-数据统计分析 / 02-数据统计分析_设计文档` and pass the same catalog path.
- [ ] **Step 4: Inspect real outputs.** Open each Word file, count expected headings/table rows, assert no removed headings or placeholders, run `validate.py`, export PDFs, and inspect rendered pages for clipping/overlap/blank pages.

### Task 9: Final verification and branch handoff

**Files:**
- Modify: `SKILL.md`, `references/filling_rules.md`, affected tests and `scripts/fill_document.py` as required by preceding tasks.

- [ ] **Step 1: Run all verification commands.**

Run: `python -m unittest discover -s tests`; `python -m compileall -q scripts tests`; `python -m pip check`; `python F:\Codex\skills\.system\skill-creator\scripts\quick_validate.py F:\Codex\skills\document-filler\.worktrees\n08-template-update-20260721`; `git diff --check`.

- [ ] **Step 2: Check branch scope.** Confirm only the new branch contains these changes, the original `main` worktree still has its pre-existing `SKILL.md` state, and generated outputs are outside the repository.
- [ ] **Step 3: Commit implementation and report branch/output paths.** Use `git status`, `git diff --stat`, and a focused commit message before handing the branch back.

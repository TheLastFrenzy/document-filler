# N07 Table Landing Requirement Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add `01-库表落地方式_需求文档` generation for service directory `N07-库表落地方式`.

**Architecture:** Follow the existing modular N07 pattern. Add a shared ledger/attachment parser for table-landing work orders, then add a focused Word builder that clones the real template's paragraph and table prototypes.

**Tech Stack:** Python, openpyxl, python-docx, ZIP/XLSX package inspection, existing `materials.shared.word_sections` helpers, unittest.

---

### Task 1: Failing Tests

**Files:**
- Create: `tests/test_n07_table_landing_requirement.py`

- [ ] **Step 1: Write failing unit tests**

Add tests that create a minimal ledger with merged rows, two embedded workbook attachments, and a Word template matching `业务场景` / `需求说明`.

Expected behaviors:
- `resolve_output_path(temp_dir, "01-库表落地方式_需求文档")` returns `01-需求文档.docx`.
- `read_table_landing_work_orders()` groups merged rows by `工单号`, keeps program counts once, and reads landing rows from first-sheet workbook columns `落地库名`, `落地表名`, `表中文名称`.
- `build_table_landing_requirement_document()` replaces the template demand source table and per-work-order detail tables.

- [ ] **Step 2: Run tests and verify RED**

Run:

```powershell
python -m unittest tests.test_n07_table_landing_requirement
```

Expected: FAIL because `materials.n07.table_landing_requirement` and registry support do not exist.

### Task 2: Shared Table-Landing Parser

**Files:**
- Create: `scripts/materials/n07/table_landing.py`

- [ ] **Step 1: Implement minimal data classes**

Define:
- `TableLandingTask`
- `TableLandingWorkOrder`

Fields should cover demand number, work order number, title, description, program count, source rows, result table values, target user, update cycle, update requirement, attachment path, and parsed task rows.

- [ ] **Step 2: Read merged ledger rows**

Use openpyxl with merged-cell inheritance. Required headers:
- `服务目录`
- `需求单号`
- `工单号`
- `工单标题`
- `程序数`
- `工单描述`
- `结果表清单`
- `数据统计分析执行周期`
- `数据更新要求`
- `自测报告附件`
- `下发前置机中文名`

Group by `工单号`, collect each source row's `结果表清单`, and only count `程序数` once per work order.

- [ ] **Step 3: Extract embedded workbook attachments**

Reuse the existing VML/OLE anchor reader shape where possible. Accept direct embedded `.xlsx` packages and OLE Package payloads that contain an XLSX ZIP. Match anchors to work orders by source row and `自测报告附件` column.

- [ ] **Step 4: Parse first worksheet task rows**

Read each extracted workbook's first sheet. For each data row, map:
- `落地库名` -> landing database
- `落地表名` -> landing table
- `表中文名称` -> business scene

Raise clear errors when required workbook columns or attachments are missing.

### Task 3: Word Builder

**Files:**
- Create: `scripts/materials/n07/table_landing_requirement.py`

- [ ] **Step 1: Capture template prototypes**

Open the template with python-docx. Find:
- `Heading 2` exact text `业务场景`
- `Heading 2` exact text `需求说明`

Use the first table between these headings as the demand source table prototype. Use the first work-order heading and first six-column detail table after `需求说明` as prototypes for generated sections.

- [ ] **Step 2: Build business scene content**

Under `业务场景`, preserve existing non-table paragraphs and replace the old table with:

```text
服务周期内，共有{需求单数}张需求单，{工单数}张工单涉及{程序数合计}个库表落地共享任务。具体需求单、工单和产出如下表：
```

Then clone the 5-column table with rows:

```text
序号 / 对应需求单编号 / 对应工单编号 / 工单内容 / 涉及共享任务数量（个）
```

- [ ] **Step 3: Build requirement sections**

For each work order, generate:
- `Heading 2`: `{工单号}_{工单标题}`
- `Normal`: `需求口径：{工单描述}`
- `Normal`: `涉及共享任务数量：{程序数}`
- 6-column table: `序号 / 落地库 / 落地表 / 对象用户 / 库表来源 / 业务场景`
- `Normal`: `计划供数方式：库表下发`
- `Normal`: `计划更新频率：{数据统计分析执行周期}`
- `Normal`: `更新时间：{数据更新要求}`

Use task rows from the embedded workbook and source table names from the ledger's `结果表清单`, aligned by row index.

- [ ] **Step 4: Save and update TOC**

Create the output directory, save `.docx`, and call `update_toc_via_com(output)`.

### Task 4: Registration And Docs

**Files:**
- Modify: `scripts/materials/registry.py`
- Modify: `SKILL.md`
- Modify: `references/filling_rules.md`

- [ ] **Step 1: Register material**

Add aliases:
- `01-库表落地方式_需求文档`
- `01-需求文档`

Default filename: `01-需求文档.docx`.

- [ ] **Step 2: Document usage**

Add user-facing support text for `N07-库表落地方式` and its `01-库表落地方式_需求文档` material, including the need for workbook attachments in `自测报告附件`.

### Task 5: Verification

**Files:**
- No new files unless tests reveal a focused fix is needed.

- [ ] **Step 1: Run targeted unit tests**

```powershell
python -m unittest tests.test_n07_table_landing_requirement tests.test_cli_contract tests.test_skill_metadata
```

Expected: all tests pass.

- [ ] **Step 2: Run real-file smoke generation**

```powershell
python scripts/fill_document.py --service-dir "N07-库表落地方式" --material-type "01-库表落地方式_需求文档" --excel "C:\Users\p\Desktop\新调整\验收台账清单（更新）.xlsx" --template "C:\Users\p\Desktop\新调整\N07 数据共享、开发、授权订阅服务\02-库表落地方式\01-需求文档 .docx" --output "<temp>\01-需求文档.docx"
```

Expected: command exits 0 and output `.docx` opens through python-docx with generated work-order headings and detail tables.

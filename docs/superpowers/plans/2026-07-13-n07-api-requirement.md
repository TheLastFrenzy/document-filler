# N07 API Requirement Document Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a modular `01-API接口开发_需求文档` material that generates `C:\Users\p\Desktop\新调整\结果\01-需求文档.docx` from the real N07 ledger, template, and embedded self-test report.

**Architecture:** Keep legacy material branches unchanged. Add a lazy material registry for new materials, shared modules for merged-ledger reading, embedded DOCX extraction, and Word section operations, plus one N07-specific builder. The N07 builder owns all API requirement business rules and exposes `build_api_requirement_document(excel_path, service_dir, template_path, output_path)`.

**Tech Stack:** Python 3, openpyxl, olefile, python-docx, OOXML, unittest, Microsoft Word COM, PyMuPDF.

---

### Task 1: New-material registry and output contract

**Files:**
- Create: `scripts/materials/__init__.py`
- Create: `scripts/materials/registry.py`
- Modify: `scripts/fill_document.py:48-67`
- Test: `tests/test_api_requirement.py`

- [ ] **Step 1: Write the failing registry tests**

```python
def test_api_requirement_registration_uses_public_output_filename(self):
    module = load_fill_document_module()
    with tempfile.TemporaryDirectory() as temp_dir:
        output = module.resolve_output_path(temp_dir, "01-API接口开发_需求文档")
    self.assertEqual(Path(output).name, "01-需求文档.docx")

def test_unknown_material_keeps_legacy_output_behavior(self):
    module = load_fill_document_module()
    with tempfile.TemporaryDirectory() as temp_dir:
        output = module.resolve_output_path(temp_dir, "01-数据报表_需求文档")
    self.assertEqual(Path(output).name, "01-数据报表_需求文档.docx")
```

- [ ] **Step 2: Run the registry tests and verify RED**

Run: `python -m unittest tests.test_api_requirement.ApiRequirementTest.test_api_requirement_registration_uses_public_output_filename`

Expected: FAIL because the output name is `01-API接口开发_需求文档.docx` or the material is not registered.

- [ ] **Step 3: Implement the registry**

```python
# scripts/materials/registry.py
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
```

Update `resolve_output_path` so a directory first checks `get_material_spec(material_type).default_filename`, then falls back to the existing `MATERIAL_OUTPUT_EXTENSIONS` behavior.

- [ ] **Step 4: Run the registry tests and verify GREEN**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: registry tests PASS; builder-import tests remain pending until the builder exists.

- [ ] **Step 5: Commit registry work**

```powershell
git add scripts/materials/__init__.py scripts/materials/registry.py scripts/fill_document.py tests/test_api_requirement.py
git commit -m "添加可扩展材料注册表"
```

### Task 2: Merged ledger work-order reader

**Files:**
- Create: `scripts/materials/shared/__init__.py`
- Create: `scripts/materials/shared/ledger.py`
- Test: `tests/test_api_requirement.py`

- [ ] **Step 1: Write failing merged-row tests**

```python
def test_read_api_work_orders_groups_merged_rows_and_counts_programs_once(self):
    ledger = make_api_ledger(self.temp_path / "ledger.xlsx")
    orders = read_api_work_orders(ledger, "N07-API接口开发")
    self.assertEqual(len(orders), 1)
    self.assertEqual(orders[0].program_count, 3)
    self.assertEqual(orders[0].source_rows, (2, 3, 4))
    self.assertEqual(
        [item.chinese_name for item in orders[0].interfaces],
        ["境外人员获取令牌接口", "境外人员身份认证申请接口", "境外人员身份认证请求接口"],
    )
```

The fixture creates rows 2-4, merges the work-order columns, stores `程序数=3` only in the merged top-left cell, and writes one `结果表清单` value per row.

- [ ] **Step 2: Run the ledger test and verify RED**

Run: `python -m unittest tests.test_api_requirement.ApiRequirementTest.test_read_api_work_orders_groups_merged_rows_and_counts_programs_once`

Expected: ERROR because `materials.shared.ledger` does not exist.

- [ ] **Step 3: Implement ledger models and grouping**

```python
@dataclass
class ApiInterface:
    chinese_name: str
    english_name: str
    source_row: int
    input_groups: list = field(default_factory=list)
    output_groups: list = field(default_factory=list)
    purpose: str = ""

@dataclass
class ApiWorkOrder:
    demand_no: str
    work_order_no: str
    title: str
    description: str
    program_count: int
    source_rows: tuple[int, ...]
    interfaces: list[ApiInterface]
    self_report_path: Path | None = None
```

Implement `merged_value_getter(ws)`, `parse_interface_name(value)`, and `read_api_work_orders(excel_path, service_dir)`. Validate the required columns `服务目录`, `需求单号`, `工单号`, `工单标题`, `程序数`, `工单描述`, `结果表清单`, and `自测报告附件`. Group by work-order number and only parse `程序数` from the first grouped record.

- [ ] **Step 4: Run ledger tests and verify GREEN**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: all ledger tests PASS.

- [ ] **Step 5: Commit ledger work**

```powershell
git add scripts/materials/shared/__init__.py scripts/materials/shared/ledger.py tests/test_api_requirement.py
git commit -m "支持N07合并行台账聚合"
```

### Task 3: Embedded DOCX extraction and API parameter parser

**Files:**
- Create: `scripts/materials/shared/embedded_docx.py`
- Create: `scripts/materials/n07/__init__.py`
- Create: `scripts/materials/n07/api_requirement.py`
- Test: `tests/test_api_requirement.py`

- [ ] **Step 1: Write failing parser tests**

```python
def test_parse_self_report_preserves_input_groups_and_output_tables(self):
    report = make_self_report(self.temp_path / "self-test.docx")
    interfaces = [ApiInterface("身份认证申请接口", "api_apply", 2)]
    parsed = parse_api_report(report, interfaces)
    item = parsed[0]
    self.assertEqual([g.label for g in item.input_groups], ["请求头", "请求体"])
    self.assertEqual(item.input_groups[0].headers, ["参数", "参数说明", "是否必选"])
    self.assertEqual(item.output_groups[0].rows[0][0], "bizSerialNum")
```

```python
def test_extract_embedded_docx_maps_attachment_anchor_to_merged_work_order(self):
    reports = extract_embedded_docx_by_work_order(ledger, orders, "自测报告附件", temp_dir)
    self.assertEqual(set(reports), {"WO-1"})
    self.assertTrue(reports["WO-1"].read_bytes().startswith(b"PK\x03\x04"))
```

- [ ] **Step 2: Run extraction/parser tests and verify RED**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: ERROR because extraction and parsing functions are undefined.

- [ ] **Step 3: Implement generic OLE DOCX extraction**

Implement `package_stream_from_ole` with this exact fallback order:

```python
def package_stream_from_ole(data: bytes) -> bytes:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".bin") as temp_file:
        temp_file.write(data)
        temp_path = Path(temp_file.name)
    ole = None
    try:
        ole = olefile.OleFileIO(str(temp_path))
        streams = {
            "/".join(parts).lower(): "/".join(parts)
            for parts in ole.listdir(streams=True, storages=False)
        }
        if "package" in streams:
            payload = ole.openstream(streams["package"]).read()
        else:
            native_name = next(
                (name for key, name in streams.items() if "ole10native" in key), ""
            )
            if not native_name:
                raise ValueError("OLE对象缺少Package或Ole10Native流")
            native = ole.openstream(native_name).read()
            offset = native.find(b"PK\x03\x04")
            if offset < 0:
                raise ValueError("Ole10Native流中未找到DOCX载荷")
            payload = native[offset:]
        if not payload.startswith(b"PK\x03\x04"):
            raise ValueError("内嵌附件不是DOCX文件")
        return payload
    finally:
        if ole is not None:
            ole.close()
        temp_path.unlink(missing_ok=True)
```

Implement the extraction entry point with this public contract:

```python
def extract_embedded_docx_by_work_order(
    excel_path, work_orders, attachment_header, work_dir
) -> dict[str, Path]:
    attachment_col = find_header_column(excel_path, attachment_header)
    reports = {}
    for anchor in read_ole_anchors(excel_path):
        if anchor.column != attachment_col:
            continue
        order = next(
            (item for item in work_orders if anchor.row in item.source_rows),
            None,
        )
        if order is None:
            continue
        payload = package_stream_from_ole(anchor.payload)
        path = Path(work_dir) / f"{order.work_order_no}_self_report.docx"
        path.write_bytes(payload)
        reports[order.work_order_no] = path
    missing = [item.work_order_no for item in work_orders if item.work_order_no not in reports]
    if missing:
        raise ValueError(f"以下工单缺少可解析的自测报告附件: {', '.join(missing)}")
    return reports
```

`find_header_column` reads row 1 and returns a one-based index. `read_ole_anchors` returns `OleAnchor(row, column, payload)` values after resolving worksheet relationships, VML `x:Anchor` coordinates, and `xl/embeddings` targets. Errors include the work-order number and the missing or malformed attachment reason.

- [ ] **Step 4: Implement report block parsing**

```python
@dataclass
class ParameterGroup:
    label: str
    headers: list[str]
    rows: list[list[str]]

def parse_api_report(report_path: Path, interfaces: list[ApiInterface]) -> list[ApiInterface]:
    blocks = read_docx_blocks(report_path)
    test_start = find_text_block(blocks, "测试内容")
    positions = match_interface_positions(blocks[test_start + 1 :], interfaces)
    for index, interface in enumerate(interfaces):
        start = positions[interface.chinese_name]
        end = positions[interfaces[index + 1].chinese_name] if index + 1 < len(interfaces) else len(blocks)
        section = blocks[start:end]
        input_index = find_text_block(section, "输入参数")
        output_index = find_text_block(section, "输出参数")
        interface.input_groups = parameter_groups(section[input_index + 1 : output_index])
        interface.output_groups = parameter_groups(section[output_index + 1 :])
        if not interface.input_groups or not interface.output_groups:
            raise ValueError(f"接口{interface.chinese_name}缺少输入或输出参数表")
    return interfaces
```

`read_docx_blocks` returns paragraph text and complete table matrices in document order. `match_interface_positions` requires exact normalized Chinese-name matches and raises with the unmatched interface names. `parameter_groups` attaches the nearest preceding non-empty paragraph label such as `请求头：`, `请求体：`, or `IdAuthData：` to each table; it uses an empty label for a single unlabeled table and drops trailing “测试结果” content.

- [ ] **Step 5: Run extraction/parser tests and verify GREEN**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: extraction and parsing tests PASS.

- [ ] **Step 6: Commit extraction/parser work**

```powershell
git add scripts/materials/shared/embedded_docx.py scripts/materials/n07/__init__.py scripts/materials/n07/api_requirement.py tests/test_api_requirement.py
git commit -m "解析API自测报告参数"
```

### Task 4: Word template transformation and humanized purpose text

**Files:**
- Create: `scripts/materials/shared/word_sections.py`
- Modify: `scripts/materials/n07/api_requirement.py`
- Test: `tests/test_api_requirement.py`

- [ ] **Step 1: Write failing document-content tests**

```python
def test_build_api_requirement_document_replaces_dynamic_sections_and_preserves_static_content(self):
    output = self.temp_path / "01-需求文档.docx"
    build_api_requirement_document(ledger, "N07-API接口开发", template, output)
    doc = Document(output)
    text = "\n".join(p.text for p in doc.paragraphs)
    self.assertIn("共提供3个API接口服务", text)
    self.assertIn("WO-1_市公安局-出入境证件身份认证-共享接口", text)
    self.assertIn("需求口径：", text)
    self.assertIn("共享开放方案", text)
    self.assertIn("API接口对外服务", text)
    self.assertNotIn("企业电子票据证件号码校验接口", text)
```

```python
def test_purpose_text_is_grounded_and_avoids_banned_ai_phrases(self):
    purpose = build_interface_purpose(order, interface)
    self.assertIn("身份认证", purpose)
    self.assertNotIn("赋能", purpose)
    self.assertNotIn("至关重要", purpose)
    self.assertNotIn("确保", purpose)
```

- [ ] **Step 2: Run document tests and verify RED**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: FAIL because the builder does not transform the template.

- [ ] **Step 3: Implement Word section helpers**

Implement:

```python
def find_heading(doc, style_name, exact_text):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == style_name and paragraph.text.strip() == exact_text:
            return paragraph._p
    raise ValueError(f"模板中未找到{exact_text}章节")

def elements_between(start, end):
    values = []
    current = start.getnext()
    while current is not None and current is not end:
        values.append(current)
        current = current.getnext()
    return values

def replace_between(start, end, elements):
    parent = start.getparent()
    for current in elements_between(start, end):
        parent.remove(current)
    anchor = start
    for element in elements:
        anchor.addnext(element)
        anchor = element

def paragraph_element(text, style_id, first_line_indent=None):
    paragraph = OxmlElement("w:p")
    properties = OxmlElement("w:pPr")
    style = OxmlElement("w:pStyle")
    style.set(qn("w:val"), style_id)
    properties.append(style)
    if first_line_indent is not None:
        indent = OxmlElement("w:ind")
        indent.set(qn("w:firstLine"), str(first_line_indent))
        properties.append(indent)
    paragraph.append(properties)
    run = OxmlElement("w:r")
    node = OxmlElement("w:t")
    node.text = text
    run.append(node)
    paragraph.append(run)
    return paragraph

def table_element(headers, rows, template_table=None):
    table = deepcopy(template_table._tbl) if template_table is not None else OxmlElement("w:tbl")
    for row in list(table.findall(qn("w:tr"))):
        table.remove(row)
    for values, is_header in [(headers, True), *[(row, False) for row in rows]]:
        table_row = OxmlElement("w:tr")
        for value in values:
            table_row.append(cell_element(value, bold=is_header, shaded=is_header))
        table.append(table_row)
    set_fixed_table_grid(table, len(headers))
    return table
```

The helpers preserve all document nodes outside replacement ranges and generate fixed-layout tables with copied borders, header shading, fonts, alignment, and cell margins.

- [ ] **Step 4: Implement API purpose generation**

`build_interface_purpose(order, interface)` uses interface semantics and parsed fields:

- 令牌: describe credential acquisition and later calls.
- 申请: describe submitted authentication context and returned serial/random data.
- 请求/查询: describe conditions received and verification/query results returned.
- fallback: describe the concrete input/output flow.

Normalize the result to one natural paragraph and reject the phrases `赋能`, `彰显`, `至关重要`, `确保`, and `重要支撑`.

- [ ] **Step 5: Implement all template transformations**

`build_api_requirement_document(excel_path, service_dir, template_path, output_path)` performs:

1. Read grouped work orders.
2. Extract and parse self-test DOCX files.
3. Replace the final business-scene paragraph.
4. Rebuild the demand list table and summary.
5. Replace sample work-order sections before `共享开放方案`.
6. Preserve `共享开放方案`.
7. Replace all sample API sections after `API接口清单`.
8. Save output, update TOC through the existing COM helper, and return the output path.

- [ ] **Step 6: Run document tests and verify GREEN**

Run: `python -m unittest tests.test_api_requirement -v`

Expected: all API requirement tests PASS.

- [ ] **Step 7: Commit Word builder work**

```powershell
git add scripts/materials/shared/word_sections.py scripts/materials/n07/api_requirement.py tests/test_api_requirement.py
git commit -m "生成N07 API需求文档"
```

### Task 5: Main dispatch and skill documentation

**Files:**
- Modify: `scripts/fill_document.py:3604-3642`
- Modify: `SKILL.md`
- Modify: `references/filling_rules.md`
- Modify: `agents/openai.yaml` only if its material summary is stale
- Test: `tests/test_api_requirement.py`
- Test: `tests/test_skill_metadata.py`

- [ ] **Step 1: Write a failing dispatch test**

```python
def test_fill_document_dispatches_registered_api_requirement_material(self):
    with mock.patch("materials.registry.load_material_builder") as load_builder:
        builder = load_builder.return_value
        builder.return_value = "out.docx"
        result = fill_document("ledger.xlsx", "N07-API接口开发", "01-API接口开发_需求文档", "template.docx", "out.docx")
    self.assertEqual(result, "out.docx")
    builder.assert_called_once_with(
        excel_path="ledger.xlsx",
        service_dir="N07-API接口开发",
        template_path="template.docx",
        output_path="out.docx",
    )
```

- [ ] **Step 2: Run the dispatch test and verify RED**

Run: `python -m unittest tests.test_api_requirement.ApiRequirementTest.test_fill_document_dispatches_registered_api_requirement_material`

Expected: FAIL because `fill_document` does not query the registry.

- [ ] **Step 3: Implement lazy registered-material dispatch**

At the top of `fill_document(excel_path, service_dir, material_type, template_path, output_path, catalog_path=None)`, after output resolution:

```python
spec = get_material_spec(material_type)
if spec:
    builder = load_material_builder(spec)
    return builder(
        excel_path=excel_path,
        service_dir=service_dir,
        template_path=template_path,
        output_path=output_path,
    )
```

Update unsupported-material text and CLI help to include the new type.

- [ ] **Step 4: Document the material**

Add the material to `SKILL.md` with its command example, required columns, attachment behavior, and default output file. Add a complete rule section to `references/filling_rules.md` covering business scene, demand list, work-order content, parameter extraction, purpose text, defaults, and strict errors.

- [ ] **Step 5: Run dispatch, metadata, and full regression tests**

Run: `python -m unittest tests.test_api_requirement tests.test_skill_metadata -v`

Run: `python -m unittest discover -s tests -p "test_*.py"`

Expected: all tests PASS.

- [ ] **Step 6: Commit dispatch and documentation**

```powershell
git add scripts/fill_document.py SKILL.md references/filling_rules.md agents/openai.yaml tests/test_api_requirement.py
git commit -m "接入API需求文档材料"
```

### Task 6: Real-file generation and visual verification

**Files:**
- Generate: `C:\Users\p\Desktop\新调整\结果\01-需求文档.docx`
- Generate temporarily: `C:\Users\p\Desktop\新调整\结果\01-需求文档.pdf`
- Generate temporarily: rendered page PNG files under a temporary verification directory

- [ ] **Step 1: Generate the real document**

Run:

```powershell
python scripts/fill_document.py `
  --service-dir "N07-API接口开发" `
  --material-type "01-API接口开发_需求文档" `
  --excel "C:\Users\p\Desktop\新调整\验收台账清单（更新）.xlsx" `
  --template "C:\Users\p\Desktop\新调整\N07 数据共享、开发、授权订阅服务\01-API接口开发\01-需求文档.docx" `
  --output "C:\Users\p\Desktop\新调整\结果"
```

Expected: creates `01-需求文档.docx` without overwriting the template.

- [ ] **Step 2: Verify document semantics**

Run a python-docx verifier that asserts:

- one demand and one work order;
- exactly three generated API Heading 3 sections;
- business scene contains the work-order title and API total 3;
- demand table total equals 3;
- every API has input and output tables;
- static `共享开放方案` text and its image relationship remain;
- sample API headings are absent;
- banned purpose phrases are absent.

Expected: verifier exits 0 and prints a JSON summary with three interfaces and no missing sections.

- [ ] **Step 3: Convert to PDF and render pages**

Use Word COM to save PDF, then use PyMuPDF to render every page at 150 DPI into a temporary verification directory.

Expected: PDF exists, page count is non-zero, and every rendered page has a non-white content bounding box.

- [ ] **Step 4: Inspect rendered pages**

Check contact sheets and full-resolution pages for:

- clipped or overlapping table text;
- tables wider than page margins;
- blank pages;
- headings stranded at page bottoms;
- missing static diagram;
- broken Chinese fonts.

If any issue is found, add a regression assertion where possible, fix the builder, regenerate, and repeat Steps 2-4.

- [ ] **Step 5: Run final verification**

Run: `python -m unittest discover -s tests -p "test_*.py"`

Run: `python F:\Codex\skills\.system\skill-creator\scripts\quick_validate.py F:\Codex\skills\document-filler` with `PYTHONUTF8=1`.

Run: `git status -sb` and `git diff --check`.

Expected: all tests pass, skill validation succeeds, no whitespace errors, and only intentional commits remain.

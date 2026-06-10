#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
验收文档自动填充脚本 - 支持多种材料类型
用法:
  01-需求文档:
    python fill_document.py --service-dir "N08-数据报表服务" --material-type "01-数据报表_需求文档" --excel "台账清单.xlsx" --template "模板.docx" --output "输出.docx"
  02-设计文档:
    python fill_document.py --service-dir "N08-数据报表服务" --material-type "02-数据报表_设计文档" --excel "台账清单.xlsx" --template "模板.docx" --catalog "数据目录数据.xlsx" --output "输出.docx"
"""

import argparse, re, sys, os, subprocess, io, zipfile, tempfile

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── Dependencies ──
def ensure_module(name, pip_name=None):
    try:
        __import__(name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name or name, "-q"])

ensure_module("openpyxl")
ensure_module("docx", "python-docx")

import openpyxl
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

HEADER_BG = "F1F1F1"

# ══════════════════════════════════════════════════════════════
# Shared Helpers
# ══════════════════════════════════════════════════════════════

def read_excel(excel_path, service_dir):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    rows = []
    for r in range(2, ws.max_row + 1):
        rd = {}
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=r, column=c).value
            rd[headers[c - 1]] = str(v).strip() if v is not None else ""
        if rd.get("服务目录") == service_dir:
            rd["_row"] = r
            rows.append(rd)
    if not rows:
        raise ValueError(f"未找到匹配数据: 服务目录={service_dir}")
    return rows

def add_solid_borders(tblPr):
    borders = OxmlElement("w:tblBorders")
    for edge in ["top", "left", "bottom", "right", "insideH", "insideV"]:
        e = OxmlElement("w:" + edge)
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), "4")
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), "000000")
        borders.append(e)
    tblPr.append(borders)

def mp(text, style_id, indent=None):
    p = OxmlElement("w:p")
    pr = OxmlElement("w:pPr")
    ps = OxmlElement("w:pStyle")
    ps.set(qn("w:val"), style_id)
    pr.append(ps)
    if indent:
        i = OxmlElement("w:ind")
        i.set(qn("w:firstLine"), str(indent))
        pr.append(i)
    p.append(pr)
    r = OxmlElement("w:r")
    tt = OxmlElement("w:t")
    tt.text = text
    tt.set(qn("xml:space"), "preserve")
    r.append(tt)
    p.append(r)
    return p

def make_cell_oxml(text, grid_span=None, bold=False, bg=None):
    tc = OxmlElement("w:tc")
    tcPr = OxmlElement("w:tcPr")
    if grid_span:
        gs = OxmlElement("w:gridSpan")
        gs.set(qn("w:val"), str(grid_span))
        tcPr.append(gs)
    if bg:
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:fill"), bg)
        tcPr.append(shd)
    tc.append(tcPr)
    p = OxmlElement("w:p")
    r = OxmlElement("w:r")
    if bold:
        rp = OxmlElement("w:rPr")
        b = OxmlElement("w:b")
        rp.append(b)
        r.append(rp)
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    p.append(r)
    tc.append(p)
    return tc

def make_table_oxml(headers, rows_data, col_widths=None):
    tbl = OxmlElement("w:tbl")
    tp = OxmlElement("w:tblPr")
    ts = OxmlElement("w:tblStyle")
    ts.set(qn("w:val"), "TableGrid")
    tp.append(ts)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), "5000")
    tw.set(qn("w:type"), "pct")
    tp.append(tw)
    add_solid_borders(tp)
    tbl.append(tp)
    ncols = len(headers)
    if col_widths is None:
        col_widths = [str(9000 // ncols)] * ncols
    tg = OxmlElement("w:tblGrid")
    for w in col_widths:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), w)
        tg.append(gc)
    tbl.append(tg)
    hdr = OxmlElement("w:tr")
    for h in headers:
        hdr.append(make_cell_oxml(h, bold=True, bg=HEADER_BG))
    tbl.append(hdr)
    for rd in rows_data:
        tr = OxmlElement("w:tr")
        for val in rd:
            tr.append(make_cell_oxml(val))
        tbl.append(tr)
    return tbl

def insert_after_oxml(ref_elem, new_elems):
    parent = ref_elem.getparent()
    children = list(parent)
    for i, child in enumerate(children):
        if child is ref_elem:
            for j, elem in enumerate(new_elems):
                parent.insert(i + 1 + j, elem)
            return

def parse_report_names(text_after_marker):
    lines = text_after_marker.strip().split("\n")
    names = []
    for line in lines:
        line = line.strip()
        if line and len(line) > 1:
            line = re.sub(r"^[\d]+[、．.）\)]\s*", "", line)
            line = line.strip("，。；,.;")
            if len(line) > 1:
                names.append(line)
    return names

def update_toc_via_com(doc_path):
    if sys.platform != "win32":
        print("非Windows环境，跳过TOC更新，请手动在Word中更新域。")
        return
    ps_script = f'''
$word = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open("{doc_path}")
    $tocCount = $doc.TablesOfContents.Count
    for ($i = $tocCount; $i -ge 1; $i--) {{ $doc.TablesOfContents.Item($i).Delete() }}
    for ($i = 1; $i -le $doc.Paragraphs.Count; $i++) {{
        if ($doc.Paragraphs.Item($i).Range.Text -match "文档介绍" -or $doc.Paragraphs.Item($i).Range.Text -match "需求来源") {{
            if ($i -gt 1) {{
                $tocRange = $doc.Paragraphs.Item($i - 1).Range
                $toc = $doc.TablesOfContents.Add($tocRange, $true, 1, 3, $false, "", $true, $true)
                $toc.Update()
            }}
            break
        }}
    }}
    $doc.Save()
    $doc.Close()
    Write-Host "TOC updated"
}} catch {{
    Write-Host "TOC error: $_"
}} finally {{
    if ($word) {{ try {{ $word.Quit() }} catch {{}} }}
}}
'''
    try:
        result = subprocess.run(
            ["powershell", "-ExecutionPolicy", "Bypass", "-Command", ps_script],
            capture_output=True, text=True, timeout=120
        )
        if "TOC updated" in result.stdout:
            print("TOC 已更新")
        else:
            print(f"TOC: {result.stdout.strip()[:200]}")
    except subprocess.TimeoutExpired:
        print("TOC 超时，请手动在Word中更新域")
    except Exception as e:
        print(f"TOC 失败: {e}，请手动在Word中更新域")

# ══════════════════════════════════════════════════════════════
# 01-数据报表_需求文档 Fill Logic
# ══════════════════════════════════════════════════════════════

def compute_stats(data_rows):
    unique_reqs = len(set(r["需求单号"] for r in data_rows if r["需求单号"]))
    unique_wos = len(set(r["工单号"] for r in data_rows if r["工单号"]))
    total_cnt = sum(int(r["报表统计次数"]) for r in data_rows if r["报表统计次数"].isdigit())
    return unique_reqs, unique_wos, total_cnt

def fill_requirement_doc(data_rows, template_path, output_path):
    """填充 01-数据报表_需求文档"""
    req_count, wo_count, total_reports = compute_stats(data_rows)
    print(f"结果: {len(data_rows)} 条, 需求单={req_count}, 工单={wo_count}, 报表={total_reports}")

    doc = Document(template_path)
    body = doc.element.body

    # ── Update count description ──
    count_text = f"服务周期内，共有{req_count}张需求单，{wo_count}张工单涉及{total_reports}次数据报表服务。具体需求单、工单和产出如下表："
    for p in doc.paragraphs:
        if p.style.name == "Body Text" and "服务周期内" in p.text:
            for run in p.runs:
                run.text = ""
            if p.runs:
                p.runs[0].text = count_text
            break

    # ── Table 2 rebuild ──
    t2 = doc.tables[2]
    tbl_elem = t2._tbl
    add_solid_borders(tbl_elem.find(qn('w:tblPr')))
    for tr_elem in list(tbl_elem.findall(qn("w:tr")))[1:]:
        tbl_elem.remove(tr_elem)
    for idx, rd in enumerate(data_rows):
        tr = OxmlElement("w:tr")
        for val in [str(idx + 1), rd["需求单号"], rd["工单号"], rd["工单内容"], rd["报表统计次数"]]:
            tr.append(make_cell_oxml(val))
        tbl_elem.append(tr)
    tr = OxmlElement("w:tr")
    tr.append(make_cell_oxml("合计", grid_span=2))
    tr.append(make_cell_oxml(""))
    tr.append(make_cell_oxml(""))
    tr.append(make_cell_oxml(str(total_reports)))
    tbl_elem.append(tr)

    # ── Section 3 ──
    style_ids = {
        "Heading 1": doc.styles["Heading 1"].style_id,
        "Heading 2": doc.styles["Heading 2"].style_id,
        "Heading 3": doc.styles["Heading 3"].style_id,
        "Body Text": doc.styles["Body Text"].style_id,
    }

    h1_content = h1_other = None
    for p in doc.paragraphs:
        if p.style.name == "Heading 1":
            if "需求内容" in p.text:
                h1_content = p._element
            elif "其他要求" in p.text:
                h1_other = p._element
                break
    if h1_content is None or h1_other is None:
        raise ValueError("模板中未找到 需求内容 或 其他要求 标题")

    remove = False
    to_remove = []
    for child in list(body):
        if child is h1_content:
            remove = True
            continue
        elif child is h1_other:
            remove = False
            continue
        if remove:
            to_remove.append(child)
    for elem in to_remove:
        body.remove(elem)

    new_elems = []
    for rd in data_rows:
        gongdan = rd["工单内容"]
        biz = rd["业务说明"]
        data_req = rd["数据需求"]
        delivery = rd["交付要求"]

        new_elems.append(mp(gongdan, style_ids["Heading 2"]))
        new_elems.append(mp("业务说明", style_ids["Heading 3"]))

        biz_match = re.search(r"本次工作拟产出以下(\d+)份报表成果", biz)
        if not biz_match:
            biz_match = re.search(r"本次工作拟产出(\d+)份报表成果", biz)

        if biz_match:
            clean_biz = biz[:biz_match.start()].strip().rstrip("\n\r")
            new_elems.append(mp(clean_biz, style_ids["Body Text"], 480))
            report_names = parse_report_names(biz[biz_match.end():])
            if report_names:
                new_elems.append(mp(
                    f"本次工作拟产出以下{len(report_names)}份报表成果：",
                    style_ids["Body Text"], 480))
                table_rows = [[str(i + 1), name] for i, name in enumerate(report_names)]
                new_elems.append(make_table_oxml(["序号", "名称"], table_rows))
                print(f"  [{gongdan[:30]}] 内嵌表格: {len(report_names)} 行")
        else:
            new_elems.append(mp(biz, style_ids["Body Text"], 480))

        new_elems.append(mp("数据需求", style_ids["Heading 3"]))
        new_elems.append(mp(data_req, style_ids["Body Text"], 480))
        new_elems.append(mp("交付要求", style_ids["Heading 3"]))
        new_elems.append(mp(delivery, style_ids["Body Text"], 480))

    insert_after_oxml(h1_content, new_elems)
    print(f"第3节: {len(new_elems)} 个元素")

    doc.save(output_path)
    print(f"已保存: {output_path}")
    update_toc_via_com(output_path)
    return output_path


# ══════════════════════════════════════════════════════════════
# 02-数据报表_设计文档 Fill Logic
# ══════════════════════════════════════════════════════════════

INDICATOR_HEADERS = ["所属报表名称", "指标名称", "指标定义", "指标业务口径", "指标展示形式", "与其他指标的关联关系", "备注"]
INDICATOR_DATA = [
    ["目录信息", "总条数", "各报表目录的总数据量条数", "无", "文本", "无", "总条数"],
    ["各报表目录详情", "库名", "目录所在库名", "无", "文本", "无", "库名"],
    ["各报表目录详情", "资源表名", "该目录所挂载的资源表名", "无", "文本", "无", "资源表名"],
    ["各报表目录详情", "表注释", "表中文名", "无", "文本", "无", "表注释"],
    ["各报表目录详情", "字段名", "字段英文名", "无", "文本", "无", "字段名"],
    ["各报表目录详情", "字段注释", "字段中文名", "无", "文本", "无", "字段注释"],
    ["各报表目录详情", "总数", "该目录中字段总条数", "无", "文本", "无", "总数"],
    ["各报表目录详情", "空值数", "该目录字段空值数", "无", "文本", "无", "空值数"],
    ["各报表目录详情", "空值率", "该目录字段空值率", "无", "文本", "无", "空值率"],
    ["各报表目录详情", "样例数据", "该字段示例", "无", "文本", "无", "样例数据"],
]

def mk_indicator_table():
    tbl = OxmlElement("w:tbl")
    tp = OxmlElement("w:tblPr")
    ts = OxmlElement("w:tblStyle")
    ts.set(qn("w:val"), "TableGrid")
    tp.append(ts)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), "5000")
    tw.set(qn("w:type"), "pct")
    tp.append(tw)
    add_solid_borders(tp)
    tbl.append(tp)
    n = 7
    tg = OxmlElement("w:tblGrid")
    for _ in range(n):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(9000 // n))
        tg.append(gc)
    tbl.append(tg)
    tr_h = OxmlElement("w:tr")
    for h_text in INDICATOR_HEADERS:
        tr_h.append(make_cell_oxml(h_text, bold=True, bg=HEADER_BG))
    tbl.append(tr_h)
    for rd in INDICATOR_DATA:
        tr = OxmlElement("w:tr")
        for v in rd:
            tr.append(make_cell_oxml(v))
        tbl.append(tr)
    return tbl

def mk_biz_table(items):
    tbl = OxmlElement("w:tbl")
    tp = OxmlElement("w:tblPr")
    ts = OxmlElement("w:tblStyle")
    ts.set(qn("w:val"), "TableGrid")
    tp.append(ts)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), "5000")
    tw.set(qn("w:type"), "pct")
    tp.append(tw)
    add_solid_borders(tp)
    tbl.append(tp)
    tg = OxmlElement("w:tblGrid")
    for w in ["2500", "2500"]:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), w)
        tg.append(gc)
    tbl.append(tg)
    for title, content in items:
        tr = OxmlElement("w:tr")
        tr.append(make_cell_oxml(title, bold=True, bg=HEADER_BG))
        tr.append(make_cell_oxml(content))
        tbl.append(tr)
    return tbl

def mk_ds_table(table_title, fields):
    tbl = OxmlElement("w:tbl")
    tp = OxmlElement("w:tblPr")
    ts = OxmlElement("w:tblStyle")
    ts.set(qn("w:val"), "TableGrid")
    tp.append(ts)
    tw = OxmlElement("w:tblW")
    tw.set(qn("w:w"), "5000")
    tw.set(qn("w:type"), "pct")
    tp.append(tw)
    add_solid_borders(tp)
    tbl.append(tp)
    tg = OxmlElement("w:tblGrid")
    for _ in range(4):
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), "2250")
        tg.append(gc)
    tbl.append(tg)
    # Row 0: merged title (name + code)
    tr0 = OxmlElement("w:tr")
    tr0.append(make_cell_oxml(table_title, grid_span=4, bold=True))
    tbl.append(tr0)
    # Row 1: column headers
    tr1 = OxmlElement("w:tr")
    for h_text in ["字段中文名称", "字段英文名称", "数据类型", "备注"]:
        tr1.append(make_cell_oxml(h_text, bold=True, bg=HEADER_BG))
    tbl.append(tr1)
    for f in fields:
        tr = OxmlElement("w:tr")
        for key in ["数据项名称", "英文名称", "数据类型", "备注"]:
            tr.append(make_cell_oxml(f.get(key, "")))
        tbl.append(tr)
    return tbl

def extract_images_from_excel(excel_path):
    """Extract embedded images from Excel, return dict mapping (row,col) -> bytes."""
    try:
        with zipfile.ZipFile(excel_path, "r") as z:
            all_images = [z.read(f) for f in sorted(n for n in z.namelist() if n.startswith("xl/media/image") and n.endswith(".png"))]
    except Exception:
        all_images = []
    wb_img = openpyxl.load_workbook(excel_path)
    ws_img = wb_img.active
    positions = sorted([
        (r, c) for r in range(2, ws_img.max_row + 1)
        for c in range(1, ws_img.max_column + 1)
        if ws_img.cell(row=r, column=c).value and "DISPIMG" in str(ws_img.cell(row=r, column=c).value)
    ])
    return {pos: all_images[i] if i < len(all_images) else None for i, pos in enumerate(positions)}


def extract_images_via_cellimages(excel_path):
    """Parse xl/cellimages.xml (WPS format) to get image bytes by name."""
    import zipfile as _zf
    result = {}
    with _zf.ZipFile(excel_path, 'r') as zf:
        if 'xl/cellimages.xml' not in zf.namelist():
            return result
        ci = zf.read('xl/cellimages.xml').decode('utf-8')
        name_to_rid = {}
        for block in re.findall(r'<etc:cellImage>(.*?)</etc:cellImage>', ci, re.DOTALL):
            nm = re.search(r'name="([^"]+)"', block)
            rm = re.search(r'r:embed="(rId\d+)"', block)
            if nm and rm:
                name_to_rid[nm.group(1)] = rm.group(1)
        rels = zf.read('xl/_rels/cellimages.xml.rels').decode('utf-8')
        rid_to_file = {}
        for m in re.finditer(r'Id="(rId\d+)".*?Target="(media/[^"]+)"', rels):
            rid_to_file[m.group(1)] = m.group(2)
        for name, rid in name_to_rid.items():
            if rid in rid_to_file:
                try:
                    result[name] = zf.read('xl/' + rid_to_file[rid])
                except:
                    pass
    return result


def match_images_to_cells(excel_path, img_bytes_by_name, img_cols, row_numbers):
    """Match DISPIMG formulas to images by name for given row numbers."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    result = {}
    for r in row_numbers:
        for cn in img_cols:
            if cn not in headers:
                continue
            ci = headers.index(cn) + 1
            val = ws.cell(row=r, column=ci).value
            if val and 'DISPIMG' in str(val):
                m = re.search(r'ID_([A-F0-9]+)', str(val))
                if m:
                    iname = "ID_" + m.group(1)
                    if iname in img_bytes_by_name:
                        result[(r, cn)] = img_bytes_by_name[iname]
    wb.close()
    return result

def load_catalog_data(catalog_path, all_codes):
    """Load resource info and field data from catalog Excel."""
    ensure_module("pandas")
    import pandas as pd
    df_r = pd.read_excel(catalog_path, sheet_name="关联资源信息")
    df_r["数据目录代码"] = df_r["数据目录代码"].astype(str).str.strip()
    rmap = {}
    for _, r in df_r[df_r["数据目录代码"].isin(all_codes)].iterrows():
        rmap[r["数据目录代码"]] = {
            "资源名称": str(r.get("资源名称", "")).strip(),
            "资源编码": str(r.get("资源编码", "")).strip()
        }
    df_i = pd.read_excel(catalog_path, sheet_name="数据项")
    df_i["数据目录代码"] = df_i["数据目录代码"].astype(str).str.strip()
    fmap = {c: [] for c in all_codes}
    for _, r in df_i[df_i["数据目录代码"].isin(all_codes)].iterrows():
        desc = str(r.get("字段描述", "")).strip()
        if not desc or desc.lower() in ("nan", "none"):
            desc = "No"
        fmap[r["数据目录代码"]].append({
            "数据项名称": str(r.get("数据项名称", "")).strip(),
            "英文名称": str(r.get("英文名称", "")).strip(),
            "数据类型": str(r.get("数据类型", "")).strip(),
            "备注": desc
        })
    return rmap, fmap

def fill_design_doc(data_rows, template_path, output_path, catalog_path):
    """Fill 02-数据报表_设计文档"""
    print(f"结果: {len(data_rows)} 条记录")

    # Extract images from ledger
    img_by_pos = extract_images_from_excel(template_path)  # placeholder; we use excel in extract call
    # Actually we need to extract from the original excel passed via --excel, but we already have data_rows.
    # We need to re-load the excel for images. Let us accept excel_path as additional param.
    # For now, we accept that images are extracted from the same excel that was read.
    # But data_rows are already filtered. We need the original excel path.

    # Since data_rows already passed in, we need the excel path separately for images.
    # Let me adjust: fill_design_doc needs excel_path too.
    pass  # Will be restructured below

def fill_design_doc_full(excel_path, data_rows, template_path, output_path, catalog_path):
    """Complete 02-设计文档 fill, with image extraction from excel."""
    img_by_pos = extract_images_from_excel(excel_path)

    # Map images to rows based on DISPIMG column
    # Need to map data_rows back to their row indices. Re-read excel for positions.
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col_idx = {h: i + 1 for i, h in enumerate(headers)}

    for r in range(2, ws.max_row + 1):
        svc = str(ws.cell(row=r, column=col_idx.get("服务目录", 1)).value or "").strip()
        if svc != data_rows[0].get("服务目录", ""):
            continue
        gd = str(ws.cell(row=r, column=col_idx.get("工单内容", 1)).value or "").strip()
        for rd in data_rows:
            if rd["工单内容"] == gd:
                img_content_col = col_idx.get("02-数据报表_设计文档-数据内容截图")
                img_logic_col = col_idx.get("数据处理逻辑")
                if img_content_col:
                    rd["_img_content"] = img_by_pos.get((r, img_content_col))
                if img_logic_col:
                    rd["_img_logic"] = img_by_pos.get((r, img_logic_col))
                    rd["_logic_is_img"] = "DISPIMG" in rd.get("数据处理逻辑", "")
                break

    # Collect all directory codes
    all_codes = set()
    code_col = "02-数据报表_设计文档-数据来源库表清单对应数据目录代码"
    for rd in data_rows:
        codes_str = rd.get(code_col, "")
        rd["_codes"] = [c.strip() for c in codes_str.split("\n") if c.strip()]
        all_codes.update(rd["_codes"])

    # Load catalog
    rmap, fmap = load_catalog_data(catalog_path, all_codes)

    # Open template and find insertion point
    doc = Document(template_path)
    body = doc.element.body
    all_c = list(body)

    h1d_idx = None
    for i, c in enumerate(all_c):
        tag = c.tag.split("}")[-1]
        if tag == "p":
            for pi, p in enumerate(doc.paragraphs):
                if p._element is c and p.style.name == "Heading 1" and "数据报表设计" in p.text:
                    h1d_idx = i
                    break
        if h1d_idx is not None:
            break
    if h1d_idx is None:
        raise ValueError("模板中未找到「数据报表设计」 Heading 1 标题")

    for i in range(len(all_c) - 1, h1d_idx, -1):
        body.remove(all_c[i])

    S = {
        "H2": doc.styles["Heading 2"].style_id,
        "H3": doc.styles["Heading 3"].style_id,
        "H4": doc.styles["Heading 4"].style_id,
        "BT": doc.styles["Body Text"].style_id,
        "NL": doc.styles["Normal"].style_id,
    }

    print("构建文档内容...")
    for idx, rd in enumerate(data_rows):
        gd = rd["工单内容"]
        codes = rd.get("_codes", [])

        body.append(mp(gd, S["H2"]))
        body.append(mp("业务分析", S["H3"]))
        body.append(mk_biz_table([
            ("内容描述", rd.get("内容描述", "")),
            ("业务场景", rd.get("业务场景", "")),
            ("数据内容", rd.get("数据内容", "")),
            ("结果形式", rd.get("结果形式", "")),
        ]))

        body.append(mp("数据内容", S["H3"]))
        body.append(mp("", S["NL"]))

        body.append(mp("数据来源库表清单", S["H3"]))
        body.append(mp("为完成本次数据报表，需使用到以下数据：", S["BT"], 480))

        if codes:
            for code in codes:
                res = rmap.get(code, {})
                rn = res.get("资源名称", "")
                rc = res.get("资源编码", "")
                h4_title = rc if rc else code
                table_title = f"{rn} {rc}" if rn and rc else code
                body.append(mp(h4_title, S["H4"]))
                fs = fmap.get(code, [])
                if fs:
                    body.append(mk_ds_table(table_title, fs))
                else:
                    body.append(mp("未匹配到此目录的任何表结构信息，请手动补充。", S["BT"], 480))
        else:
            body.append(mp("（无数据来源库表清单）", S["BT"], 480))

        body.append(mp("数据处理逻辑", S["H3"]))
        if rd.get("_logic_is_img") and rd.get("_img_logic"):
            body.append(mp("", S["NL"]))
        else:
            body.append(mp(rd.get("数据处理逻辑", ""), S["BT"], 480))

        body.append(mp("报表指标设计", S["H3"]))
        body.append(mk_indicator_table())

    # Insert images
    h3s = [(i, p) for i, p in enumerate(doc.paragraphs) if p.style.name == "Heading 3"]
    for ri, pi in enumerate([i for i, p in h3s if p.text.strip() == "数据内容"]):
        if ri < len(data_rows) and data_rows[ri].get("_img_content") and pi + 1 < len(doc.paragraphs):
            np = doc.paragraphs[pi + 1]
            if np.style.name == "Normal" and not np.text.strip():
                tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                tf.write(data_rows[ri]["_img_content"])
                tf.close()
                for re_e in np._element.findall(qn("w:r")):
                    np._element.remove(re_e)
                np.add_run().add_picture(tf.name, width=Inches(5.5))
                os.unlink(tf.name)

    for ri, pi in enumerate([i for i, p in h3s if p.text.strip() == "数据处理逻辑"]):
        if ri < len(data_rows) and data_rows[ri].get("_logic_is_img") and data_rows[ri].get("_img_logic") and pi + 1 < len(doc.paragraphs):
            np = doc.paragraphs[pi + 1]
            if np.style.name == "Normal" and not np.text.strip():
                tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                tf.write(data_rows[ri]["_img_logic"])
                tf.close()
                for re_e in np._element.findall(qn("w:r")):
                    np._element.remove(re_e)
                np.add_run().add_picture(tf.name, width=Inches(5.5))
                os.unlink(tf.name)

    doc.save(output_path)
    print(f"已保存: {output_path}")
    update_toc_via_com(output_path)
    return output_path


# ══════════════════════════════════════════════════════════════

# ══════════════════════════════════════════════════════════════
# 03-数据报表_上线记录 Fill Logic
# ══════════════════════════════════════════════════════════════

def fill_launch_record_doc(excel_path, data_rows, template_path, output_path):
    """Fill 03-数据报表_上线记录 template."""
    img_cols = ["上线交付截图1", "上线交付截图2", "使用记录截图1", "使用记录截图2"]
    print("提取图片...")
    imgs = extract_images_via_cellimages(excel_path)
    row_nums = set(rd["_row"] for rd in data_rows)
    cell_imgs = match_images_to_cells(excel_path, imgs, img_cols, row_nums)
    print(f"匹配到 {len(cell_imgs)} 张图片")

    doc = Document(template_path)
    body = doc.element.body

    S = {}
    for sn in ["Heading 1", "Heading 2", "Heading 3", "Body Text", "Normal"]:
        try:
            S[sn] = doc.styles[sn].style_id
        except:
            S[sn] = sn

    # Find key positions
    src_para = None
    desc_para = None
    for i, p in enumerate(doc.paragraphs):
        if p.style.name == "Heading 1" and "需求来源" in p.text:
            src_para = i
            if i + 1 < len(doc.paragraphs):
                desc_para = i + 1
            break
    if src_para is None:
        raise ValueError("模板中未找到'需求来源'章节")

    children = list(body)
    src_body = next(j for j, c in enumerate(children) if c is doc.paragraphs[src_para]._element)

    # Find table after 需求来源
    tbl_body = None
    for j in range(src_body + 1, len(children)):
        if children[j].tag == qn('w:tbl'):
            tbl_body = j
            break

    # Find first 工单 Heading 1
    gd_body = None
    for j in range(tbl_body + 1 if tbl_body else src_body + 1, len(children)):
        if children[j].tag == qn('w:p'):
            for p in doc.paragraphs:
                if p._element is children[j] and p.style.name == "Heading 1" and "上线记录" in p.text:
                    gd_body = j
                    break
        if gd_body:
            break

    # Update description
    ureq = len(set(r["需求单号"] for r in data_rows))
    ugd = len(set(r["工单号"] for r in data_rows))
    trep = sum(int(r.get("报表统计次数", 0) or 0) for r in data_rows)
    nd = f"服务周期内，共有{ureq}张需求单，{ugd}张工单涉及{trep}次数据报表服务。具体需求单、工单和产出如下表："
    dp = doc.paragraphs[desc_para]
    if dp.runs:
        dp.runs[0].text = nd
    for rn in dp.runs[1:]:
        rn.text = ""
    
    # Fill table with solid borders
    tbl = children[tbl_body]
    for tr in tbl.findall(qn('w:tr'))[1:]:
        tbl.remove(tr)
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        tbl.insert(0, tblPr)
    for eb in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(eb)
    for ts in tblPr.findall(qn('w:tblStyle')):
        tblPr.remove(ts)
    add_solid_borders(tblPr)

    for idx, rd in enumerate(data_rows):
        tr = OxmlElement("w:tr")
        tr.append(make_cell_oxml(str(idx + 1)))
        tr.append(make_cell_oxml(rd.get("需求单号", "")))
        tr.append(make_cell_oxml(rd.get("工单号", "")))
        tr.append(make_cell_oxml(rd.get("工单内容", "")))
        tr.append(make_cell_oxml(rd.get("报表统计次数", "")))
        tbl.append(tr)

    # Total row
    trt = OxmlElement("w:tr")
    trt.append(make_cell_oxml("合计", grid_span=2, bold=True))
    trt.append(make_cell_oxml(""))
    trt.append(make_cell_oxml(""))
    trt.append(make_cell_oxml(str(trep), bold=True))
    tbl.append(trt)
    print(f"表格: {len(data_rows)} 行 + 合计")

    # Remove old 工单 content
    if gd_body:
        children = list(body)
        for j in range(len(children) - 1, gd_body - 1, -1):
            body.remove(children[j])

    # Build 工单 sections
    body.append(mp("", S["Normal"]))
    print(f"构建 {len(data_rows)} 个工单章节...")
    for rd in data_rows:
        gd = rd.get("工单内容", "")
        body.append(mp(f"{gd}的上线记录", S["Heading 1"]))
        body.append(mp("产出说明", S["Heading 2"]))
        body.append(mp(f"需求编号：\t{rd.get('需求单号', '')}\t对应工单编号：\t{rd.get('工单号', '')}", S["Body Text"]))
        body.append(mp(f"需求描述：{rd.get('需求描述', '')}", S["Body Text"], 480))
        body.append(mp(f"统计报表：{rd.get('报表统计次数', '')}次。", S["Normal"]))
        body.append(mp("上线交付截图", S["Heading 2"]))
        body.append(mp("", S["Body Text"]))
        body.append(mp("", S["Body Text"]))
        body.append(mp("使用记录", S["Heading 2"]))
        body.append(mp("使用记录如下：", S["Body Text"]))
        body.append(mp("", S["Body Text"]))
        body.append(mp("", S["Body Text"]))

    doc.save(output_path)

    # Insert images
    print("插入图片...")
    doc2 = Document(output_path)
    h2s = [(i, p) for i, p in enumerate(doc2.paragraphs) if p.style.name == "Heading 2"]
    gi = 0
    for pi, p in h2s:
        if p.text.strip() == "上线交付截图":
            if gi < len(data_rows):
                rn = data_rows[gi]["_row"]
                for o, cn in enumerate(["上线交付截图1", "上线交付截图2"]):
                    tpi = pi + 1 + o
                    if tpi < len(doc2.paragraphs) and (rn, cn) in cell_imgs:
                        tp = doc2.paragraphs[tpi]
                        tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        tf.write(cell_imgs[(rn, cn)])
                        tf.close()
                        tp.add_run().add_picture(tf.name, width=Inches(5.5))
                        os.unlink(tf.name)
        elif p.text.strip() == "使用记录":
            if gi < len(data_rows):
                rn = data_rows[gi]["_row"]
                for o, cn in enumerate(["使用记录截图1", "使用记录截图2"]):
                    tpi = pi + 2 + o
                    if tpi < len(doc2.paragraphs) and (rn, cn) in cell_imgs:
                        tp = doc2.paragraphs[tpi]
                        tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
                        tf.write(cell_imgs[(rn, cn)])
                        tf.close()
                        tp.add_run().add_picture(tf.name, width=Inches(5.5))
                        os.unlink(tf.name)
                gi += 1

    doc2.save(output_path)
    update_toc_via_com(output_path)
    return output_path

# Dispatcher
# ══════════════════════════════════════════════════════════════

def fill_document(excel_path, service_dir, material_type, template_path, output_path, catalog_path=None):
    print(f"台账清单: {excel_path}")
    print(f"筛选条件: 服务目录={service_dir}")
    print(f"材料类型: {material_type}")
    data_rows = read_excel(excel_path, service_dir)

    if material_type == "01-数据报表_需求文档":
        return fill_requirement_doc(data_rows, template_path, output_path)
    elif material_type == "02-数据报表_设计文档":
        if not catalog_path:
            raise ValueError("02-设计文档需要 --catalog 参数指定数据目录数据路径")
        if not os.path.exists(catalog_path):
            raise FileNotFoundError(f"数据目录数据文件不存在: {catalog_path}")
        return fill_design_doc_full(excel_path, data_rows, template_path, output_path, catalog_path)
    elif material_type == "03-数据报表_上线记录":
        return fill_launch_record_doc(excel_path, data_rows, template_path, output_path)
    else:
        raise ValueError(f"不支持的材料类型: {material_type}。当前支持: 01-数据报表_需求文档, 02-数据报表_设计文档, 03-数据报表_上线记录")


# ══════════════════════════════════════════════════════════════
# CLI
# ══════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description="验收文档自动填充工具")
    parser.add_argument("--service-dir", required=True, help="服务目录筛选值，如 N08-数据报表服务")
    parser.add_argument("--material-type", required=True, help="材料类型，如 01-数据报表_需求文档 或 02-数据报表_设计文档")
    parser.add_argument("--excel", required=True, help="台账清单Excel文件路径")
    parser.add_argument("--template", required=True, help="Word模板文件路径")
    parser.add_argument("--output", required=True, help="输出Word文档路径")
    parser.add_argument("--catalog", default=None, help="数据目录数据Excel路径（02-设计文档必填）")
    args = parser.parse_args()

    try:
        result = fill_document(
            excel_path=args.excel,
            service_dir=args.service_dir,
            material_type=args.material_type,
            template_path=args.template,
            output_path=args.output,
            catalog_path=args.catalog,
        )
        print(f"\n✅ 完成: {result}")
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()



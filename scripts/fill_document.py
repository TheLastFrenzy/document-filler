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

import argparse, re, sys, os, subprocess, io, zipfile, tempfile, importlib.util
from pathlib import Path

if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        if hasattr(sys.stdout, "buffer"):
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

# ── Dependencies ──
def ensure_module(name, pip_name=None):
    try:
        __import__(name)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name or name, "-q"])

ensure_module("openpyxl")
ensure_module("docx", "python-docx")
ensure_module("olefile")

import openpyxl
from docx import Document
from docx.shared import Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import struct

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
            rd["_vml_row"] = r - 1
            rd["_vml_col"] = headers.index("业务逻辑") if "业务逻辑" in headers else -1
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


def xml_safe(t):
    if not t: return ""
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f-\x9f]', '', t)
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

def make_cell_oxml(text, grid_span=None, bold=False, bg=None, align=None):
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
    if align:
        pp = OxmlElement("w:pPr")
        jc = OxmlElement("w:jc")
        jc.set(qn("w:val"), align)
        pp.append(jc)
        p.append(pp)
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


def extract_ole_by_cell(excel_path):
    import olefile as _ole
    result = {}
    with zipfile.ZipFile(excel_path, 'r') as zf:
        vml = zf.read('xl/drawings/vmlDrawing1.vml').decode('utf-8')
        sx = zf.read('xl/worksheets/sheet1.xml').decode('utf-8')
        rx = zf.read('xl/worksheets/_rels/sheet1.xml.rels').decode('utf-8')
        vml_shapes = {}
        for m in re.finditer(r'<v:shape\b(.*?)</v:shape>', vml, re.DOTALL):
            body = m.group(1); id_m = re.search(r'id="([^"]+)"', body)
            am = re.search(r'<x:Anchor>([^<]+)</x:Anchor>', body)
            if id_m and am:
                parts = am.group(1).split(',')
                vml_shapes[id_m.group(1)] = {'row': int(parts[2].strip()), 'col': int(parts[0].strip()), 'ole': 'o:ole="t"' in body}
        ole_map = {}
        for m in re.finditer(r'<oleObject[^>]*shapeId="(\d+)"[^>]*r:id="(rId\d+)"', sx):
            sn = m.group(1)
            if sn not in ole_map: ole_map[sn] = m.group(2)
        rid2f = {}
        for m in re.finditer(r'<Relationship[^>]*Id="(rId\d+)"[^>]*oleObject[^>]*Target="([^"]+)"', rx):
            rid2f[m.group(1)] = m.group(2).split('/')[-1]
        for sn, rid in ole_map.items():
            if rid not in rid2f: continue
            for vid, vs in vml_shapes.items():
                if vs.get('ole') and vid.endswith('_s' + sn):
                    fn = rid2f[rid]; vr = vs['row']; vc = vs['col']
                    try: od = zf.read(f'xl/embeddings/{fn}')
                    except: continue
                    tf = tempfile.NamedTemporaryFile(suffix='.bin', delete=False)
                    tf.write(od); tf.close()
                    ole = _ole.OleFileIO(tf.name)
                    wd = ole.openstream('WordDocument').read()
                    ole.close(); os.unlink(tf.name)
                    chunks = []; i = 0
                    while i < len(wd) - 1:
                        cc = struct.unpack_from('<H', wd, i)[0]
                        if 0x20 <= cc <= 0xFFFF and cc != 0xFFFE:
                            si = i
                            while i < len(wd) - 1:
                                cc2 = struct.unpack_from('<H', wd, i)[0]
                                if cc2 in (0x000D, 0x0007, 0x0000): break
                                i += 2
                            try:
                                text = wd[si:i].decode('utf-16-le', errors='replace')
                                if len(text.strip()) > 2: chunks.append(text)
                            except: pass
                        i += 2
                    raw = ''.join(chunks)
                    clean = re.sub(r'[^\u4e00-\u9fff\u3000-\u303f\uff00-\uffefa-zA-Z0-9\s\.\,\;\:\!\?\-\+\=\(\)\[\]\{\}\/\\\@\#\%\&\*_\u201c\u201d\u2018\u2019\u3001\u3002\u2014\u2026\n\r\t]', '', raw)
                    clean = re.sub(r'\n{3,}', '\n\n', clean)
                    logic = clean
                    m2 = re.search(r'需求处理逻辑(.+)', clean, re.DOTALL)
                    if m2: logic = m2.group(1)
                    em2 = re.search(r'(需求验收标准|完成时间|资源保证|风险评估)', logic)
                    if em2: logic = logic[:em2.start()]
                    result[(vr, vc)] = logic.strip()
                    break
    return result
def compute_stats(data_rows):
    unique_reqs = len(set(r["需求单号"] for r in data_rows if r["需求单号"]))
    unique_wos = len(set(r["工单号"] for r in data_rows if r["工单号"]))
    total_cnt = sum(int(r["报表统计次数"]) for r in data_rows if r["报表统计次数"].isdigit())
    return unique_reqs, unique_wos, total_cnt


def summarize_biz_logic(text):
    if not text: return [("1", "详见附件需求规格说明书")]
    steps = []; text = text.strip()
    markers = list(re.finditer(r'(?:^|(?<=[\n\u3002\u3001])\s*)(\d+)[.\u3001)\uff09]\s*', text))
    if len(markers) >= 2:
        for i, m in enumerate(markers):
            start = m.end(); end = markers[i+1].start() if i+1 < len(markers) else len(text)
            cp = text[start:end].strip()
            if len(cp) > 3: steps.append((str(i+1), xml_safe(cp[:250])))
    else:
        segs = re.split(r'(?<=[。、])\s*', text); sn = 1; cur = ""
        for seg in segs:
            seg = seg.strip()
            if not seg: continue
            if len(cur) + len(seg) > 200 and cur:
                steps.append((str(sn), xml_safe(cur[:250]))); sn += 1; cur = seg
            else:
                cur = (cur + "。" + seg) if cur else seg
        if cur: steps.append((str(sn), xml_safe(cur[:250])))
    if not steps: steps = [("1", xml_safe(text[:200]))]
    return steps[:15]


def parse_report_list(text):
    if not text: return []
    items = []
    for line in text.strip().split('\n'):
        line = line.strip()
        if not line: continue
        parts = line.rsplit(None, 1) if ' ' in line else [line, '']
        cn = parts[0].strip(); en = parts[1].strip() if len(parts) > 1 else ''
        if cn: items.append((cn, en))
    return items


def norm_table_name(name):
    text = (name or "").strip()
    if "." in text:
        text = text.split(".")[-1]
    return re.sub(r"[^A-Za-z0-9_]", "", text).upper()


def merged_value_getter(ws):
    merged_values = {}
    for merged_range in ws.merged_cells.ranges:
        value = ws.cell(merged_range.min_row, merged_range.min_col).value
        for row in range(merged_range.min_row, merged_range.max_row + 1):
            for col in range(merged_range.min_col, merged_range.max_col + 1):
                merged_values[(row, col)] = value

    def get(row, col):
        value = ws.cell(row, col).value
        if value is None:
            value = merged_values.get((row, col))
        return str(value).strip() if value is not None else ""

    return get


def read_stats_requirement_groups(excel_path, service_dir):
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    ws = wb.active
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    if "服务目录" not in headers:
        raise ValueError("台账清单缺少「服务目录」列")
    get = merged_value_getter(ws)
    service_col = headers.index("服务目录") + 1
    logic_col = headers.index("业务逻辑") if "业务逻辑" in headers else -1
    groups = []
    by_key = {}

    for row in range(2, ws.max_row + 1):
        if get(row, service_col) != service_dir:
            continue
        record = {header: get(row, idx + 1) for idx, header in enumerate(headers) if header}
        key = record.get("工单号") or f"{record.get('需求单号', '')}|{record.get('工单内容', '')}|{row}"
        if key not in by_key:
            record["_row"] = row
            record["_vml_row"] = row - 1
            record["_vml_col"] = logic_col
            record["results"] = []
            by_key[key] = record
            groups.append(record)

        result_text = record.get("统计分析结果表清单", "")
        for result in parse_report_list(result_text):
            if result not in by_key[key]["results"]:
                by_key[key]["results"].append(result)

    for group in groups:
        group["统计分析结果表清单"] = "\n".join(
            f"{cn} {en}".strip() for cn, en in group.get("results", [])
        )

    if not groups:
        raise ValueError(f"未找到匹配数据: 服务目录={service_dir}")
    return groups


def resolve_stats_relation_workbook(template_path, relation_hint=None):
    candidates = []
    if relation_hint:
        candidates.append(Path(relation_hint))
    template = Path(template_path)
    candidates.append(template.with_name("04-数据统计分析_结果表及使用说明.xlsx"))
    candidates.append(template.parent / "04-数据统计分析_结果表及使用说明.xlsx")

    seen = set()
    for candidate in candidates:
        if not candidate or candidate in seen or not candidate.exists():
            continue
        seen.add(candidate)
        try:
            wb = openpyxl.load_workbook(candidate, read_only=True, data_only=True)
            if "2、表融合关系" in wb.sheetnames:
                wb.close()
                return str(candidate)
            wb.close()
        except Exception:
            continue
    return None


def _parse_relation_title(text):
    value = str(text or "").strip()
    if not value:
        return ""
    match = re.search(r"([A-Za-z][A-Za-z0-9_$.]{2,})\s*$", value)
    return norm_table_name(match.group(1)) if match else ""


def load_stats_relation_descriptions(relation_path):
    wb = openpyxl.load_workbook(relation_path, data_only=True)
    if "2、表融合关系" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["2、表融合关系"]
    descriptions = {}
    current_result = ""
    for row in range(1, ws.max_row + 1):
        first_value = ws.cell(row, 1).value
        parsed_result = _parse_relation_title(first_value)
        if parsed_result:
            current_result = parsed_result
        for col in range(1, ws.max_column + 1):
            label = str(ws.cell(row, col).value or "").strip()
            if label == "文字描述" and current_result:
                desc = ws.cell(row, col + 1).value if col + 1 <= ws.max_column else ""
                if desc:
                    descriptions[current_result] = str(desc).strip()
                break
    wb.close()
    return descriptions


def _clean_logic_step(text):
    text = xml_safe(str(text or ""))
    text = re.sub(r"^\s*\d+\s*[\.、．)\uff09]?\s*", "", text)
    text = re.sub(r"[\t ]+", " ", text)
    return text.strip(" \r\n\t；;。")


def split_logic_description(text):
    text = str(text or "").strip()
    if not text:
        return []
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    markers = list(re.finditer(r"(?:^|\n)\s*(\d+)\s*(?:[\.、．)\uff09]|\t|\s)\s*", normalized))
    steps = []
    if len(markers) >= 2:
        for idx, marker in enumerate(markers):
            start = marker.end()
            end = markers[idx + 1].start() if idx + 1 < len(markers) else len(normalized)
            piece = _clean_logic_step(normalized[start:end])
            if piece:
                steps.append(piece)
    else:
        for line in normalized.split("\n"):
            piece = _clean_logic_step(line)
            if piece:
                steps.append(piece)

    if len(steps) <= 1:
        source = steps[0] if steps else _clean_logic_step(normalized)
        steps = [
            _clean_logic_step(piece)
            for piece in re.split(r"(?<=[。；;])\s*", source)
            if _clean_logic_step(piece)
        ]

    return steps[:15]


def build_stats_requirement_business_logic_steps(results, descriptions):
    if not results:
        return [("1", "未匹配到统计分析结果表，请手动补充业务逻辑说明。")]

    all_steps = []
    for cn, en in results:
        desc = descriptions.get(norm_table_name(en), "")
        pieces = split_logic_description(desc)
        table_label = f"{cn}（{en}）" if en else cn
        if not pieces:
            all_steps.append(f"{table_label}：未在结果表及使用说明中匹配到文字描述，请手动补充。")
        elif len(results) == 1:
            all_steps.extend(pieces)
        else:
            merged = "；".join(piece.rstrip("。；;") for piece in pieces)
            all_steps.append(f"{table_label}：{merged}。")

    return [(str(idx), xml_safe(step[:500])) for idx, step in enumerate(all_steps, start=1)]


def make_biz_table(header_text, rows_data):
    tbl = OxmlElement("w:tbl"); tp = OxmlElement("w:tblPr")
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), "5000"); tw.set(qn("w:type"), "pct"); tp.append(tw)
    add_solid_borders(tp); tbl.append(tp)
    tg = OxmlElement("w:tblGrid")
    tg.append(OxmlElement("w:gridCol")); tg[-1].set(qn("w:w"), "900")
    tg.append(OxmlElement("w:gridCol")); tg[-1].set(qn("w:w"), "8100"); tbl.append(tg)
    tr0 = OxmlElement("w:tr")
    tr0.append(make_cell_oxml(header_text, grid_span=2, bold=True, bg=HEADER_BG, align="center"))
    tbl.append(tr0)
    tr1 = OxmlElement("w:tr")
    tr1.append(make_cell_oxml("步骤", bold=True, bg=HEADER_BG))
    tr1.append(make_cell_oxml("说明", bold=True, bg=HEADER_BG))
    tbl.append(tr1)
    for sn, st in rows_data:
        tr = OxmlElement("w:tr")
        tr.append(make_cell_oxml(sn))
        tr.append(make_cell_oxml(st))
        tbl.append(tr)
    return tbl


def make_stats_source_table(headers, rows_data):
    tbl = OxmlElement("w:tbl"); tp = OxmlElement("w:tblPr")
    tw = OxmlElement("w:tblW"); tw.set(qn("w:w"), "5000"); tw.set(qn("w:type"), "pct"); tp.append(tw)
    add_solid_borders(tp); tbl.append(tp)
    n = len(headers)
    tg = OxmlElement("w:tblGrid")
    for _ in range(n):
        gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), str(9000 // n)); tg.append(gc)
    tbl.append(tg)
    tr = OxmlElement("w:tr")
    for h in headers:
        tr.append(make_cell_oxml(h, bold=True, bg=HEADER_BG))
    tbl.append(tr)
    for rd in rows_data:
        tr = OxmlElement("w:tr")
        for v in rd: tr.append(make_cell_oxml(v))
        tbl.append(tr)
    return tbl
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
        for m in re.finditer(r'<Relationship[^>]*Id="(rId\d+)"[^>]*Target="([^"]+)"', rels):
            target = m.group(2)
            if target.startswith("../"):
                target = "xl/" + target[3:]
            elif target.startswith("xl/"):
                target = target
            else:
                target = "xl/" + target.lstrip("/")
            rid_to_file[m.group(1)] = target
        for name, rid in name_to_rid.items():
            if rid in rid_to_file:
                try:
                    result[name] = zf.read(rid_to_file[rid])
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
    # Extract images via cellimages.xml for name-based matching (not positional)
    imgs = extract_images_via_cellimages(excel_path)
    row_nums = set(rd["_row"] for rd in data_rows)
    img_cols = ["02-数据报表_设计文档-数据内容截图", "数据处理逻辑"]
    cell_imgs = match_images_to_cells(excel_path, imgs, img_cols, row_nums)
    print(f"匹配到 {len(cell_imgs)} 张图片")

    for rd in data_rows:
        rd["_img_content"] = cell_imgs.get((rd["_row"], "02-数据报表_设计文档-数据内容截图"))
        rd["_img_logic"] = cell_imgs.get((rd["_row"], "数据处理逻辑"))
        rd["_logic_is_img"] = "DISPIMG" in rd.get("数据处理逻辑", "")

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

# ================================================================================
# 01-数据统计分析_需求文档 Fill Logic
# ================================================================================

def fill_stats_requirement_doc(excel_path, data_rows, template_path, output_path, relation_path=None):
    """填充 01-数据统计分析_需求文档"""
    print(f"数据: {len(data_rows)} 条")

    relation_descriptions = {}
    if relation_path:
        print(f"读取表融合关系文字描述: {relation_path}")
        relation_descriptions = load_stats_relation_descriptions(relation_path)
        print(f"文字描述匹配: {len(relation_descriptions)} 个结果表")

    bts = {}
    if not relation_descriptions:
        print("提取业务逻辑附件...")
        obc = extract_ole_by_cell(excel_path)
        for rd in data_rows:
            k = (rd.get('_vml_row', -1), rd.get('_vml_col', -1))
            if k in obc:
                bts[rd.get("工单内容", "")] = obc[k]
        print(f"附件匹配: {len(bts)} 个")

    doc = Document(template_path)
    body = doc.element.body

    S = {}
    for sn in ["Heading 1", "Heading 2", "Heading 3", "Body Text", "Normal"]:
        try: S[sn] = doc.styles[sn].style_id
        except: S[sn] = sn

    # Find "需求来源" description paragraph to update count
    src_para = desc_para = None
    for i, p in enumerate(doc.paragraphs):
        if p.style.name == "Heading 2" and "需求来源" in p.text:
            src_para = i; desc_para = i + 1 if i + 1 < len(doc.paragraphs) else None; break
    if src_para is None:
        raise ValueError("未找到需求来源")

    # Find "需求内容" H1
    content_h1 = None
    for i, p in enumerate(doc.paragraphs):
        if p.style.name == "Heading 1" and "需求内容" in p.text:
            content_h1 = i; break

    # Build body children index
    children = list(body)
    src_body = next(j for j, c in enumerate(children) if c is doc.paragraphs[src_para]._element)
    tbl_body = next(j for j in range(src_body + 1, len(children)) if children[j].tag == qn('w:tbl'))
    content_body = next(j for j, c in enumerate(children) if c is doc.paragraphs[content_h1]._element)

    # Find first H2 after 需求内容 (the first old gongdan)
    gd_body = None
    for j in range(content_body + 1, len(children)):
        if children[j].tag == qn('w:p'):
            for p in doc.paragraphs:
                if p._element is children[j] and p.style.name == "Heading 2":
                    gd_body = j; break
        if gd_body is not None: break

    # Find "其他要求" H1
    other_h1_body = None
    for j in range(content_body + 1, len(children)):
        if children[j].tag == qn('w:p'):
            for p in doc.paragraphs:
                if p._element is children[j] and p.style.name == "Heading 1" and "其他要求" in p.text:
                    other_h1_body = j; break
        if other_h1_body is not None: break

    # Update count description
    ureq = len(set(r["需求单号"] for r in data_rows))
    ugd = len(set(r["工单号"] for r in data_rows))
    trep = sum(int(r.get("报表统计次数", 0) or 0) for r in data_rows)
    nd = f"服务周期内，共有{ureq}张需求单，{ugd}张工单涉及{trep}次数据统计分析。具体需求单、工单和产出如下表："
    
    dp = doc.paragraphs[desc_para]
    if dp.runs:
        dp.runs[0].text = nd
    for rn in dp.runs[1:]:
        rn.text = ""

    # Rebuild table
    tbl = children[tbl_body]
    for tr in tbl.findall(qn('w:tr'))[1:]:
        tbl.remove(tr)
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr"); tbl.insert(0, tblPr)
    for eb in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(eb)
    for ts in tblPr.findall(qn('w:tblStyle')):
        tblPr.remove(ts)
    add_solid_borders(tblPr)

    # Fixed layout + column widths 1:2:2:4:1
    tw_el = tblPr.find(qn('w:tblW'))
    if tw_el is not None:
        tw_el.set(qn('w:type'), 'pct'); tw_el.set(qn('w:w'), '5000')
    for tl in tblPr.findall(qn('w:tblLayout')):
        tblPr.remove(tl)
    tl = OxmlElement('w:tblLayout'); tl.set(qn('w:type'), 'fixed'); tblPr.append(tl)
    tg = tbl.find(qn('w:tblGrid'))
    if tg is not None:
        for gc in list(tg):
            tg.remove(gc)
        for w in ["900", "1800", "1800", "3600", "900"]:
            gc = OxmlElement("w:gridCol"); gc.set(qn("w:w"), w); tg.append(gc)

    for idx, rd in enumerate(data_rows):
        tr = OxmlElement("w:tr")
        tr.append(make_cell_oxml(str(idx + 1)))
        tr.append(make_cell_oxml(rd.get("需求单号", "")))
        tr.append(make_cell_oxml(rd.get("工单号", "")))
        tr.append(make_cell_oxml(rd.get("工单内容", "")))
        tr.append(make_cell_oxml(rd.get("报表统计次数", "")))
        tbl.append(tr)

    trt = OxmlElement("w:tr")
    trt.append(make_cell_oxml("合计", grid_span=2, bold=True))
    trt.append(make_cell_oxml(""))
    trt.append(make_cell_oxml(""))
    trt.append(make_cell_oxml(str(trep), bold=True))
    tbl.append(trt)
    print(f"表格: {len(data_rows)} 行 + 合计")

    # Remove old gongdan content between first gongdan H2 and 其他要求
    if gd_body is not None:
        children2 = list(body)
        end_idx = other_h1_body if other_h1_body else len(children2)
        for j in range(end_idx - 1, gd_body - 1, -1):
            body.remove(children2[j])

    # Find insertion point: right before "其他要求"
    insert_before = None
    children3 = list(body)
    for j, child in enumerate(children3):
        if child.tag == qn('w:p'):
            for p in doc.paragraphs:
                if p._element is child and p.style.name == "Heading 1" and "其他要求" in p.text:
                    insert_before = child
                    break
        if insert_before is not None: break

    # Build new gongdan content
    print(f"构建 {len(data_rows)} 个工单章节...")
    new_elems = [mp("", S["Normal"])]

    # Need bold mp variant
    def mp_bold(text, style_id, indent=None):
        p = OxmlElement("w:p"); pr = OxmlElement("w:pPr")
        ps = OxmlElement("w:pStyle"); ps.set(qn("w:val"), style_id); pr.append(ps)
        if indent:
            i = OxmlElement("w:ind"); i.set(qn("w:firstLine"), str(indent)); pr.append(i)
        p.append(pr)
        r = OxmlElement("w:r"); rp = OxmlElement("w:rPr"); rp.append(OxmlElement("w:b")); r.append(rp)
        t = OxmlElement("w:t"); t.text = xml_safe(text); t.set(qn("xml:space"), "preserve")
        r.append(t); p.append(r); return p

    for rd in data_rows:
        gd = rd.get("工单内容", "")
        new_elems.append(mp(xml_safe(gd), S["Heading 2"]))
        new_elems.append(mp("业务描述", S["Heading 3"]))
        new_elems.append(mp(xml_safe(rd.get("业务描述", "")), S["Body Text"], 480))
        ri = rd.get("results") or parse_report_list(rd.get("统计分析结果表清单", "")); nr = len(ri)
        new_elems.append(mp(f"本次工作拟产出{nr}个统计分析结果表。", S["Body Text"]))
        if ri:
            rr = [[str(i + 1), xml_safe(cn), xml_safe(en)] for i, (cn, en) in enumerate(ri)]
            new_elems.append(make_stats_source_table(["序号", "结果中文表名称", "表名"], rr))
        new_elems.append(mp("", S["Body Text"]))
        new_elems.append(mp("业务逻辑", S["Heading 3"]))
        gk = rd.get("工单内容", "")
        if relation_descriptions:
            steps = build_stats_requirement_business_logic_steps(ri, relation_descriptions)
        else:
            biz = bts.get(gk, "")
            steps = summarize_biz_logic(biz) if biz else [("1", "无附件，请手动补充业务逻辑说明。")]
        new_elems.append(make_biz_table("业务逻辑说明", steps))
        new_elems.append(mp("数据加工周期", S["Heading 3"]))
        new_elems.append(mp(f"数据统计分析执行周期：{xml_safe(rd.get('数据统计分析执行周期', ''))}。", S["Body Text"]))
        new_elems.append(mp(f"数据更新要求：{xml_safe(rd.get('数据更新要求', ''))}。", S["Body Text"]))
        new_elems.append(mp(f"数据量对后续运维的特殊要求：{xml_safe(rd.get('数据量对后续运维的特殊要求', ''))}。", S["Body Text"]))

    # Insert new elements at the right position
    if insert_before is not None:
        parent = insert_before.getparent()
        idx = list(parent).index(insert_before)
        for elem in reversed(new_elems):
            parent.insert(idx, elem)
    else:
        for elem in new_elems:
            body.append(elem)

    doc.save(output_path)
    print(f"已保存: {output_path}")
    update_toc_via_com(output_path)
    return output_path


# ================================================================================
# 02-数据统计分析_设计文档 Fill Logic
# ================================================================================

def image_id_from_formula(formula_text):
    match = re.search(r'DISPIMG\("([^"]+)"', str(formula_text or ""))
    return match.group(1) if match else ""


def load_stats_design_usage_data(relation_path):
    if not relation_path:
        raise ValueError("02-数据统计分析_设计文档需要模板同目录的 04-数据统计分析_结果表及使用说明.xlsx")

    wb = openpyxl.load_workbook(relation_path, data_only=False)
    source_map = {}
    if "1、数据源表list" in wb.sheetnames:
        ws = wb["1、数据源表list"]
        headers = [str(ws.cell(1, col).value or "").strip() for col in range(1, ws.max_column + 1)]
        for row in range(2, ws.max_row + 1):
            rd = {header: str(ws.cell(row, idx + 1).value or "").strip() for idx, header in enumerate(headers) if header}
            result = norm_table_name(rd.get("数据融合加工表", ""))
            source = norm_table_name(rd.get("资源信息（表名）", ""))
            if result and source:
                source_map.setdefault(result, []).append(rd)

    relation_map = {}
    if "2、表融合关系" in wb.sheetnames:
        ws = wb["2、表融合关系"]
        current = ""
        for row in range(1, ws.max_row + 1):
            title = ws.cell(row, 1).value
            parsed = _parse_relation_title(title)
            if parsed:
                current = parsed
            for col in range(1, ws.max_column + 1):
                label = str(ws.cell(row, col).value or "").strip()
                if label == "文字描述" and current:
                    relation_map.setdefault(current, {})["description"] = str(ws.cell(row, col + 1).value or "").strip()
                    break
                if label == "数据处理流程图" and current:
                    relation_map.setdefault(current, {})["image_formula"] = str(ws.cell(row, col + 1).value or "").strip()
                    break

    detail_map = {}
    if "4、数据统计分析结果表详情" in wb.sheetnames:
        ws = wb["4、数据统计分析结果表详情"]
        current = ""
        row = 1
        while row <= ws.max_row:
            value_b = str(ws.cell(row, 2).value or "").strip()
            if norm_table_name(value_b).startswith("FUSION_"):
                current = norm_table_name(value_b)
                row += 1
                continue
            headers_here = [str(ws.cell(row, col).value or "").strip() for col in range(1, ws.max_column + 1)]
            if current and "字段中文名" in headers_here and "字段英文名" in headers_here:
                header_map = {header: idx + 1 for idx, header in enumerate(headers_here) if header}
                row += 1
                fields = []
                while row <= ws.max_row:
                    first = str(ws.cell(row, 1).value or "").strip()
                    second = str(ws.cell(row, 2).value or "").strip()
                    if re.match(r"4\.\d+", first) or norm_table_name(second).startswith("FUSION_"):
                        row -= 1
                        break
                    if any(str(ws.cell(row, col).value or "").strip() for col in range(1, ws.max_column + 1)):
                        fields.append({header: str(ws.cell(row, col).value or "").strip() for header, col in header_map.items()})
                    row += 1
                detail_map[current] = fields
            row += 1
    wb.close()

    return source_map, relation_map, detail_map, extract_images_via_cellimages(relation_path)


def load_stats_design_catalog(catalog_path, needed_names=None):
    wb = openpyxl.load_workbook(catalog_path, data_only=True, read_only=True)
    if "关联资源信息" not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb["关联资源信息"]
    headers = [str(ws.cell(1, col).value or "").strip() for col in range(1, ws.max_column + 1)]
    rows = {}
    for values in ws.iter_rows(min_row=2, values_only=True):
        rd = {header: str(values[idx] or "").strip() for idx, header in enumerate(headers) if header and idx < len(values)}
        code = norm_table_name(rd.get("资源编码", ""))
        if needed_names and code not in needed_names:
            continue
        if code and code not in rows:
            rows[code] = rd
    wb.close()
    return rows


def make_stats_design_entity_row(table_name, source_rd, resource_info, table_type, seq):
    key = norm_table_name(table_name)
    catalog_rd = resource_info.get(key, {})
    if source_rd:
        directory_code = source_rd.get("资源编目（非必填）") or catalog_rd.get("数据目录代码", "")
        resource_name = source_rd.get("资源名称") or catalog_rd.get("资源名称", "")
        source_name = source_rd.get("资源信息（表名）") or key
    else:
        directory_code = catalog_rd.get("数据目录代码", "")
        resource_name = catalog_rd.get("资源名称", "")
        source_name = key
    return [
        str(seq),
        directory_code,
        resource_name,
        source_name,
        table_type,
        catalog_rd.get("业务数据更新周期", ""),
    ]


def find_heading_paragraph(doc, style_name, contains):
    for paragraph in doc.paragraphs:
        if paragraph.style.name == style_name and contains in paragraph.text:
            return paragraph
    return None


def remove_content_after_heading(body, heading_paragraph):
    found = False
    for child in list(body):
        if child is heading_paragraph._element:
            found = True
            continue
        if found and child.tag != qn("w:sectPr"):
            body.remove(child)


def append_body_element(body, element):
    sect_pr = None
    for child in list(body):
        if child.tag == qn("w:sectPr"):
            sect_pr = child
            break
    if sect_pr is not None:
        body.insert(list(body).index(sect_pr), element)
    else:
        body.append(element)


def append_image_paragraph(doc, body, image_bytes, missing_text):
    paragraph = doc.add_paragraph("")
    append_body_element(body, paragraph._element)
    if image_bytes:
        tf = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tf.write(image_bytes)
        tf.close()
        paragraph.add_run().add_picture(tf.name, width=Inches(5.8))
        os.unlink(tf.name)
    else:
        paragraph.text = missing_text
    return paragraph


def fill_stats_design_doc(excel_path, data_rows, template_path, output_path, catalog_path, relation_path=None):
    """填充 02-数据统计分析_设计文档，只生成 Word 文档。"""
    if not catalog_path:
        raise ValueError("02-数据统计分析_设计文档需要 --catalog 参数指定数据目录数据路径")
    if not os.path.exists(catalog_path):
        raise FileNotFoundError(f"数据目录数据文件不存在: {catalog_path}")
    if not relation_path:
        relation_path = resolve_stats_relation_workbook(template_path)
    if not relation_path:
        raise FileNotFoundError("未找到 04-数据统计分析_结果表及使用说明.xlsx，请放在模板同目录或通过 --catalog 传入")

    print(f"数据: {len(data_rows)} 个工单")
    programs = []
    for group in data_rows:
        for result_cn, result_en in group.get("results") or parse_report_list(group.get("统计分析结果表清单", "")):
            if result_en:
                item = dict(group)
                item["result_cn"] = result_cn
                item["result_en"] = norm_table_name(result_en)
                programs.append(item)
    print(f"程序: {len(programs)} 个")

    source_map, relation_map, detail_map, image_bytes = load_stats_design_usage_data(relation_path)
    needed_names = {item["result_en"] for item in programs}
    for result_en in list(needed_names):
        for source in source_map.get(result_en, []):
            source_name = norm_table_name(source.get("资源信息（表名）", ""))
            if source_name:
                needed_names.add(source_name)
    resource_info = load_stats_design_catalog(catalog_path, needed_names)

    doc = Document(template_path)
    body = doc.element.body
    design_h1 = find_heading_paragraph(doc, "Heading 1", "数据统计分析设计")
    if design_h1 is None:
        raise ValueError("模板中未找到「数据统计分析设计」章节")

    source_heading = find_heading_paragraph(doc, "Heading 2", "需求来源")
    if source_heading:
        source_index = next(
            (
                idx for idx, paragraph in enumerate(doc.paragraphs)
                if paragraph.style.name == "Heading 2" and "需求来源" in paragraph.text
            ),
            -1,
        )
        if source_index + 1 < len(doc.paragraphs):
            desc = doc.paragraphs[source_index + 1]
            unique_reqs = len({item.get("需求单号", "") for item in programs if item.get("需求单号", "")})
            unique_orders = len({item.get("工单号", "") for item in programs if item.get("工单号", "")})
            desc.text = f"服务周期内，共有{unique_reqs}张需求单，{unique_orders}张工单涉及{len(programs)}次数据统计分析服务。具体需求单、工单和产出如下表："
        if len(doc.tables) >= 3:
            tbl = doc.tables[2]._tbl
            for tr in tbl.findall(qn("w:tr"))[1:]:
                tbl.remove(tr)
            tbl_pr = tbl.find(qn("w:tblPr"))
            if tbl_pr is None:
                tbl_pr = OxmlElement("w:tblPr")
                tbl.insert(0, tbl_pr)
            for border in tbl_pr.findall(qn("w:tblBorders")):
                tbl_pr.remove(border)
            add_solid_borders(tbl_pr)
            for idx, item in enumerate(programs, start=1):
                tr = OxmlElement("w:tr")
                for value in [
                    str(idx),
                    item.get("需求单号", ""),
                    item.get("工单号", ""),
                    item.get("工单内容", ""),
                    item["result_en"],
                ]:
                    tr.append(make_cell_oxml(xml_safe(value)))
                tbl.append(tr)

    remove_content_after_heading(body, design_h1)

    style = {}
    for local_name, style_name in [("H2", "Heading 2"), ("H3", "Heading 3"), ("H4", "Heading 4"), ("BT", "Body Text"), ("NL", "Normal")]:
        try:
            style[local_name] = doc.styles[style_name].style_id
        except Exception:
            style[local_name] = style_name

    append_body_element(body, mp(
        f"服务单服务周期内，形成了{len(programs)}个数据统计分析程序，各程序分析设计和融合加工的具体过程如下：",
        style["NL"],
    ))

    for item in programs:
        result_cn = item["result_cn"]
        result_en = item["result_en"]
        append_body_element(body, mp(xml_safe(result_cn), style["H2"]))

        append_body_element(body, mp("涉及到的实体表", style["H3"]))
        entity_rows = []
        seq = 1
        seen_sources = set()
        for source in source_map.get(result_en, []):
            table_name = source.get("资源信息（表名）", "")
            key = norm_table_name(table_name)
            if not key or key in seen_sources:
                continue
            seen_sources.add(key)
            entity_rows.append(make_stats_design_entity_row(table_name, source, resource_info, "源表", seq))
            seq += 1
        entity_rows.append(make_stats_design_entity_row(result_en, None, resource_info, "目标表", seq))
        append_body_element(body, make_table_oxml(
            ["序号", "数据目录/编码（如有）", "数据目录中文名称（目录名）", "表名", "表类型", "数据更新周期"],
            entity_rows,
            ["550", "1600", "2600", "2800", "900", "1300"],
        ))

        append_body_element(body, mp("数据统计分析设计", style["H3"]))
        append_body_element(body, mp(xml_safe(f"{result_cn} {result_en}"), style["H4"]))
        fields = []
        for field in detail_map.get(result_en, []):
            fields.append([
                field.get("字段中文名", ""),
                field.get("字段英文名", ""),
                field.get("字段类型", ""),
                "不可为空",
                "唯一",
                field.get("字段注释", "") or "No",
            ])
        if not fields:
            fields = [["未匹配到字段信息", "", "", "", "", "请补充"]]
        append_body_element(body, make_table_oxml(
            ["字段中文名", "字段英文名", "字段类型", "是否为空", "主键/外键", "字段说明"],
            fields,
            ["1700", "1800", "1500", "1100", "1100", "2800"],
        ))

        append_body_element(body, mp("数据处理流程图", style["H3"]))
        img_id = image_id_from_formula(relation_map.get(result_en, {}).get("image_formula", ""))
        append_image_paragraph(doc, body, image_bytes.get(img_id), "未匹配到数据处理流程图，请补充。")

        append_body_element(body, mp("数据加工逻辑", style["H3"]))
        steps = split_logic_description(relation_map.get(result_en, {}).get("description", ""))
        if not steps:
            steps = ["未匹配到数据加工逻辑说明，请手动补充。"]
        append_body_element(body, make_biz_table(
            "数据统计分析加工逻辑说明",
            [(str(idx), step) for idx, step in enumerate(steps, start=1)],
        ))

    doc.save(output_path)
    print(f"已保存: {output_path}")
    update_toc_via_com(output_path)
    return output_path

# Dispatcher
# ══════════════════════════════════════════════════════════════

def fill_stats_result_usage_workbook(excel_path, service_dir, template_path, output_path, catalog_path):
    """Fill 04-数据统计分析_结果表及使用说明 workbook template."""
    if not catalog_path:
        raise ValueError("04-数据统计分析_结果表及使用说明需要 --catalog 参数指定数据目录数据路径")
    if not os.path.exists(catalog_path):
        raise FileNotFoundError(f"数据目录数据文件不存在: {catalog_path}")
    ensure_module("PIL", "pillow")
    builder_path = os.path.join(os.path.dirname(__file__), "build_stats_result_usage_workbook.py")
    spec = importlib.util.spec_from_file_location("build_stats_result_usage_workbook", builder_path)
    module = importlib.util.module_from_spec(spec)
    if spec.loader is None:
        raise RuntimeError(f"无法加载生成脚本: {builder_path}")
    spec.loader.exec_module(module)
    return module.build_stats_result_usage_workbook(
        ledger_path=excel_path,
        service_dir=service_dir,
        template_path=template_path,
        output_path=output_path,
        catalog_path=catalog_path,
    )


def fill_stats_test_pdf(excel_path, service_dir, template_path, output_path):
    """Fill 03-数据统计分析_测试文档 PDF."""
    ensure_module("PIL", "pillow")
    ensure_module("fitz", "pymupdf")
    ensure_module("reportlab")
    builder_path = os.path.join(os.path.dirname(__file__), "build_stats_test_pdf.py")
    spec = importlib.util.spec_from_file_location("build_stats_test_pdf", builder_path)
    module = importlib.util.module_from_spec(spec)
    if spec.loader is None:
        raise RuntimeError(f"无法加载生成脚本: {builder_path}")
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module.build_stats_test_pdf(
        ledger_path=excel_path,
        service_dir=service_dir,
        template_path=template_path,
        output_path=output_path,
    )


def fill_document(excel_path, service_dir, material_type, template_path, output_path, catalog_path=None):
    print(f"台账清单: {excel_path}")
    print(f"筛选条件: 服务目录={service_dir}")
    print(f"材料类型: {material_type}")

    if material_type == "01-数据报表_需求文档":
        data_rows = read_excel(excel_path, service_dir)
        return fill_requirement_doc(data_rows, template_path, output_path)
    elif material_type == "02-数据报表_设计文档":
        data_rows = read_excel(excel_path, service_dir)
        if not catalog_path:
            raise ValueError("02-设计文档需要 --catalog 参数指定数据目录数据路径")
        if not os.path.exists(catalog_path):
            raise FileNotFoundError(f"数据目录数据文件不存在: {catalog_path}")
        return fill_design_doc_full(excel_path, data_rows, template_path, output_path, catalog_path)
    elif material_type == "01-数据统计分析_需求文档":
        data_rows = read_stats_requirement_groups(excel_path, service_dir)
        relation_path = resolve_stats_relation_workbook(template_path, catalog_path)
        return fill_stats_requirement_doc(excel_path, data_rows, template_path, output_path, relation_path)
    elif material_type == "02-数据统计分析_设计文档":
        if not catalog_path:
            raise ValueError("02-数据统计分析_设计文档需要 --catalog 参数指定数据目录数据路径")
        if not os.path.exists(catalog_path):
            raise FileNotFoundError(f"数据目录数据文件不存在: {catalog_path}")
        data_rows = read_stats_requirement_groups(excel_path, service_dir)
        relation_path = resolve_stats_relation_workbook(template_path)
        return fill_stats_design_doc(excel_path, data_rows, template_path, output_path, catalog_path, relation_path)
    elif material_type == "04-数据统计分析_结果表及使用说明":
        return fill_stats_result_usage_workbook(excel_path, service_dir, template_path, output_path, catalog_path)
    elif material_type == "03-数据统计分析_测试文档":
        return fill_stats_test_pdf(excel_path, service_dir, template_path, output_path)
    elif material_type == "03-数据报表_上线记录":
        data_rows = read_excel(excel_path, service_dir)
        return fill_launch_record_doc(excel_path, data_rows, template_path, output_path)
    else:
        raise ValueError(f"不支持的材料类型: {material_type}。当前支持: 01-数据报表_需求文档, 01-数据统计分析_需求文档, 02-数据报表_设计文档, 02-数据统计分析_设计文档, 03-数据报表_上线记录, 03-数据统计分析_测试文档, 04-数据统计分析_结果表及使用说明")


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



from copy import deepcopy
from pathlib import Path
import subprocess
import sys

from docx.oxml import OxmlElement
from docx.oxml.ns import qn


HEADER_BG = "F1F1F1"


def find_heading(document, style_name, exact_text):
    for paragraph in document.paragraphs:
        if paragraph.style.name == style_name and paragraph.text.strip() == exact_text:
            return paragraph
    raise ValueError(f"模板中未找到{exact_text}章节")


def element_text(element):
    return "".join(node.text or "" for node in element.iter(qn("w:t"))).strip()


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


def replace_after(start, elements):
    parent = start.getparent()
    current = start.getnext()
    while current is not None and current.tag != qn("w:sectPr"):
        following = current.getnext()
        parent.remove(current)
        current = following
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
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = str(text or "")
    run.append(text_node)
    paragraph.append(run)
    return paragraph


def _table_borders(properties):
    existing = properties.find(qn("w:tblBorders"))
    if existing is not None:
        properties.remove(existing)
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        border = OxmlElement(f"w:{edge}")
        border.set(qn("w:val"), "single")
        border.set(qn("w:sz"), "4")
        border.set(qn("w:space"), "0")
        border.set(qn("w:color"), "000000")
        borders.append(border)
    properties.append(borders)


def _cell_element(text, width, bold=False, shaded=False):
    cell = OxmlElement("w:tc")
    properties = OxmlElement("w:tcPr")
    cell_width = OxmlElement("w:tcW")
    cell_width.set(qn("w:w"), str(width))
    cell_width.set(qn("w:type"), "dxa")
    properties.append(cell_width)
    vertical = OxmlElement("w:vAlign")
    vertical.set(qn("w:val"), "center")
    properties.append(vertical)
    margins = OxmlElement("w:tcMar")
    for side, value in (("top", "80"), ("left", "100"), ("bottom", "80"), ("right", "100")):
        margin = OxmlElement(f"w:{side}")
        margin.set(qn("w:w"), value)
        margin.set(qn("w:type"), "dxa")
        margins.append(margin)
    properties.append(margins)
    if shaded:
        shading = OxmlElement("w:shd")
        shading.set(qn("w:val"), "clear")
        shading.set(qn("w:fill"), HEADER_BG)
        properties.append(shading)
    cell.append(properties)

    paragraph = OxmlElement("w:p")
    paragraph_properties = OxmlElement("w:pPr")
    alignment = OxmlElement("w:jc")
    alignment.set(qn("w:val"), "center")
    paragraph_properties.append(alignment)
    paragraph.append(paragraph_properties)
    run = OxmlElement("w:r")
    run_properties = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    fonts.set(qn("w:ascii"), "Times New Roman")
    fonts.set(qn("w:hAnsi"), "Times New Roman")
    fonts.set(qn("w:eastAsia"), "宋体")
    run_properties.append(fonts)
    size = OxmlElement("w:sz")
    size.set(qn("w:val"), "21")
    run_properties.append(size)
    if bold:
        run_properties.append(OxmlElement("w:b"))
    run.append(run_properties)
    text_node = OxmlElement("w:t")
    text_node.set(qn("xml:space"), "preserve")
    text_node.text = str(text or "")
    run.append(text_node)
    paragraph.append(run)
    cell.append(paragraph)
    return cell


def table_element(headers, rows, column_widths=None):
    headers = [str(value or "") for value in headers]
    if not headers:
        raise ValueError("参数表缺少表头")
    column_count = len(headers)
    if column_widths is None:
        base = 9000 // column_count
        column_widths = [base] * column_count
        column_widths[-1] += 9000 - sum(column_widths)

    table = OxmlElement("w:tbl")
    properties = OxmlElement("w:tblPr")
    style = OxmlElement("w:tblStyle")
    style.set(qn("w:val"), "TableGrid")
    properties.append(style)
    width = OxmlElement("w:tblW")
    width.set(qn("w:w"), "5000")
    width.set(qn("w:type"), "pct")
    properties.append(width)
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    properties.append(layout)
    _table_borders(properties)
    table.append(properties)

    grid = OxmlElement("w:tblGrid")
    for value in column_widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(value))
        grid.append(column)
    table.append(grid)

    header_row = OxmlElement("w:tr")
    row_properties = OxmlElement("w:trPr")
    row_properties.append(OxmlElement("w:tblHeader"))
    header_row.append(row_properties)
    for index, header in enumerate(headers):
        header_row.append(_cell_element(header, column_widths[index], bold=True, shaded=True))
    table.append(header_row)

    for source_row in rows:
        values = [str(value or "") for value in source_row[:column_count]]
        values.extend([""] * (column_count - len(values)))
        row = OxmlElement("w:tr")
        for index, value in enumerate(values):
            row.append(_cell_element(value, column_widths[index]))
        table.append(row)
    return table


def clone_element(element):
    return deepcopy(element)


def update_toc_via_com(doc_path):
    if sys.platform != "win32":
        return False
    escaped_path = str(Path(doc_path).resolve()).replace("'", "''")
    script = f"""
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{escaped_path}')
    foreach ($field in $doc.Fields) {{ $field.Update() | Out-Null }}
    foreach ($toc in $doc.TablesOfContents) {{ $toc.Update() }}
    $doc.Save()
    Write-Output 'UPDATED'
}} finally {{
    if ($doc) {{ $doc.Close() }}
    if ($word) {{ $word.Quit() }}
}}
"""
    result = subprocess.run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        timeout=120,
    )
    return result.returncode == 0 and "UPDATED" in result.stdout

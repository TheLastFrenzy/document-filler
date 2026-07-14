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


def clone_paragraph_with_text(prototype, text):
    paragraph = OxmlElement("w:p")
    properties = prototype.find(qn("w:pPr"))
    if properties is not None:
        paragraph.append(deepcopy(properties))

    prototype_run = prototype.find(qn("w:r"))
    if prototype_run is None:
        prototype_run = prototype.find(".//" + qn("w:r"))
    run = OxmlElement("w:r")
    if prototype_run is not None:
        run_properties = prototype_run.find(qn("w:rPr"))
        if run_properties is not None:
            run.append(deepcopy(run_properties))
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


def _cell_element(text, width, bold=False, shaded=False, align=None):
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
    if align:
        paragraph_properties = OxmlElement("w:pPr")
        alignment = OxmlElement("w:jc")
        alignment.set(qn("w:val"), align)
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
    size.set(qn("w:val"), "20")
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
    else:
        source_total = sum(int(value) for value in column_widths)
        if source_total <= 0 or len(column_widths) != column_count:
            raise ValueError("参数表列宽与表头不匹配")
        column_widths = [round(int(value) * 9000 / source_total) for value in column_widths]
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
        header_row.append(_cell_element(header, column_widths[index], bold=True, shaded=True, align="center"))
    table.append(header_row)

    for source_row in rows:
        values = [str(value or "") for value in source_row[:column_count]]
        values.extend([""] * (column_count - len(values)))
        row = OxmlElement("w:tr")
        for index, value in enumerate(values):
            row.append(_cell_element(value, column_widths[index]))
        table.append(row)
    return table


def _cell_grid_span(cell):
    properties = cell.find(qn("w:tcPr"))
    span = properties.find(qn("w:gridSpan")) if properties is not None else None
    return int(span.get(qn("w:val"))) if span is not None else 1


def _resize_unmerged_row(row, column_count):
    cells = list(row.findall(qn("w:tc")))
    if any(_cell_grid_span(cell) != 1 for cell in cells):
        return
    while len(cells) > column_count:
        row.remove(cells.pop())
    prototype = cells[-1] if cells else None
    while len(cells) < column_count:
        cell = deepcopy(prototype) if prototype is not None else OxmlElement("w:tc")
        row.append(cell)
        cells.append(cell)


def _replace_cell_text(cell, text):
    prototype_paragraph = cell.find(qn("w:p"))
    for child in list(cell):
        if child.tag != qn("w:tcPr"):
            cell.remove(child)
    if prototype_paragraph is None:
        prototype_paragraph = OxmlElement("w:p")
    cell.append(clone_paragraph_with_text(prototype_paragraph, text))


def _set_cell_width(cell, width, total_width):
    properties = cell.find(qn("w:tcPr"))
    if properties is None:
        properties = OxmlElement("w:tcPr")
        cell.insert(0, properties)
    cell_width = properties.find(qn("w:tcW"))
    if cell_width is None:
        cell_width = OxmlElement("w:tcW")
        properties.insert(0, cell_width)
    width_type = cell_width.get(qn("w:type"), "pct")
    if width_type == "pct":
        cell_width.set(qn("w:w"), str(round(width * 5000 / total_width)))
    else:
        cell_width.set(qn("w:w"), str(width))


def _clone_table_row(prototype, values, widths, adjust_widths):
    row = deepcopy(prototype)
    logical_columns = len(values)
    _resize_unmerged_row(row, logical_columns)
    logical_index = 0
    for cell in row.findall(qn("w:tc")):
        span = _cell_grid_span(cell)
        value = values[logical_index] if logical_index < logical_columns else ""
        _replace_cell_text(cell, value)
        if adjust_widths and logical_index < len(widths):
            cell_width = sum(widths[logical_index : logical_index + span])
            _set_cell_width(cell, cell_width, sum(widths))
        logical_index += span
    return row


def _set_repeat_table_header(row):
    properties = row.find(qn("w:trPr"))
    if properties is None:
        properties = OxmlElement("w:trPr")
        row.insert(0, properties)
    if properties.find(qn("w:tblHeader")) is None:
        properties.append(OxmlElement("w:tblHeader"))


def clone_table_with_data(prototype, headers, rows, column_widths=None, footer=None):
    headers = [str(value or "") for value in headers]
    if not headers:
        raise ValueError("参数表缺少表头")
    column_count = len(headers)
    if column_widths is not None and len(column_widths) != column_count:
        raise ValueError("参数表列宽与表头不匹配")

    table = deepcopy(prototype)
    prototype_rows = list(table.findall(qn("w:tr")))
    if not prototype_rows:
        raise ValueError("模板表格缺少原型行")
    header_prototype = prototype_rows[0]
    data_prototype = prototype_rows[1] if len(prototype_rows) > 1 else prototype_rows[0]
    footer_prototype = prototype_rows[-1] if len(prototype_rows) > 2 else data_prototype
    prototype_column_count = sum(
        _cell_grid_span(cell) for cell in header_prototype.findall(qn("w:tc"))
    )

    grid = table.find(qn("w:tblGrid"))
    source_grid = [] if grid is None else [
        int(column.get(qn("w:w"), "0")) for column in grid.findall(qn("w:gridCol"))
    ]
    total_width = sum(source_grid) or 9000
    if column_widths is None:
        if len(source_grid) == column_count and all(source_grid):
            widths = source_grid
        else:
            widths = [total_width // column_count] * column_count
            widths[-1] += total_width - sum(widths)
    else:
        source_total = sum(int(value) for value in column_widths)
        if source_total <= 0:
            raise ValueError("参数表列宽必须为正数")
        widths = [round(int(value) * total_width / source_total) for value in column_widths]
        widths[-1] += total_width - sum(widths)
    adjust_widths = column_widths is not None or prototype_column_count != column_count

    if grid is None:
        grid = OxmlElement("w:tblGrid")
        properties = table.find(qn("w:tblPr"))
        table.insert(table.index(properties) + 1 if properties is not None else 0, grid)
    for column in list(grid):
        grid.remove(column)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        grid.append(column)

    for row in prototype_rows:
        table.remove(row)
    header_row = _clone_table_row(header_prototype, headers, widths, adjust_widths)
    _set_repeat_table_header(header_row)
    table.append(header_row)
    for source_row in rows:
        values = [str(value or "") for value in source_row[:column_count]]
        values.extend([""] * (column_count - len(values)))
        table.append(_clone_table_row(data_prototype, values, widths, adjust_widths))
    if footer is not None:
        values = [str(value or "") for value in footer[:column_count]]
        values.extend([""] * (column_count - len(values)))
        table.append(_clone_table_row(footer_prototype, values, widths, adjust_widths))
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

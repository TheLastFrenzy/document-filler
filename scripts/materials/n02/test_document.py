import datetime as dt
import os
import re
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path

from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.oxml.ns import qn
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.table import Table

from materials.shared.ledger_sheet import select_ledger_sheet
from materials.shared.office_word import save_word_document
from materials.shared.word_sections import (
    clone_paragraph_with_text,
    clone_table_with_data,
    find_heading,
    replace_between,
)


TASK_NAME_BOX = (54, 29, 794, 76)
STRATEGY_BOX = (930, 29, 1200, 76)
SUCCESS_BOX = (1211, 29, 1300, 76)
LAUNCH_BOX = (1490, 29, 1665, 76)
MIN_FONT_SIZE = 12


@dataclass(frozen=True)
class N02Task:
    task_name: str
    strategy: str
    strategy_text: str
    success_failure: str
    launch_time: object
    launch_time_text: str
    test_time_text: str
    source_department: str
    table_name: str


def _safe_name(value):
    text = re.sub(r"[\\/:*?\"<>|]+", "_", str(value or "").strip())
    text = re.sub(r"\s+", "_", text)
    return text or "task"


def _load_font(size):
    windir = Path(os.environ.get("WINDIR", r"C:\Windows"))
    fonts = windir / "Fonts"
    candidates = [
        "msyh.ttc",
        "msyhbd.ttc",
        "simhei.ttf",
        "simhei.ttf",
        "simsun.ttc",
        "arial.ttf",
    ]
    for name in candidates:
        font_path = fonts / name
        if font_path.exists():
            try:
                return ImageFont.truetype(str(font_path), size=size)
            except OSError:
                continue
    return ImageFont.load_default()


def _fit_font(draw, text, max_width, start_size=24, min_size=MIN_FONT_SIZE):
    for size in range(start_size, min_size - 1, -1):
        font = _load_font(size)
        bbox = draw.textbbox((0, 0), text, font=font)
        if bbox[2] - bbox[0] <= max_width:
            return font
    return _load_font(min_size)


def _centered_position(draw, box, text, font, left_padding=0):
    left, top, right, bottom = box
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]
    x = left + left_padding
    if left_padding == 0:
        x = left + max(0, (right - left - text_width) // 2)
    y = top + max(0, (bottom - top - text_height) // 2) - bbox[1]
    return x, y


def normalize_execution_strategy(value):
    text = str(value or "").strip()
    if not text:
        return ""
    daily = re.fullmatch(r"(\d{1,2}):(\d{2})", text)
    if daily:
        return f"每天{int(daily.group(1)):02d}点"
    monthly = re.fullmatch(r"(\d{1,2})\s*,\s*(\d{1,2}):(\d{2})", text)
    if monthly:
        return f"每月{int(monthly.group(1))}号的{int(monthly.group(2)):02d}点"
    period = re.fullmatch(r"(\d+)(hour|min)", text, flags=re.IGNORECASE)
    if period:
        return f"每{period.group(1)}{period.group(2).lower()}"
    return text


def _format_launch_time_text(value):
    if isinstance(value, dt.datetime):
        return value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, dt.date):
        return dt.datetime.combine(value, dt.time()).strftime("%Y-%m-%d %H:%M:%S")
    text = str(value or "").strip()
    if not text:
        return ""
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", text):
        return f"{text} 00:00:00"
    return text


def _format_test_time(value):
    if isinstance(value, dt.datetime):
        return value.strftime("%Y%m%d")
    if isinstance(value, dt.date):
        return value.strftime("%Y%m%d")
    text = str(value or "").strip()
    if not text:
        return ""
    try:
        return dt.datetime.fromisoformat(text.replace("/", "-")).strftime("%Y%m%d")
    except ValueError:
        return text[:10].replace("-", "")[:8]


def _select_value(row, index):
    if index is None or index >= len(row):
        return ""
    return row[index]


def read_n02_tasks(excel_path):
    workbook = load_workbook(excel_path, data_only=True, read_only=False)
    sheet = select_ledger_sheet(workbook, required_headers=("任务名称", "执行策略", "成功/失败", "上线时间"))
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]
    header_index = {header: idx for idx, header in enumerate(headers)}
    required = ("任务名称", "执行策略", "成功/失败", "上线时间", "来源委办", "表名")
    missing = [name for name in required if name not in header_index]
    if missing:
        raise ValueError(f"台账清单缺少必要列: {', '.join(missing)}")

    tasks = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        task_name = str(_select_value(row, header_index["任务名称"]) or "").strip()
        if not task_name:
            continue
        raw_strategy = _select_value(row, header_index["执行策略"])
        launch_time = _select_value(row, header_index["上线时间"])
        tasks.append(
            N02Task(
                task_name=task_name,
                strategy=str(raw_strategy or "").strip(),
                strategy_text=normalize_execution_strategy(raw_strategy),
                success_failure=str(_select_value(row, header_index["成功/失败"]) or "").strip(),
                launch_time=launch_time,
                launch_time_text=_format_launch_time_text(launch_time),
                test_time_text=_format_test_time(launch_time),
                source_department=str(_select_value(row, header_index["来源委办"]) or "").strip(),
                table_name=str(_select_value(row, header_index["表名"]) or "").strip(),
            )
        )

    workbook.close()
    if not tasks:
        raise ValueError(f"未找到匹配数据: 台账清单={excel_path}")
    return tasks


def _render_text(draw, box, text, *, fill, start_size, min_size=MIN_FONT_SIZE, align="center", left_padding=0):
    text = str(text or "")
    if not text:
        return
    font = _fit_font(draw, text, box[2] - box[0] - 8, start_size=start_size, min_size=min_size)
    x, y = _centered_position(draw, box, text, font, left_padding=left_padding if align == "left" else 0)
    draw.text((x, y), text, fill=fill, font=font)


def render_task_screenshot(task, template_image_path, output_path):
    image = Image.open(template_image_path).convert("RGB")
    draw = ImageDraw.Draw(image)
    background = image.getpixel((10, 10))
    for box in (TASK_NAME_BOX, STRATEGY_BOX, SUCCESS_BOX, LAUNCH_BOX):
        draw.rectangle(box, fill=background)

    _render_text(draw, TASK_NAME_BOX, task.task_name, fill=(33, 102, 173), start_size=24, align="left", left_padding=8)
    _render_text(draw, STRATEGY_BOX, task.strategy_text, fill=(66, 66, 66), start_size=18)
    _render_text(draw, SUCCESS_BOX, task.success_failure, fill=(192, 0, 0), start_size=22)
    _render_text(draw, LAUNCH_BOX, task.launch_time_text, fill=(66, 66, 66), start_size=20)

    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    image.save(output)
    return str(output)


def _set_cell_text(cell, text):
    prototype = deepcopy(cell.paragraphs[0]._p)
    tc = cell._tc
    for child in list(tc):
        if child.tag != qn("w:tcPr"):
            tc.remove(child)
    tc.append(clone_paragraph_with_text(prototype, text))


def _replace_image(cell, image_path, width):
    if not cell.paragraphs:
        paragraph = cell.add_paragraph()
    else:
        paragraph = cell.paragraphs[-1]
    for run in list(paragraph.runs):
        paragraph._p.remove(run._r)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    paragraph.add_run().add_picture(str(image_path), width=width)


def build_n02_test_document(excel_path, service_dir, template_path, output_path):
    tasks = read_n02_tasks(excel_path)
    document = Document(template_path)

    test_range_heading = find_heading(document, "Heading 1", "测试范围")
    test_environment_heading = find_heading(document, "Heading 1", "测试环境")
    test_case_heading = find_heading(document, "Heading 1", "测试案例")
    test_conclusion_heading = find_heading(document, "Heading 1", "测试结论")
    attachment_heading = find_heading(document, "Heading 1", "附件：测试结果截图")

    scope_elements = []
    current = test_range_heading._p.getnext()
    while current is not None and current is not test_environment_heading._p:
        scope_elements.append(current)
        current = current.getnext()
    scope_paragraph = next((item for item in scope_elements if item.tag.endswith("}p")), None)
    scope_table = next((item for item in scope_elements if item.tag.endswith("}tbl")), None)
    if scope_paragraph is None or scope_table is None:
        raise ValueError("模板中未找到测试范围的段落或表格原型")

    case_elements = []
    current = test_case_heading._p.getnext()
    while current is not None and current is not test_conclusion_heading._p:
        case_elements.append(current)
        current = current.getnext()
    case_title_prototype = next((item for item in case_elements if item.tag.endswith("}p")), None)
    case_table_prototype = next((item for item in case_elements if item.tag.endswith("}tbl")), None)
    conclusion_elements = []
    current = test_conclusion_heading._p.getnext()
    while current is not None and current is not attachment_heading._p:
        conclusion_elements.append(current)
        current = current.getnext()
    conclusion_paragraph = next((item for item in conclusion_elements if item.tag.endswith("}p")), None)
    if case_title_prototype is None or case_table_prototype is None or conclusion_paragraph is None:
        raise ValueError("模板中未找到测试案例或测试结论原型")

    template_image = Path(template_path).with_name("图片1.png")
    if not template_image.exists():
        raise FileNotFoundError(f"未找到截图模板图片: {template_image}")

    screenshot_width = document.inline_shapes[0].width if document.inline_shapes else None
    if screenshot_width is None:
        raise ValueError("模板中未找到截图原型")

    with tempfile.TemporaryDirectory(prefix="document_filler_n02_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        scope_text = f"针对服务期内新增的{len(tasks)}个归集任务进行测试，涉及{len({task.source_department for task in tasks})}个委办。"
        conclusion_text = f"通过对上述{len(tasks)}个归集任务测试，任务均能正常运行，测试通过。"

        scope_table_data = [[str(index), task.task_name, task.source_department, task.table_name] for index, task in enumerate(tasks, start=1)]
        scope_table_element = clone_table_with_data(
            scope_table,
            ["序号", "任务名", "来源委办", "表名"],
            scope_table_data,
        )

        case_elements_out = []
        for index, task in enumerate(tasks, start=1):
            screenshot_path = temp_dir_path / f"{index:02d}_{_safe_name(task.task_name)}.png"
            render_task_screenshot(task, template_image, screenshot_path)
            case_title = clone_paragraph_with_text(case_title_prototype, task.task_name)
            case_table_xml = deepcopy(case_table_prototype)
            case_table = Table(case_table_xml, document)
            _set_cell_text(case_table.rows[7].cells[1], "栾希旺")
            _set_cell_text(case_table.rows[7].cells[3], task.test_time_text)
            _replace_image(case_table.rows[6].cells[1], screenshot_path, screenshot_width)
            case_elements_out.extend([case_title, case_table_xml])

        replace_between(
            test_range_heading._p,
            test_environment_heading._p,
            [clone_paragraph_with_text(scope_paragraph, scope_text), scope_table_element],
        )
        replace_between(
            test_case_heading._p,
            test_conclusion_heading._p,
            case_elements_out,
        )
        replace_between(
            test_conclusion_heading._p,
            attachment_heading._p,
            [clone_paragraph_with_text(conclusion_paragraph, conclusion_text)],
        )

        work_dir = temp_dir_path / "work"
        return save_word_document(document, output_path, work_dir)

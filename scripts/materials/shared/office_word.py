import subprocess
from pathlib import Path

from materials.shared.word_sections import update_toc_via_com


def escape_powershell_string(value):
    return str(value).replace("'", "''")


def convert_legacy_doc_template(template_path, work_dir, runner=None):
    template = Path(template_path)
    if template.suffix.lower() == ".docx":
        return template
    if template.suffix.lower() != ".doc":
        raise ValueError(f"API接口测试报告模板仅支持 .doc 或 .docx: {template}")

    Path(work_dir).mkdir(parents=True, exist_ok=True)
    output = Path(work_dir) / f"{template.stem}.docx"
    script = f"""
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{escape_powershell_string(str(template.resolve()))}')
    $doc.SaveAs([ref] '{escape_powershell_string(str(output.resolve()))}', [ref] 16)
    Write-Output 'CONVERTED'
}} finally {{
    if ($doc) {{ $doc.Close($false) }}
    if ($word) {{ $word.Quit() }}
}}
"""
    run = runner or subprocess.run
    result = run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        errors="replace",
        timeout=120,
    )
    if not output.exists() or output.stat().st_size == 0:
        raise RuntimeError(
            f"无法将旧版 .doc 模板转换为 .docx: {template}; "
            f"stdout={result.stdout.strip()} stderr={result.stderr.strip()}"
        )
    return output


def convert_docx_to_legacy_doc(docx_path, output_path, runner=None):
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    script = f"""
$word = $null
$doc = $null
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{escape_powershell_string(str(Path(docx_path).resolve()))}')
    foreach ($field in $doc.Fields) {{ $field.Update() | Out-Null }}
    foreach ($toc in $doc.TablesOfContents) {{ $toc.Update() }}
    $doc.SaveAs([ref] '{escape_powershell_string(str(output.resolve()))}', [ref] 0)
    Write-Output 'CONVERTED'
}} finally {{
    if ($doc) {{ $doc.Close($false) }}
    if ($word) {{ $word.Quit() }}
}}
"""
    run = runner or subprocess.run
    result = run(
        ["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        errors="replace",
        timeout=120,
    )
    if not output.exists() or output.stat().st_size == 0:
        raise RuntimeError(
            f"无法将 .docx 输出转换为旧版 .doc: {output}; "
            f"stdout={result.stdout.strip()} stderr={result.stderr.strip()}"
        )
    return output


def save_word_document(
    document,
    output_path,
    work_dir,
    legacy_converter=None,
    toc_updater=None,
):
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    if output.suffix.lower() == ".doc":
        Path(work_dir).mkdir(parents=True, exist_ok=True)
        docx_output = Path(work_dir) / f"{output.stem}.docx"
        document.save(docx_output)
        convert = legacy_converter or convert_docx_to_legacy_doc
        convert(docx_output, output)
        return str(output)
    document.save(output)
    update = toc_updater or update_toc_via_com
    update(output)
    return str(output)

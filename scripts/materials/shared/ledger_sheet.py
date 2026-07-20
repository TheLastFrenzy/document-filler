def select_ledger_sheet(workbook, required_headers=("服务目录",)):
    if "台账清单" in workbook.sheetnames:
        return workbook["台账清单"]

    required = set(required_headers)
    for sheet in workbook.worksheets:
        headers = {
            str(sheet.cell(1, column).value or "").strip()
            for column in range(1, sheet.max_column + 1)
        }
        if required.issubset(headers):
            return sheet

    return workbook.active

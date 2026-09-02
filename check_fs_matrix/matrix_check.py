"""
Checks whether an Excel file contains a sheet that is some version of the
"FS Matrix" sheet - the actual sheet name varies across submissions (e.g.
"FI FS Matrix", "FSマトリクス"), so every known variant is kept in one list
here rather than checking a single hardcoded name.
"""

import os

import openpyxl

# Known variants of the FS Matrix sheet name - add more here as new formats
# show up in submitted files.
MATRIX_SHEET_NAME_VARIANTS = [
    "FI FSﾏﾄﾘｸｽ",
    "FI FSﾏﾄﾘｸｽ",
    "FI_FSﾏﾄﾘｸｽ",
    "FI FSmatrix",
    "FI FSマトリクス",
    "FI_FSマトリクス",
    "ＦＩ FSMATRIX",
    "FI Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "MOT FSﾏﾄﾘｸｽ",
    "MOT FSﾏﾄﾘｸｽ",
    "MOT_FSﾏﾄﾘｸｽ",
    "MOT FSmatrix",
    "MOT FSマトリクス",
    "MOT FSMATRIX",
    "MOT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "BAT FSﾏﾄﾘｸｽ",
    "BAT FSﾏﾄﾘｸｽ",
    "BAT_FSﾏﾄﾘｸｽ",
    "BAT FSmatrix",
    "BAT FSマトリクス",
    "BAT FSMATRIX",
    "BAT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "EVCC FSﾏﾄﾘｸｽ",
    "EVCC FSﾏﾄﾘｸｽ",
    "EVCC_FSﾏﾄﾘｸｽ",
    "EVCC FSmatrix",
    "EVCC FSマトリクス",
    "EVCC FSMATRIX",
    "EVCC Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "ZEV2IN1 FSﾏﾄﾘｸｽ",
    "ZEV2IN1 FSﾏﾄﾘｸｽ",
    "ZEV2IN1_FSﾏﾄﾘｸｽ",
    "ZEV2IN1 FSmatrix",
    "ZEV2IN1 FSマトリクス",
    "ZEV2IN1 FSMATRIX",
    "ZEV2IN1 Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "AREAFRONT FSﾏﾄﾘｸｽ",
    "AREAFRONT FSﾏﾄﾘｸｽ",
    "AREAFRONT_FSﾏﾄﾘｸｽ",
    "AREAFRONT FSmatrix",
    "AREAFRONT FSマトリクス",
    "AREAFRONT FSMATRIX",
    "AREAFRONT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "CORE FSﾏﾄﾘｸｽ",
    "CORE FSﾏﾄﾘｸｽ",
    "CORE_FSﾏﾄﾘｸｽ",
    "CORE FSmatrix",
    "CORE FSマトリクス",
    "CORE FSMATRIX",
    "CORE Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "ICM FSﾏﾄﾘｸｽ",
    "ICM FSﾏﾄﾘｸｽ",
    "ICM_FSﾏﾄﾘｸｽ",
    "ICM FSmatrix",
    "ICM FSマトリクス",
    "ICM_FSマトリクス",
    "ＩＣＭ FSMATRIX",
    "ICM Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",

]


BUNBAN_LABEL = "部番"
DIAGNOSTIC_CODE_LABEL = "診断識別コード"
OK_NG_SYMBOLS = ("○", "×", "（×）","◯","●","〇","０","0","Ｏ","ｏ","O","o","Ｘ","ｘ","X","x","-","ー","－","‐")


def _find_cell(ws, target_value, exact=True):
    """Returns the first cell matching target_value, or None if not found.
    exact=True requires the cell's value to equal target_value exactly;
    exact=False matches a cell whose (string) value merely contains
    target_value as a substring - e.g. a "部番NO" or "部番／型式" header cell,
    not just a cell whose whole value is exactly "部番"."""
    for row in ws.iter_rows():
        for cell in row:
            if exact:
                if cell.value == target_value:
                    return cell
            elif isinstance(cell.value, str) and target_value in cell.value:
                return cell
    return None


def _starts_with_number(value):
    if isinstance(value, (int, float)):
        return True
    if isinstance(value, str) and value.strip() and value.strip()[0].isdigit():
        return True
    return False


def _merged_range_containing(ws, cell):
    """Returns the merged cell range containing `cell`, or None if it isn't
    part of one. Needed because openpyxl only stores a value in a merged
    range's top-left cell - every other cell in the range reads back as
    None, which would otherwise look like a blank row right after the
    header rather than the header itself spanning multiple rows."""
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return merged_range
    return None


def check_fs_matrix_symbols(local_path, sheet_name):
    """Returns (result, message). Dynamically locates every column in the
    "部番" row whose value starts with a number, and the row right after
    "診断識別コード" (or right after that header's merged range, if it spans
    more than one row) as the starting row. The table's last row is found by
    walking down the "診断識別コード" column itself until the first blank
    cell - the row just before that is the shared last row used for every
    numeric column, not a separate stopping point per column. Every cell in
    that shared row range must be 〇 or ×; a blank cell within that range
    counts as an error, not as the end of the table."""
    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return "NG", f"ファイルを開けない: {e}"

    if sheet_name not in wb.sheetnames:
        return "NG", f"「{sheet_name}」が見つかりません"

    ws = wb[sheet_name]

    bunban_cell = _find_cell(ws, BUNBAN_LABEL, exact=False)
    if bunban_cell is None:
        return "NG", f"「{BUNBAN_LABEL}」が見つかりません"

    numeric_columns = [cell.column for cell in ws[bunban_cell.row] if _starts_with_number(cell.value)]
    if not numeric_columns:
        return "NG", f"「{BUNBAN_LABEL}」行に番号で始まるセルが見つかりません"

    diagnostic_cell = _find_cell(ws, DIAGNOSTIC_CODE_LABEL)
    if diagnostic_cell is None:
        return "NG", f"「{DIAGNOSTIC_CODE_LABEL}」が見つかりません"

    # If the header cell is merged across multiple rows (e.g. spanning the
    # row below it too), data starts right after the merge ends, not just
    # one row below wherever the header's own anchor cell happens to be.
    diagnostic_merge = _merged_range_containing(ws, diagnostic_cell)
    header_end_row = diagnostic_merge.max_row if diagnostic_merge else diagnostic_cell.row
    start_row = header_end_row + 1

    # Walk down the 診断識別コード column itself to find the shared last row -
    # the row right before the first blank cell in that column.
    last_row = start_row - 1
    row_num = start_row
    while ws.cell(row=row_num, column=diagnostic_cell.column).value is not None:
        last_row = row_num
        row_num += 1

    if last_row < start_row:
        return "NG", f"「{DIAGNOSTIC_CODE_LABEL}」の下にデータがありません"

    for col in numeric_columns:
        for row_num in range(start_row, last_row + 1):
            cell = ws.cell(row=row_num, column=col)
            if cell.value not in OK_NG_SYMBOLS:
                return "NG", f"{cell.coordinate}に〇×ではない値があります"

    return "OK", f"「{sheet_name}」の部番・診断識別コード表にデータ不備は有りません"


def check_matrix_sheet(local_path):
    """Returns (result, message). Finds whichever MATRIX_SHEET_NAME_VARIANTS
    name exists as a sheet in local_path, reports that as found, then
    validates its symbol table via check_fs_matrix_symbols. If both pass,
    the OK message combines the "found" message with check_fs_matrix_symbols'
    own message. NG if the file can't be opened, none of the variants are
    found, or the symbol table check fails."""
    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return "NG", f"ファイルを開けない: {e}"

    for variant in MATRIX_SHEET_NAME_VARIANTS:
        if variant in wb.sheetnames:
            found_message = 'FSマトリクスファイルにツールが想定している"FI FSMatrix"シートがある。'
            symbol_result, symbol_message = check_fs_matrix_symbols(local_path, variant)
            if symbol_result != "OK":
                return "NG", symbol_message
            return "OK", f"{found_message}\n{symbol_message}"

    return "NG", 'FSマトリクスファイルにツールが想定している"FI FSMatrix"シートが見つからない。'


def check_matrix_sheet_for_all_files(local_paths):
    """Runs check_matrix_sheet on every given local file path - for a column
    that may have more than one file attached, every one of them needs a
    matching sheet. Returns one combined (result, message): OK only if every
    file passed; otherwise NG naming whichever file failed."""
    per_file_results = [(path, check_matrix_sheet(path)) for path in local_paths]

    for path, (result, message) in per_file_results:
        if result != "OK":
            return "NG", f"{os.path.basename(path)}: {message}"

    return "OK", " / ".join(f"{os.path.basename(path)}: {message}" for path, (result, message) in per_file_results)

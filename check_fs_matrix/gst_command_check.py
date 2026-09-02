"""
Checks GST Command Support files attached to commentCommandSupport for the
existence of a "SID一覧" sheet.

Different from the FS Matrix check (matrix_check.py): a missing attachment
is NOT an error here. If nothing is attached, check_gst_command_files()
simply returns OK with no message - the caller doesn't need to special-case
"no attachment" separately, it can always call this function and let it
decide there is nothing to check. When there IS at least one attachment,
every file whose name contains "CommandSupportGST" and has an Excel
extension gets checked for SID一覧 - files that don't match that pattern are
ignored, not flagged.
"""

import os

import openpyxl

SID_LIST_SHEET_NAME = "SID一覧"

# Row labels to check - each found label defines its own row; the columns
# checked within that row start two columns after the label ends (the label
# itself may be a single cell or merged across several columns), skipping
# one gap column right after it.
J1979_LABELS = ("SAE J1979", "SAE J1979-2", "SAE J1979-3")
# The column just before this one marks the end of the checked range, for
# every J1979_LABELS row.
J1979_APPLY_LABEL = "J1979を適用する場合"
SID_OK_NG_SYMBOLS = ("○", "×", "（×）","◯","●","〇","０","0","Ｏ","ｏ","O","o","Ｘ","ｘ","X","x","-","ー","－","‐")


def is_command_support_gst_file(name):
    return "CommandSupportGST" in name and name.lower().endswith((".xlsx", ".xls"))


def _find_cell(ws, target_value):
    """Returns the first cell whose value equals target_value, or None if
    not found."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == target_value:
                return cell
    return None


def _merged_range_containing(ws, cell):
    """Returns the merged cell range containing `cell`, or None if it isn't
    part of one. Needed because a J1979_LABELS cell (e.g. "SAE J1979") may
    itself be merged across several columns - the actual label's rightmost
    column is that merge's max_col, not the anchor cell's own column."""
    for merged_range in ws.merged_cells.ranges:
        if cell.coordinate in merged_range:
            return merged_range
    return None


def check_sid_list_symbols(local_path, sheet_name, file_name):
    """Returns (result, message). For each of J1979_LABELS found as a cell
    in sheet_name, checks that label's row through the column right before
    J1979_APPLY_LABEL's column - every cell in that range must be 〇 or ×;
    anything else is an error. NG if J1979_APPLY_LABEL or none of
    J1979_LABELS can be found at all.

    The checked range starts two columns after the label's own rightmost
    column: one column is the label itself (or, if the label cell is merged
    across several columns, the whole merge) and the next column after that
    is a skipped gap column - data starts only after both."""
    wb = openpyxl.load_workbook(local_path, data_only=True)
    ws = wb[sheet_name]

    apply_cell = _find_cell(ws, J1979_APPLY_LABEL)
    if apply_cell is None:
        return "NG", f"'{file_name}'に「{J1979_APPLY_LABEL}」が見つかりません"

    last_column = apply_cell.column - 1

    found_labels = [(label, cell) for label, cell in ((l, _find_cell(ws, l)) for l in J1979_LABELS) if cell]
    if not found_labels:
        return "NG", f"'{file_name}'に{'/'.join(J1979_LABELS)}のいずれも見つかりません"

    for label, cell in found_labels:
        label_merge = _merged_range_containing(ws, cell)
        label_end_column = label_merge.max_col if label_merge else cell.column
        start_column = label_end_column + 2
        for col in range(start_column, last_column + 1):
            target = ws.cell(row=cell.row, column=col)
            if target.value not in SID_OK_NG_SYMBOLS:
                return "NG", f"'{file_name}'の{target.coordinate}に〇×ではない値があります"

    return "OK", f"'{file_name}'のSID一覧判定表にデータ不備はありません"


def check_sid_list_sheet(local_path):
    """Returns (result, message). OK if SID_LIST_SHEET_NAME exists as a
    sheet in local_path and its J1979 symbol table passes check_sid_list_symbols;
    NG if the file can't be opened, the sheet is missing, or that table has
    a problem. The file name is embedded directly in the message."""
    file_name = os.path.basename(local_path)

    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return "NG", f"'{file_name}'を開けない: {e}"

    if SID_LIST_SHEET_NAME not in wb.sheetnames:
        return "NG", 'GSTCommandファイルにツールが想定している"SID一覧"シートが見つからない。'

    found_message = 'GSTCommandファイルにツールが想定している"SID一覧"シートがある。'
    symbol_result, symbol_message = check_sid_list_symbols(local_path, SID_LIST_SHEET_NAME, file_name)
    if symbol_result != "OK":
        return "NG", symbol_message

    return "OK", f"{found_message}\n{symbol_message}"


def check_gst_command_files(local_paths):
    """Runs check_sid_list_sheet on every given local file path whose name
    matches the CommandSupportGST Excel-file pattern - others are ignored.
    Returns one combined (result, message): OK (with no message) if
    local_paths is empty or none of them match the pattern - this is the
    "skip silently" case, not an error. Otherwise OK only if every matching
    file passed; NG with whichever failing file's message."""
    matching_paths = [path for path in local_paths if is_command_support_gst_file(os.path.basename(path))]

    if not matching_paths:
        return "OK", ""

    per_file_results = [check_sid_list_sheet(path) for path in matching_paths]

    for result, message in per_file_results:
        if result != "OK":
            return "NG", message

    return "OK", " / ".join(message for _, message in per_file_results if message)

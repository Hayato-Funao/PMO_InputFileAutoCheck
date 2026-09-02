"""
Checks against two FIXED reference files - neither is the file a user
submits via SharePoint/PowerApps:

  - server/excel.xlsm (SERVER_REFERENCE_FILE_PATH, a local path on the server)
  - a second file that also lives on SharePoint (not tied to any particular
    submission - the caller downloads it and passes in its local path)

Checks performed:
  - compares server/excel.xlsm's "元表" sheet against the second fixed
    file's "FI FSmatrix" sheet (not 元表 on that side)
  - checks server/excel.xlsm's OK_SHEET_NAME/NG_SHEET_NAME sheets, starting
    right after the "HILSコード" header row, for rows with data in column A
    but nothing in columns B and C

check_definition_file_parts(sharepoint_reference_local_path) is the entry
point other scripts should call - it runs every check above and returns
each as its own (label, result, message) tuple. Since neither file depends
on which item is being processed, callers should call this once per run,
not once per item.

【2026-09-02改修で導入した動的コピー機構（resolve_and_copy_master_file等）は同日中に取消】
①(check_mao)とのマージ作業で「原本を直接改修・参照しない」「バージョン番号をハードコード
しない」という動的解決の仕組みを一度実装したが、ユーザー判断により取消し、元の固定パス定数
（SERVER_REFERENCE_FILE_PATH）へ戻した。本番用ファイルへの変更は②担当者に委ねる。
"""

import openpyxl

# TODO: dummy link - replace with the real local path of the server-side
# reference/definition file that a submitted file's 元表 gets compared against
SERVER_REFERENCE_FILE_PATH = r"C:\Users\RJ067219\OneDrive - Honda\デスクトップ\work\2026_タスク\08_タスク\Check_PythonCode\XPX定義ファイル管理_5.22_FI_ELEC.xlsm"#XPX定義ファイル管理ファイルのパス

MOTOHYOU_SHEET_NAME = "元表"
# The sheet on the second (SharePoint) fixed reference file that server/excel.xlsm's
# 元表 gets compared against - not 元表 on that side, but FI FSmatrix.
SHAREPOINT_COMPARISON_SHEET_NAME = "FI FSmatrix"
OK_SHEET_NAME = "fcl_list_o_FI"
NG_SHEET_NAME = "fcl_list_x_FI"
# The data-integrity check below starts right after this header row, so the
# sheet's own title row (row 1: just the sheet name in column A) and the
# header row itself never get mistaken for actual data.
HILS_CODE_LABEL = "HILSコード"


def get_sheet_rows(local_path, sheet_name):
    """Returns (rows, error_message). rows is the sheet's cell values as a
    list of row tuples if sheet_name exists in local_path; otherwise rows is
    None and error_message says why (file unreadable, or sheet not found)."""
    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return None, f"ファイルを開けない: {e}"

    if sheet_name not in wb.sheetnames:
        return None, f"「{sheet_name}」が見つかりません"

    rows = [row for row in wb[sheet_name].iter_rows(values_only=True)]
    return rows, None


def compare_motohyou_sheets(server_file_path, downloaded_file_path):
    """Returns (result, message). Reads 元表 from the server-side reference
    file (server_file_path) and FI FSmatrix from the second fixed reference
    file that also lives on SharePoint (downloaded_file_path, already
    downloaded by the caller), saves each as row data via get_sheet_rows,
    then compares the two cell-for-cell."""
    server_rows, error = get_sheet_rows(server_file_path, MOTOHYOU_SHEET_NAME)
    if error:
        return "NG", f"サーバー側の定義ファイル - {error}"

    sharepoint_rows, error = get_sheet_rows(downloaded_file_path, SHAREPOINT_COMPARISON_SHEET_NAME)
    if error:
        return "NG", f"SharePoint側の定義ファイル - {error}"

    if server_rows != sharepoint_rows:
        return "NG", f"「{MOTOHYOU_SHEET_NAME}」と「{SHAREPOINT_COMPARISON_SHEET_NAME}」の内容が一致しません"

    return "OK", f"「{MOTOHYOU_SHEET_NAME}」と「{SHAREPOINT_COMPARISON_SHEET_NAME}」の内容が一致しました"


def _find_cell(ws, target_value):
    """Returns the first cell whose value equals target_value, or None if
    not found."""
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == target_value:
                return cell
    return None


def check_o_x_sheet_data(local_path, sheet_name):
    """Returns (result, message) for the given sheet (OK_SHEET_NAME or
    NG_SHEET_NAME) in local_path. Finds the "HILSコード" header cell and
    starts checking from the row right after it - this skips both the
    sheet's title row and the header row itself, so neither is mistaken for
    real data. NG if any row from there has data in column A while columns
    B and C are both empty; OK otherwise."""
    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return "NG", f"ファイルを開けない: {e}"

    if sheet_name not in wb.sheetnames:
        return "NG", f"「{sheet_name}」が見つかりません"

    ws = wb[sheet_name]

    hils_cell = _find_cell(ws, HILS_CODE_LABEL)
    if hils_cell is None:
        return "NG", f"「{sheet_name}」に「{HILS_CODE_LABEL}」が見つかりません"

    start_row = hils_cell.row + 1

    for row in ws.iter_rows(min_row=start_row, min_col=1, max_col=3):
        a_value, b_value, c_value = (cell.value for cell in row)
        if a_value is not None and b_value is None and c_value is None:
            return "NG", f"{sheet_name}にデータ不備があり。"

    return "OK", f"{sheet_name}にデータ不備がない。"


def check_definition_file_parts(sharepoint_reference_local_path):
    """Single entry point: runs every check against SERVER_REFERENCE_FILE_PATH -
    the 元表 comparison against sharepoint_reference_local_path (the second
    fixed reference file, already downloaded by the caller), plus the 〇/×
    sheet data checks - and returns them as three separate (label, result,
    message) tuples, one per topic, so each can be reported on its own line
    instead of being merged into one combined message. Callers that need a
    single overall (result, message) instead (e.g. for a SharePoint
    write-back that can only hold one result) can combine these themselves."""
    return [
        ("元表", *compare_motohyou_sheets(SERVER_REFERENCE_FILE_PATH, sharepoint_reference_local_path)),
        (OK_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, OK_SHEET_NAME)),
        (NG_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, NG_SHEET_NAME)),
    ]

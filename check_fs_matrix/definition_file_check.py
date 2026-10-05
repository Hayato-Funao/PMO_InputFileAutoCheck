"""
Checks around the XPX定義ファイル管理 workbook (元表 and its 〇/× sheets).

【2026-09-04追記】本モジュールは2種類のチェックを持つ。以前は前者だけだった:

  (1) 固定の参照ファイル2つを突き合わせる、アイテム非依存のチェック
      （`check_definition_file_parts`。下記のとおり提出物は見ない）
  (2) **案件ごとに提出されたFSマトリクス**を`元表`のA〜J列＋Q列と突き合わせる、
      アイテム依存のチェック（`load_motohyou_index`＋`reconcile_fs_matrix_parts`。
      詳細はファイル後半の見出しコメント参照）。こちらは②のOK/NG判定と集約ファイルには
      算入せず、②の詳細ファイルにだけ出す。

--- (1) 固定参照ファイル同士のチェック ---

Checks against two FIXED reference files - neither is the file a user
submits via SharePoint/PowerApps:

  - server/excel.xlsm (SERVER_REFERENCE_FILE_PATH, a local or UNC path -
    currently the \\snd89a0 share, see the constant's own note)
  - a second file that also lives on SharePoint (not tied to any particular
    submission - the caller downloads it and passes in its local path)

Checks performed:
  - compares server/excel.xlsm's "元表" sheet against the second fixed
    file's "FI FSmatrix" sheet (not 元表 on that side)
  - checks server/excel.xlsm's OK_SHEET_NAME/NG_SHEET_NAME sheets, starting
    right after the "HILSコード" header row, for rows with data in column A
    but nothing in columns B and C - the NG message lists the offending
    cells (A列のセル番地とHILSコードの値) so the problem can be located in
    the Excel file directly

check_definition_file_parts(sharepoint_reference_local_path) is the entry
point other scripts should call - it runs every check above and returns
each as its own (label, result, message) tuple. Since neither file depends
on which item is being processed, callers should call this once per run,
not once per item.

元表比較が不一致でNGになった場合、続くOK_SHEET_NAME／NG_SHEET_NAMEのチェックは実行せず、
[元表]の1件だけを返して打ち切る（不一致＝サーバー側の定義ファイルが想定の版ではないため、
同じファイルの〇/×シートを検査しても意味が無い）。

【メッセージ書式】詳細ファイル（②の結果ログ）は`check2_report.py`が
[定義ファイル管理Excel]／[FI FSマトリクスファイル]／[GSTコマンド装備表]の3グループに分けて
そのまま並べるため、本モジュールが返すmessageは「グループ見出しの下に箇条書き1行として
置ける文」にする。OK/NGはラベルではなく文末の肯定形／否定形（一致している↔一致していない、
データ不備がない↔データ不備がある）で表し、NGの場合は続けて不備箇所（セル番地等）を書く。
文言を変える場合は`matrix_check`／`gst_command_check`側の文と揃えること。

【2026-09-02改修で導入した動的コピー機構（resolve_and_copy_master_file等）は同日中に取消】
①(check_mao)とのマージ作業で「原本を直接改修・参照しない」「バージョン番号をハードコード
しない」という動的解決の仕組みを一度実装したが、ユーザー判断により取消し、元の固定パス定数
（SERVER_REFERENCE_FILE_PATH）へ戻した。本番用ファイルへの変更は②担当者に委ねる。
"""

import os
import re
import glob

import openpyxl
from openpyxl.utils import get_column_letter

# サーバー側の定義ファイル（XPX定義ファイル管理ファイル）。この`元表`シートが、SharePoint側の
# 参照ファイル（`unified_main.SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`）の`FI FSmatrix`シートと
# 比較される（2026-09-02改修。②担当者の個人環境のローカルコピーから、共有サーバーのUNCパスへ
# 変更した）。
#
# 【注意】UNCパスなので、実行アカウントが`\\snd89a0`の共有へ読み取りアクセスできる必要がある
# （タスクスケジューラで無人実行する場合は、ネットワークドライブではなくUNCパスのまま参照する
# 本設定が前提。ドライブレターは対話セッションにしか割り当たらないため）。アクセスできない場合、
# ②の[元表]／[fcl_list_o_FI]／[fcl_list_x_FI]は「ファイルを開けない」でNGになる。
#
# 【注意】ファイル名にバージョン番号(5.23)が入っているため、定義ファイルの改版時はこのパスも
# 更新が必要（SharePoint側リンクと同様、定期的な見直しが必要）。共有側の運用は「現行版を`FI\`
# 直下に置き、旧版は`FI\old\`へ移す」形なので、改版時は`FI\`直下の最新ファイル名へ更新すること。
# 本ツールではファイル名（5.24、5.25など）を固定で持たず、
# FI直下に存在する現行版を自動検出して使用する。
# 版によって`元表`シートの内容が異なる（5.23=953行／5.22=951行）ため、版を間違えると[元表]の
# OK/NG判定がそのまま変わる点に注意。
# 想定例:
#
# FI
# ├─ XPX定義ファイル管理_5.25_FI_ELEC.xlsm
# └─ old
# ├─ XPX定義ファイル管理_5.24_FI_ELEC.xlsm
# └─ XPX定義ファイル管理_5.23_FI_ELEC.xlsm
#
# 現行版が複数存在する場合は運用異常とみなしエラーとする。
DEFINITION_FILE_DIRECTORY = (
    r"\\snd89a0\proj-hils_pu3\proj-XPX"
    r"\01_Eng\97_定義ファイル管理"
    r"\01_管理\XPX\FI"
)


def resolve_definition_file_path():
    """
    FI直下に存在する XPX定義ファイル管理_*_FI_ELEC.xlsm を検索し、
    現行版ファイルのフルパスを返す。
    
    戻り値:
    現行版定義ファイルのフルパス
    
    例:
    \\snd89a0\...\FI\
    XPX定義ファイル管理_5.25_FI_ELEC.xlsm
    
    異常時:
    ・0件 -> ファイル未配置
    ・2件以上 -> 現行版が複数存在
    """
    
    pattern = os.path.join(
        DEFINITION_FILE_DIRECTORY,
        "XPX定義ファイル管理_*_FI_ELEC.xlsm"
    )
    
    files = glob.glob(pattern)
    
    if not files:
        raise FileNotFoundError(
            "XPX定義ファイル管理ファイルが見つかりません。"
        )
    
    if len(files) > 1:
        raise RuntimeError(
            "XPX定義ファイル管理ファイルが複数存在します。"
            f" 件数={len(files)}"
        )
    
    return files[0]
    
# 現行版のXPX定義ファイル。
#
# 【注意】
# ファイル名に含まれるバージョン番号（5.24等）は固定参照しない。
# 改版時はFI直下のファイルを差し替えるだけでよく、
# 本ソースの修正は不要。
SERVER_REFERENCE_FILE_PATH = resolve_definition_file_path()
#"C:\Users\RJ067219\OneDrive - Honda\デスクトップ\work\2026_タスク\08_タスク\Check_PythonCode\XPX定義ファイル管理_5.22_FI_ELEC.xlsm"
#SERVER_REFERENCE_FILE_PATH = r"\\snd89a0\proj-hils_pu3\proj-XPX\01_Eng\97_定義ファイル管理\01_管理\XPX\FI\XPX定義ファイル管理_5.24_FI_ELEC.xlsm" #本番にリンク
#SERVER_REFERENCE_FILE_PATH = r"C:\Users\RJ067219\OneDrive - Honda\デスクトップ\work\2026_タスク\08_タスク\本番ツール\check_all\XPX定義ファイル管理_5.22_FI_ELEC.xlsm" #テスト用のリンク
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

# 元表比較で不一致になった旨を表す文の語尾（肯定形／否定形）。SP側シート名は
# SHAREPOINT_COMPARISON_SHEET_NAMEだが、結果ログでは提出物の呼び名に合わせて
# 「SP上のFSマトリクスのシート」と表現する。
MOTOHYOU_SP_DESCRIPTION = "SP上のFSマトリクスのシート"

# NGメッセージへ列挙する不備セル／不一致セルの上限。定義ファイルは1000行規模なので、
# 不備が大量にある場合にメッセージが極端に長くなる（＝SharePoint列や詳細ファイルの1行が
# 読めなくなる）のを防ぐ。超過分は「ほか（計N件）」と件数だけ示すので、全件を知りたい場合は
# 列挙されたセルを直して再実行すること。
MAX_REPORTED_DEFECT_ROWS = 10

# 不一致セルの値をメッセージへ載せる際の1セルあたりの最大文字数（長い備考欄などで
# メッセージが膨れるのを防ぐ。超過分は「…」で切る）。
MAX_REPORTED_VALUE_LENGTH = 20


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


def workbook_display_name(local_path):
    """結果ログへ載せるワークブックの呼び名（拡張子なしのファイル名）を返す。
    例: `...\\XPX定義ファイル管理_5.22_FI_ELEC.xlsm` → `XPX定義ファイル管理_5.22_FI_ELEC`。
    版番号がそのまま出るので、どの版を見た結果なのかがログだけで分かる。"""
    return os.path.splitext(os.path.basename(local_path))[0]


def compare_motohyou_sheets(server_file_path, downloaded_file_path):
    """Returns (result, message) - the (label, result, message) 元表 line as
    it appears in the report. See _compare_motohyou for the details; this
    wrapper just drops the internal mismatch flag."""
    result, message, _is_mismatch = _compare_motohyou(server_file_path, downloaded_file_path)
    return result, message


def motohyou_unavailable_message(reason):
    """元表比較そのものを実行できなかった場合の[元表]行を組み立てる（結果ログの文体を
    1箇所に集めるため、呼び出し側で文を組まずにこれを使う）。

    用途: SharePoint側の元表ファイルを取得できず、`check_definition_file_parts`へ渡す
    ローカルパスが存在しない場合（`unified_main.run`）。そのまま実行させると
    「ファイルを開けない」という分かりにくい理由でNGになるため、取得失敗の理由へ差し替える。"""
    workbook_name = workbook_display_name(SERVER_REFERENCE_FILE_PATH)
    return (
        f"「{workbook_name}」の「{MOTOHYOU_SHEET_NAME}」シートと{MOTOHYOU_SP_DESCRIPTION}"
        f"内容が確認できない。{reason}"
    )


def _compare_motohyou(server_file_path, downloaded_file_path):
    """Returns (result, message, is_mismatch). Reads 元表 from the server-side
    reference file (server_file_path) and FI FSmatrix from the second fixed
    reference file that also lives on SharePoint (downloaded_file_path,
    already downloaded by the caller), saves each as row data via
    get_sheet_rows, then compares the two cell-for-cell.

    `is_mismatch`は「両シートを読めたが内容が違った」場合だけTrueにする
    （ファイルを開けない／シートが無い場合はFalse）。`check_definition_file_parts`が
    〇/×シートのチェックを打ち切るかどうかをこのフラグで判定するので、メッセージ文言の
    一致に依存させないこと。"""
    workbook_name = workbook_display_name(server_file_path)
    sentence_head = f"「{workbook_name}」の「{MOTOHYOU_SHEET_NAME}」シートと{MOTOHYOU_SP_DESCRIPTION}内容が"

    server_rows, error = get_sheet_rows(server_file_path, MOTOHYOU_SHEET_NAME)
    if error:
        return "NG", f"{sentence_head}確認できない。サーバー側の定義ファイル - {error}", False

    sharepoint_rows, error = get_sheet_rows(downloaded_file_path, SHAREPOINT_COMPARISON_SHEET_NAME)
    if error:
        return "NG", f"{sentence_head}確認できない。SharePoint側の定義ファイル - {error}", False

    if server_rows != sharepoint_rows:
        return "NG", f"{sentence_head}一致していない。{_describe_differences(server_rows, sharepoint_rows)}", True

    return "OK", f"{sentence_head}一致している。", False


def _shorten(value):
    """セル値を結果ログ用に文字列化する（未入力は`(空欄)`、長い値はMAX_REPORTED_VALUE_LENGTHで切る）。"""
    if value is None:
        return "(空欄)"
    text = str(value)
    if len(text) > MAX_REPORTED_VALUE_LENGTH:
        return text[:MAX_REPORTED_VALUE_LENGTH] + "…"
    return text


def _describe_differences(server_rows, sharepoint_rows):
    """元表とSP側シートの差分を、Excel上の位置が分かる1行の文へ整形する。
    例: `不一致セル: B2（元表=値2／FI FSmatrix=値2_不一致）`

    行数自体が違う場合はセル番地より先に行数差を書く（版違いはまず行数に出るため。
    5.23=953行／5.22=951行）。セル番地の列挙はMAX_REPORTED_DEFECT_ROWS件まで。"""
    parts = []
    if len(server_rows) != len(sharepoint_rows):
        parts.append(
            f"行数が異なる（{MOTOHYOU_SHEET_NAME}={len(server_rows)}行／"
            f"{SHAREPOINT_COMPARISON_SHEET_NAME}={len(sharepoint_rows)}行）"
        )

    differences = []
    for row_index in range(min(len(server_rows), len(sharepoint_rows))):
        server_row = server_rows[row_index]
        sharepoint_row = sharepoint_rows[row_index]
        for col_index in range(max(len(server_row), len(sharepoint_row))):
            server_value = server_row[col_index] if col_index < len(server_row) else None
            sharepoint_value = sharepoint_row[col_index] if col_index < len(sharepoint_row) else None
            if server_value != sharepoint_value:
                cell = f"{get_column_letter(col_index + 1)}{row_index + 1}"
                differences.append(
                    f"{cell}（{MOTOHYOU_SHEET_NAME}={_shorten(server_value)}／"
                    f"{SHAREPOINT_COMPARISON_SHEET_NAME}={_shorten(sharepoint_value)}）"
                )
        if len(differences) > MAX_REPORTED_DEFECT_ROWS:
            # 上限を超えた分は件数も出さない（全行を突き合わせずに打ち切るため、
            # 総件数が分からない）。列挙されたセルを直して再実行する運用にする。
            break

    if differences:
        listed = ", ".join(differences[:MAX_REPORTED_DEFECT_ROWS])
        if len(differences) > MAX_REPORTED_DEFECT_ROWS:
            listed += " ほか"
        parts.append(f"不一致セル: {listed}")

    if not parts:
        return "差分セルを特定できない。"
    return "。".join(parts) + "。"


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
    B and C are both empty; OK otherwise.

    NG時のメッセージには、Excelのどこが不備なのかをそのまま開いて辿れるよう、該当セル
    （A列のセル番地＋そのHILSコードの値）を列挙する。最初の1件で打ち切らず全行を走査
    するので、1回の実行で不備行をまとめて把握できる（列挙はMAX_REPORTED_DEFECT_ROWS件まで、
    超過分は件数のみ）。"""
    workbook_name = workbook_display_name(local_path)
    sentence_head = f"「{workbook_name}」の「{sheet_name}」に"

    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return "NG", f"{sentence_head}データ不備がないか確認できない。ファイルを開けない: {e}"

    if sheet_name not in wb.sheetnames:
        return "NG", f"「{workbook_name}」に「{sheet_name}」シートがない。"

    ws = wb[sheet_name]

    hils_cell = _find_cell(ws, HILS_CODE_LABEL)
    if hils_cell is None:
        return "NG", f"{sentence_head}データ不備がないか確認できない。「{HILS_CODE_LABEL}」が見つからない。"

    start_row = hils_cell.row + 1

    defect_rows = []  # [(行番号, A列の値)]
    for row in ws.iter_rows(min_row=start_row, min_col=1, max_col=3):
        a_value, b_value, c_value = (cell.value for cell in row)
        if a_value is not None and b_value is None and c_value is None:
            defect_rows.append((row[0].row, a_value))

    if defect_rows:
        return "NG", f"{sentence_head}データ不備がある。{_describe_defect_rows(defect_rows)}"

    return "OK", f"{sentence_head}データ不備がない。"


def _describe_defect_rows(defect_rows):
    """不備行のリスト（[(行番号, A列の値)]）を、Excel上の位置が分かる1行の文へ整形する。
    例: `A列に値があるのにB列・C列が両方空欄: A123（HILSコード=H001）, A130（HILSコード=H008）`

    詳細ファイル／SharePoint列は1件1行で書き出されるため、改行を含めず1行に収める。"""
    listed = ", ".join(
        f"A{row_number}（{HILS_CODE_LABEL}={a_value}）"
        for row_number, a_value in defect_rows[:MAX_REPORTED_DEFECT_ROWS]
    )
    if len(defect_rows) > MAX_REPORTED_DEFECT_ROWS:
        listed += f" ほか（計{len(defect_rows)}件）"
    return f"A列に値があるのにB列・C列が両方空欄: {listed}。"


# =====================================================================================
# 提出FSマトリクス ↔ 元表 突合（2026-09-04新規）
#
# 上の`_compare_motohyou`は「固定の参照ファイル2つ（サーバー側定義ファイルとSP上の元表
# ファイル）が同じ版か」をセル単位で見るチェックで、案件ごとの提出物は見ていない。ここから
# 下は**案件ごとに提出されたFSマトリクス**を`元表`と突合する、アイテム依存のチェックである。
#
# 【何を突合するか】`XPX定義ファイル管理`のVBA（`LoadFSMatrix`）がFSマトリクスから実際に
# 読み取る11フィールドだけを対象にする。`元表`側の列で言うとA〜J列＋Q列:
#
#     +0  A  診断識別コード          HILSコードの構成要素
#     +1  B  診断識別コード/枝番      HILSコードの構成要素
#     +2  C  モニタ項目
#     +3  D  担当区                （ENG以外の行はVBAが読み飛ばすため突合対象から除く）
#     +4  E  MIL                  HILSコードの構成要素（点灯→1／それ以外→0）
#     +5  F  故障確定D/C            HILSコードの構成要素
#     +6  G  SAE J2012 CODE       HILSコードの構成要素（FTBと連結される）
#     +7  H  FTB                  HILSコードの構成要素（u／z枝番のみ）
#     +8  I  システム識別番号
#     +9  J  SCSコード              HILSコードの構成要素
#     +16 Q  OBDシステムコンセプト書DWG
#
# K〜P列（GST出力項目・Readiness Group・RGID・TestResult・IUMPR）とR列以降はVBAが読み
# 取らないため突合しない（差異があってもXPXの実行結果には影響しない）。
#
# 【なぜ突合が必要か】`XPX定義ファイル管理`は上記フィールドから
# `{診断識別コード}-{枝番}_{SAE+FTB}_{MIL}_{D/C}_{SCS}`というHILSコードを組み立て、これを
# `fcl_list_o_FI`／`fcl_list_x_FI`のキーとして`Fail_setting_FI`／`DP_Setting_FI`をVLOOKUPで
# 引く。フィールドが1つでも食い違うとHILSコードが変わり、引き当てられず`#N/A`になる。
# `Convert_CSV`は`#N/A`のままCSVへ出力し、XPXはそのCSVを無検査で読むため、食い違いは
# XPXの実行時まで気付かれない。
#
# 【2つのチェックに分かれている】判定への算入有無が違うので混同しないこと:
#
#   (a) `check_fs_matrix_codes_in_motohyou` — **②のOK/NG判定に算入する**
#       FSマトリクスのコードが`元表`に在るか（片方向）。無ければNG。
#       `元表`に無いコードは`XPX定義ファイル管理`が「新規コード」として`fcl_list`へ追加し、
#       設定の無いままCSVへ出るため、XPXが引き当てられない状態になる＝実害が確定している。
#
#   (b) `check_fs_matrix_data_mismatches` — 戻り値が2つに分かれ、算入有無も分かれる
#       両方に在るコードの値（A〜J列＋Q列）の食い違いを、重要度で分けて返す:
#         * `hils_parts`  … HILSコードの構成要素（SAE／FTB／MIL／D/C／SCS）の不一致。
#                            **②のOK/NG判定に算入する**。HILSコードが変わり実害が確定するため。
#         * `other_parts` … それ以外の列だけの不一致。算入しない（実データではOBD DWGの
#                            改訂差だけで100件超になるのが常態で、算入すると全案件がNGになる）。
#
# 集約ファイル（`インプットファイルチェック結果.txt`）の②行末へは、(a)と(b)の`hils_parts`が
# NGのときだけ短文を出す（`code_check_aggregate_summary`／`data_check_aggregate_summary`）。
# 欠落コード・不一致の一覧は②の詳細ファイル側にだけ置く。グループ自体は
# `check2_report.GROUPS_EXCLUDED_FROM_AGGREGATE`に入れたままなので、`other_parts`や
# 一覧が集約ファイルへ漏れることはない。
# =====================================================================================

# `元表`／`FI FSmatrix`のヘッダー行を特定するためのアンカー文字列。VBA（`LoadFSMatrix`）が
# 「改行を除去してTrimした結果がこの文字列に完全一致するセル」を先頭20行×全使用列から探すのと
# 同じ条件で探す。`元表`はA6・FSマトリクスはA9等、絶対行はファイルごとに違うため、行番号を
# 決め打ちせず必ず探索する。
ANCHOR_HEADER = "診断識別コード"
ANCHOR_SEARCH_ROWS = 20

# 突合するフィールド: (キー, アンカーからのオフセット, 表示名, HILSコードの構成要素か)
RECONCILE_FIELDS = (
    ("diag", 0, "診断識別コード", True),
    ("branch", 1, "枝番", True),
    ("monitor", 2, "モニタ項目", False),
    ("dept", 3, "担当区", False),
    ("mil", 4, "MIL", True),
    ("dc", 5, "故障確定D/C", True),
    ("sae", 6, "SAEコード", True),
    ("ftb", 7, "FTB", True),
    ("sysid", 8, "システム識別番号", False),
    ("scs", 9, "SCSコード", True),
    ("obd", 16, "OBDシステムコンセプト書DWG", False),
)

# 突合対象の行を絞る担当区（VBAの`If Trim(CStr(dept)) <> "ENG" Then GoTo ContinueLoad`と同条件）。
# これ以外の担当区の行は`XPX定義ファイル管理`へ取り込まれないため、差異があっても無害。
RECONCILE_TARGET_DEPT = "ENG"

# 突合から除外する診断識別コードの接頭辞。VBAは`isPlaceholder`（診断識別コードが"("始まり）の
# 行を`fcl_list`へ反映しないため、`(7n)`等のプレースホルダーは差異として報告しない。
PLACEHOLDER_PREFIX = "("

# 提出FSマトリクス側の突合対象シート名（`matrix_check`はシート名の揺れを9種類許容するが、
# `XPX定義ファイル管理`のVBAはこの名前を決め打ちで探すため、突合もこの名前だけを見る）。
FS_MATRIX_SHEET_NAME = "FI FSmatrix"

# コード突合の結果を表す文（2026-09-04）。②の詳細ファイルではこの文の後ろへ欠落コードの
# 一覧を続けるが、集約ファイル（`インプットファイルチェック結果.txt`）へは
# `MOTOHYOU_CODE_NG_SUMMARY`だけを出し、一覧は載せない（一覧は詳細ファイルで確認する運用）。
MOTOHYOU_CODE_OK_MESSAGE = "FSマトリクスコードが元表シートにある。"
MOTOHYOU_CODE_NG_MESSAGE = "FSマトリクスコードが元表シートにない。"
MOTOHYOU_CODE_NG_SUMMARY = "FSマトリクスコードが元表シートにない"
# 突合そのものを実行できなかった場合の文（元表／FSマトリクスを開けない、シートが無い等）。
# 「コードが無い」とは意味が違うので、集約ファイルへもこちらの短文を出し分ける。
_CODE_UNCHECKABLE_MESSAGE = "FSマトリクスコードが元表シートにあるか確認できない。"
_CODE_UNCHECKABLE_SUMMARY = "FSマトリクスコードが元表シートにあるか確認できない"


# 値の突合（`check_fs_matrix_data_mismatches`）の結果を表す文（2026-09-04）。コード突合と同じく、
# 詳細ファイルではこの文の後ろへ不一致の一覧を続け、集約ファイルへは`..._SUMMARY`だけを出す。
MOTOHYOU_DATA_OK_MESSAGE = "FSマトリクスと元表シートのデータが一致している。"
MOTOHYOU_DATA_NG_MESSAGE = "FSマトリクスと元表シートのデータが一致しない。"
MOTOHYOU_DATA_NG_SUMMARY = "FSマトリクスと元表シートのデータが一致しない"


def data_check_aggregate_summary(hils_parts):
    """`check_fs_matrix_data_mismatches`が返した`hils_parts`から、集約ファイルの②行末へ出す
    短文を返す（2026-09-04新規）。NGが1件も無ければNone。

    不一致の一覧は②の詳細ファイル側にだけ置き、集約ファイルへは「一致しない」ことだけを示す
    （`code_check_aggregate_summary`と同じ方針）。
    """
    if any(result != "OK" for result, _message in hils_parts):
        return MOTOHYOU_DATA_NG_SUMMARY
    return None


def code_check_aggregate_summary(code_parts):
    """`check_fs_matrix_codes_in_motohyou`の結果から、集約ファイル
    （`インプットファイルチェック結果.txt`）の②行末へ出す短文を返す（2026-09-04新規）。

    欠落コードの一覧は②の詳細ファイル側にだけ置き、集約ファイルへは「無い」または
    「確認できない」ことだけを示す。NGが1件も無ければNoneを返す。

    「コードが無い」と「突合そのものを実行できなかった」は原因が別物なので短文を出し分ける
    （後者でも②はNGだが、提出者に伝える内容は「シートが無い／開けない」であり、
    「コードが元表に無い」と書くと誤解を招く）。
    """
    for result, message in code_parts:
        if result == "OK":
            continue
        if _CODE_UNCHECKABLE_MESSAGE in message:
            return _CODE_UNCHECKABLE_SUMMARY
        return MOTOHYOU_CODE_NG_SUMMARY
    return None

_WHITESPACE = re.compile(r"\s+")


def _normalize_cell(value):
    """比較用にセル値を正規化する（全角空白→半角、改行・連続空白を1つの半角空白へ、前後を除去）。

    FSマトリクスと元表は同じ内容でも改行位置・全角空白の有無が揺れるため、書式差を差異として
    報告しないようにする。`None`は空文字にする。
    """
    if value is None:
        return ""
    return _WHITESPACE.sub(" ", str(value).replace("　", " ")).strip()


def _normalize_field(raw, key):
    """フィールド1つを比較用に正規化する。SAEコードとSCSコードだけはVBA（`LoadFSMatrix`）と
    同じ追加正規化を行う（SAE: "/"→"-"と空白除去／SCS: []{}と空白の除去）。HILSコードへ載る
    のは正規化後の値なので、正規化前の書式差でNGにしないため。"""
    text = _normalize_cell(raw)
    if key == "sae":
        return text.replace("/", "-").replace(" ", "")
    if key == "scs":
        for char in "[]{} ":
            text = text.replace(char, "")
        return text
    return text


def _find_anchor_cell(ws):
    """アンカー（`診断識別コード`セル）の(行, 列)を返す。見つからなければ(None, None)。"""
    max_col = ws.max_column or 1
    max_row = min(ANCHOR_SEARCH_ROWS, ws.max_row or 1)
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            text = ws.cell(row, col).value
            if text is None:
                continue
            if str(text).replace("\n", "").replace("\r", "").strip() == ANCHOR_HEADER:
                return row, col
    return None, None


def _load_reconcile_rows(ws):
    """シートから突合対象行を読み出す。

    戻り値:
        (rows, anchor_col, error) — rowsは`{"row": 行番号, "norm": {キー: 正規化値}}`のリスト。
        アンカーが見つからない場合は(None, None, 理由)。

    VBAと同じく、診断識別コードが空欄の行（サブヘッダー・空行）、担当区が`ENG`以外の行、
    プレースホルダー（"("始まり）の行は読み飛ばす。
    """
    anchor_row, anchor_col = _find_anchor_cell(ws)
    if anchor_row is None:
        return None, None, f"「{ANCHOR_HEADER}」列が見つからない。"

    rows = []
    for row_number in range(anchor_row + 1, (ws.max_row or anchor_row) + 1):
        norm = {
            key: _normalize_field(ws.cell(row_number, anchor_col + offset).value, key)
            for key, offset, _label, _in_hils in RECONCILE_FIELDS
        }
        if not norm["diag"]:
            continue
        if norm["diag"].startswith(PLACEHOLDER_PREFIX):
            continue
        if norm["dept"] != RECONCILE_TARGET_DEPT:
            continue
        rows.append({"row": row_number, "norm": norm})

    return rows, anchor_col, None


def _reconcile_key(row):
    """突合キー（診断識別コード＋枝番）。"""
    return (row["norm"]["diag"], row["norm"]["branch"])


def _reconcile_signature(row, hils_only=False):
    """同一キーの行同士を突き合わせるための正規化値タプル。

    `hils_only=True`ならHILSコードの構成要素だけを並べる。1キーに複数行あるケースで
    「差異はあるがHILSコードには影響しない」を判別するために使う。
    """
    return tuple(
        row["norm"][key] for key, _o, _l, in_hils in RECONCILE_FIELDS if in_hils or not hils_only
    )


def _group_by_reconcile_key(rows):
    """キー→行リストのdictを返す。診断識別コード＋枝番が同じでモニタ項目だけ違う行が実在する
    （例: `71`の失火モニタが`Criteria A+B`と`Criteria A only`の2行）ため、1キー複数行を許す。"""
    grouped = {}
    for row in rows:
        grouped.setdefault(_reconcile_key(row), []).append(row)
    return grouped


def load_motohyou_index(server_reference_path=None):
    """`XPX定義ファイル管理`の`元表`シートを読み、突合用のインデックスを組み立てる。

    `元表`はアイテム（案件）に依存しないため、`unified_main.run`で**実行ごとに1回だけ**呼び、
    戻り値を全アイテムで共有する（3MB超のxlsmを案件ごとに開くと遅いため）。

    引数:
        server_reference_path: 省略時は`SERVER_REFERENCE_FILE_PATH`

    戻り値:
        (index, error) — indexは`{"grouped": {キー: [行, ...]}, "anchor_col": 列番号,
        "name": 表示名, "row_count": 行数}`。読めない場合は(None, 理由)。呼び出し側は理由を
        そのままNGメッセージに使う。
    """
    path = server_reference_path or SERVER_REFERENCE_FILE_PATH

    try:
        wb = openpyxl.load_workbook(path, data_only=True)
    except Exception as e:
        return None, f"ファイルを開けない: {e}"

    if MOTOHYOU_SHEET_NAME not in wb.sheetnames:
        return None, f"「{MOTOHYOU_SHEET_NAME}」シートがない。"

    rows, anchor_col, error = _load_reconcile_rows(wb[MOTOHYOU_SHEET_NAME])
    if error is not None:
        return None, f"「{MOTOHYOU_SHEET_NAME}」シートの{error}"

    return (
        {
            "grouped": _group_by_reconcile_key(rows),
            "anchor_col": anchor_col,
            "name": workbook_display_name(path),
            "row_count": len(rows),
        },
        None,
    )


def _clip_value(text):
    """メッセージへ載せる値を`MAX_REPORTED_VALUE_LENGTH`で切る（空欄は`(空欄)`と書く）。"""
    if not text:
        return "(空欄)"
    if len(text) > MAX_REPORTED_VALUE_LENGTH:
        return text[:MAX_REPORTED_VALUE_LENGTH] + "…"
    return text


def _describe_mismatch(key, matrix_row, motohyou_row, matrix_anchor_col, motohyou_anchor_col):
    """データ不一致1件を1行の文へ整形する（食い違ったフィールドだけ列挙し、両側のセル番地を書く）。

    戻り値:
        (message, affects_hils) — messageは1行の文（差異が無ければNone）、
        `affects_hils`は食い違ったフィールドにHILSコードの構成要素が含まれるか。
        呼び出し側はこのフラグで報告の優先順位を分ける（`reconcile_fs_matrix_parts`参照）。
        HILSコードの構成要素のラベルには`*`を付けて、XPXの実行に影響する差異だと分かるようにする。
    """
    diag, branch = key
    diffs = []
    affects_hils = False
    for field_key, offset, label, in_hils in RECONCILE_FIELDS:
        if field_key in ("diag", "branch", "dept"):
            continue  # キー自体と、突合対象を絞るだけの担当区は比較しない
        matrix_value = matrix_row["norm"][field_key]
        motohyou_value = motohyou_row["norm"][field_key]
        if matrix_value == motohyou_value:
            continue
        if in_hils:
            affects_hils = True
        matrix_cell = f"{get_column_letter(matrix_anchor_col + offset)}{matrix_row['row']}"
        motohyou_cell = f"{get_column_letter(motohyou_anchor_col + offset)}{motohyou_row['row']}"
        diffs.append(
            f"{label}{'*' if in_hils else ''}: FSマトリクス{matrix_cell}={_clip_value(matrix_value)}"
            f"／{MOTOHYOU_SHEET_NAME}{motohyou_cell}={_clip_value(motohyou_value)}"
        )
    if not diffs:
        return None, False
    return f"{diag}/{branch}（" + "、".join(diffs) + "）", affects_hils


def _join_capped(items, limit=None):
    """差異の一覧を`limit`件（既定は`MAX_REPORTED_DEFECT_ROWS`）までで連結し、超過分は件数だけ示す。"""
    cap = MAX_REPORTED_DEFECT_ROWS if limit is None else limit
    listed = " / ".join(items[:cap])
    if len(items) > cap:
        listed += f" ほか（計{len(items)}件）"
    return listed


def _open_fs_matrix_rows(local_path):
    """提出FSマトリクスの`FI FSmatrix`シートから突合対象行を読み出す。

    戻り値:
        (grouped, anchor_col, error) — groupedはキー→行リスト。読めない場合は(None, None, 理由)。
    """
    try:
        wb = openpyxl.load_workbook(local_path, data_only=True)
    except Exception as e:
        return None, None, f"ファイルを開けない: {e}"

    if FS_MATRIX_SHEET_NAME not in wb.sheetnames:
        return None, None, f"「{FS_MATRIX_SHEET_NAME}」シートがない。"

    rows, anchor_col, error = _load_reconcile_rows(wb[FS_MATRIX_SHEET_NAME])
    if error is not None:
        return None, None, error

    return _group_by_reconcile_key(rows), anchor_col, None


def check_fs_matrix_codes_in_motohyou(motohyou_index, motohyou_error, fs_matrix_local_paths):
    """【②の判定に算入する新条件（2026-09-04追加）】提出FSマトリクスの`FI FSmatrix`シートに
    ある診断識別コードが、すべて`XPX定義ファイル管理`の`元表`に存在するかを確認する。

    存在しないコードが1件でもあれば**NG**とする。`XPX定義ファイル管理`は`元表`に無いコードを
    「新規コード」として`fcl_list`へ追加してしまうため、`Fail_setting_FI`／`DP_Setting_FI`に
    設定が無いまま定義ファイルCSVへ出力され、XPXが引き当てられない状態になる。

    【片方向だけ見る理由】`元表`にあってFSマトリクスに無いコードはNGにしない。`元表`は全機種の
    マスタであり、機種別のFSマトリクスがその部分集合になるのは正常だから（実データでは
    `元表`側のみのコードが常時20件前後ある）。見るのは「FSマトリクス → `元表`」の方向のみ。

    引数:
        motohyou_index: `load_motohyou_index`が返したインデックス（読めなかった場合はNone）
        motohyou_error: `load_motohyou_index`が返した理由（読めた場合はNone）
        fs_matrix_local_paths: 提出されたFSマトリクスのローカルパス（0件以上）

    戻り値:
        [(result, message), ...]。`check2_report`のグループ見出しの下へそのまま並べられる文。
        FSマトリクスが0件の場合は空リストを返す（サイレントスキップ。
        `gst_command_check.check_gst_command_files_parts`と同じ扱い）。
    """
    if not fs_matrix_local_paths:
        return []

    if motohyou_index is None:
        return [("NG", f"{_CODE_UNCHECKABLE_MESSAGE}{motohyou_error}")]

    motohyou_name = motohyou_index["name"]
    motohyou_grouped = motohyou_index["grouped"]

    parts = []
    # ファイルが2つ以上ある場合だけ、どのファイルの行なのか分かるようファイル名を付ける
    # （`matrix_check.check_matrix_sheet_parts_for_all_files`と同じ方針）。
    show_file_name = len(fs_matrix_local_paths) > 1

    for local_path in fs_matrix_local_paths:
        prefix = f"「{os.path.basename(local_path)}」: " if show_file_name else ""

        matrix_grouped, matrix_anchor_col, error = _open_fs_matrix_rows(local_path)
        if error is not None:
            parts.append(("NG", f"{prefix}{_CODE_UNCHECKABLE_MESSAGE}{error}"))
            continue

        missing = [
            f"{diag}/{branch}"
            f"（FSマトリクス{get_column_letter(matrix_anchor_col)}{matrix_grouped[(diag, branch)][0]['row']}）"
            for diag, branch in sorted(set(matrix_grouped) - set(motohyou_grouped))
        ]

        if missing:
            # 詳細ファイル用のメッセージ。先頭は`MOTOHYOU_CODE_NG_MESSAGE`のみで、その後ろに
            # 欠落コードの一覧を続ける。集約ファイルへは`unified_main._run_check2`が
            # `MOTOHYOU_CODE_NG_SUMMARY`（一覧なしの短文）だけを出す。
            parts.append(
                (
                    "NG",
                    f"{prefix}{MOTOHYOU_CODE_NG_MESSAGE}"
                    f"「{motohyou_name}」に無いコード{len(missing)}件: {_join_capped(missing)}。",
                )
            )
        else:
            parts.append(("OK", f"{prefix}{MOTOHYOU_CODE_OK_MESSAGE}"))

    return parts


def check_fs_matrix_data_mismatches(motohyou_index, motohyou_error, fs_matrix_local_paths):
    """`FI FSmatrix`と`元表`の**両方にある**診断識別コードについて、A〜J列＋Q列の値が
    食い違っていないかを見る（2026-09-04追加）。

    `check_fs_matrix_codes_in_motohyou`が拾えない不備を補うためにある。コード自体は両方に
    存在するのに値だけ違うケース（例: `u`枝番のSAEコード・FTBが空欄で、`元表`側には
    `P0150`／`7C`が入っている）は、HILSコードが変わるためXPXが引き当てられなくなるが、
    コードの有無だけを見るチェックでは検出できない。

    【重要度で2つに分け、判定に算入するのはHILSコードに影響する分だけ】(2026-09-04)
      * HILSコードの構成要素（SAE／FTB／MIL／D/C／SCS）の不一致 → **NG。②の判定に算入し、
        集約ファイルにも短文を出す**。HILSコードが変わり実害が確定するため。
      * それ以外の列（モニタ項目・システム識別番号・OBDシステムコンセプト書DWG）だけの
        不一致 → OK扱いの参考情報。判定にも集約ファイルにも出さない。実データでは
        DWGの改訂差だけで100件超になるのが常態で、算入すると全案件がNGになるため。
    列挙も分けている。まとめて`MAX_REPORTED_DEFECT_ROWS`件で切ると、DWGの改訂差に押し出されて
    SAE／FTB等の重要な不一致が列挙から消えるため（2026-09-04の実データ確認で判明）。

    戻り値:
        (hils_parts, other_parts) — どちらも[(result, message), ...]。
        `hils_parts`は②の判定へ算入する分（不一致が無ければOK1件）、`other_parts`は
        参考情報（HILSコード外の不一致が無ければ空）。FSマトリクスが0件、または`元表`を
        読めない場合はどちらも空リスト（読めない旨は`check_fs_matrix_codes_in_motohyou`側が
        既にNGで報告するため、重複させない）。
    """
    if not fs_matrix_local_paths or motohyou_index is None:
        return [], []

    motohyou_grouped = motohyou_index["grouped"]
    motohyou_anchor_col = motohyou_index["anchor_col"]

    hils_parts = []
    other_parts = []
    show_file_name = len(fs_matrix_local_paths) > 1

    for local_path in fs_matrix_local_paths:
        prefix = f"「{os.path.basename(local_path)}」: " if show_file_name else ""

        matrix_grouped, matrix_anchor_col, error = _open_fs_matrix_rows(local_path)
        if error is not None:
            continue  # 読めない旨はコード突合側が報告済み

        hils_mismatches = []
        other_mismatches = []
        for key in sorted(set(matrix_grouped) & set(motohyou_grouped)):
            matrix_side = matrix_grouped[key]
            motohyou_side = motohyou_grouped[key]

            # 同一キーが複数行あるケースは、全フィールドの正規化値タプルの多重集合として比較する。
            # 両側で同じ組み合わせが揃っていれば差異なしとみなす（行の並び順には依存しない）。
            if sorted(_reconcile_signature(r) for r in matrix_side) == sorted(
                _reconcile_signature(r) for r in motohyou_side
            ):
                continue

            if len(matrix_side) == 1 and len(motohyou_side) == 1:
                described, affects_hils = _describe_mismatch(
                    key, matrix_side[0], motohyou_side[0], matrix_anchor_col, motohyou_anchor_col
                )
                if described:
                    (hils_mismatches if affects_hils else other_mismatches).append(described)
            else:
                # 複数行同士（`71`の失火モニタのように診断識別コード＋枝番が同じ行が複数ある
                # ケース）は、どの行とどの行が対応するかを決められないため個別の差異は書けない。
                # ただしHILSコードの構成要素だけの多重集合が一致していれば、差異はHILSコード外
                # （モニタ項目・OBD DWG等）に限られると判定できるので、影響なし側へ分類する。
                diag, branch = key
                described = (
                    f"{diag}/{branch}（同一コードの行の内容が一致しない。"
                    f"FSマトリクス{len(matrix_side)}行／{MOTOHYOU_SHEET_NAME}{len(motohyou_side)}行）"
                )
                if sorted(_reconcile_signature(r, hils_only=True) for r in matrix_side) == sorted(
                    _reconcile_signature(r, hils_only=True) for r in motohyou_side
                ):
                    other_mismatches.append(described)
                else:
                    hils_mismatches.append(described)

        # HILSコードに影響する不一致＝XPXが引き当てられなくなる実害があるので**NG**にする
        # （②の判定に算入し、集約ファイルにも短文を出す）。
        if hils_mismatches:
            hils_parts.append(
                (
                    "NG",
                    f"{prefix}{MOTOHYOU_DATA_NG_MESSAGE}"
                    f"HILSコードに影響する不一致{len(hils_mismatches)}件: "
                    f"{_join_capped(hils_mismatches)}。"
                    f"（*はHILSコードの構成要素＝XPXの実行に影響する差異）",
                )
            )
        else:
            hils_parts.append(("OK", f"{prefix}{MOTOHYOU_DATA_OK_MESSAGE}"))

        # HILSコード外のフィールドだけの差異（OBDシステムコンセプト書DWGの改訂差など）。
        # XPXの実行結果には影響しないため判定に算入せず、件数を主に示して列挙は先頭数件に絞る
        # （改訂差で100件超になるのが常態のため）。
        if other_mismatches:
            other_parts.append(
                (
                    "OK",
                    f"{prefix}HILSコード外のみの不一致が{len(other_mismatches)}件ある"
                    f"（実行への影響なし・判定対象外）: {_join_capped(other_mismatches, limit=3)}。",
                )
            )

    return hils_parts, other_parts


def check_definition_file_parts(sharepoint_reference_local_path=None):
    """Single entry point: runs every check against SERVER_REFERENCE_FILE_PATH and
    returns them as separate (label, result, message) tuples, one per topic, so
    each can be reported on its own line instead of being merged into one
    combined message. Callers that need a single overall (result, message)
    instead (e.g. for a SharePoint write-back that can only hold one result)
    can combine these themselves.

    【2026-09-04 元表比較を無効化】`元表`シートを
    `unified_main.SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`（SP上の元表ファイル）の
    `FI FSmatrix`シートとセル単位で突き合わせる[元表]の条件は、ユーザー判断により
    コメントアウトした。判定にも算入せず、結果ログにも出さない（返るタプルから[元表]が
    消え、〇/×シート2件のみになる）。

    再開する場合は下のコメントアウト部分と`unified_main.run`側の
    `_resolve_definition_reference`呼び出し（同じ日付のコメント）を両方戻すこと。
    引数`sharepoint_reference_local_path`は無効化中は使わないため省略可にしてある。

    無効化に伴い、元表が不一致のときに〇/×シートのチェックを打ち切っていた挙動も止まる
    （打ち切りの根拠が元表比較の結果だったため）。〇/×シートは常に評価される。

    `fcl_list_o_FI`／`fcl_list_x_FI`の〇/×データ不備チェックは従来どおり変更していない。"""
    # ---- [元表] 比較（2026-09-04 無効化。再開時はこのブロックのコメントを外す） ----------
    # result, message, is_mismatch = _compare_motohyou(
    #     SERVER_REFERENCE_FILE_PATH, sharepoint_reference_local_path
    # )
    # motohyou_part = (MOTOHYOU_SHEET_NAME, result, message)
    #
    # # 元表が不一致＝サーバー側の定義ファイルが想定の版ではないので、同じファイルの
    # # 〇/×シートを検査しても意味のある結果にならない。ここで打ち切る。
    # if is_mismatch:
    #     return [motohyou_part]
    #
    # return [
    #     motohyou_part,
    #     (OK_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, OK_SHEET_NAME)),
    #     (NG_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, NG_SHEET_NAME)),
    # ]
    # ------------------------------------------------------------------------------------

    return [
        (OK_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, OK_SHEET_NAME)),
        (NG_SHEET_NAME, *check_o_x_sheet_data(SERVER_REFERENCE_FILE_PATH, NG_SHEET_NAME)),
    ]

"""
matrix_check.py／gst_command_check.py／definition_file_check.py（②check_fs_matrixの
判定ロジック）を検証するための最小フィクスチャ（xlsx）を生成するモジュール
（2026-09-02新規。①(check_mao)の`テスト\\generate_fixtures.py`と同じ考え方で、
openpyxlで最小限のワークブックを組み立てる）。
"""

from openpyxl import Workbook

# matrix_check.py / gst_command_check.py が○×として許容する記号（両モジュール共通の定義を
# ここでも1つだけ使う。テストで「許容される記号」であることが分かればよいため、実データの
# 全表記ゆれを網羅する必要はない）。
OK_SYMBOL = "○"
NG_SYMBOL_INVALID = "△"  # ○×どちらでもない、想定外の記号


def generate_matrix_fixture(path, sheet_name="FSマトリクス", table_symbol=OK_SYMBOL):
    """
    matrix_check.check_matrix_sheetが想定するFSマトリクスシートの最小構成を生成する。

    1行目に「部番」ラベルとその右に番号始まりのセルを1列、2行目に「診断識別コード」ラベル、
    3〜4行目がデータ（診断識別コード列に値があり、部番列に`table_symbol`が入る）。
    `table_symbol`にOK_SYMBOL以外（例:NG_SYMBOL_INVALID）を渡すと、
    check_fs_matrix_symbolsがNGになるデータを生成できる。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = sheet_name

    sheet.cell(row=1, column=1, value="部番")
    sheet.cell(row=1, column=2, value="12345")  # 数字始まり列（判定対象列）
    sheet.cell(row=2, column=1, value="診断識別コード")
    sheet.cell(row=3, column=1, value="D001")
    sheet.cell(row=3, column=2, value=table_symbol)
    sheet.cell(row=4, column=1, value="D002")
    sheet.cell(row=4, column=2, value=table_symbol)

    workbook.save(path)


def generate_no_matrix_sheet_fixture(path):
    """matrix_check.MATRIX_SHEET_NAME_VARIANTSのいずれにも一致しないシート名のみを持つ
    ワークブックを生成する（check_matrix_sheetがNGになることの確認用）。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "無関係なシート"
    sheet.cell(row=1, column=1, value="dummy")
    workbook.save(path)


def generate_gst_command_fixture(path, table_symbol=OK_SYMBOL):
    """
    gst_command_check.check_sid_list_sheetが想定する「SID一覧」シートの最小構成を生成する。

    1行目に"SAE J1979"ラベルとその2列右から"J1979を適用する場合"の直前列までが判定対象
    （本フィクスチャでは1列だけ）に`table_symbol`を置く。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "SID一覧"

    sheet.cell(row=1, column=1, value="SAE J1979")
    sheet.cell(row=1, column=3, value=table_symbol)  # ラベル列+2列目が判定対象開始列
    sheet.cell(row=1, column=4, value="J1979を適用する場合")

    workbook.save(path)


def generate_no_sid_list_sheet_fixture(path):
    """「SID一覧」シートを持たないワークブックを生成する
    （check_sid_list_sheetがNGになることの確認用）。"""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "無関係なシート"
    sheet.cell(row=1, column=1, value="dummy")
    workbook.save(path)


def generate_master_definition_fixture(path, ok_sheet_has_defect=False, ng_sheet_has_defect=False):
    """
    definition_file_check.SERVER_REFERENCE_FILE_PATHが指す「マスタ定義ファイル管理ツール」の
    ローカルコピーを模したワークブックを生成する。「元表」「fcl_list_o_FI」「fcl_list_x_FI」の
    3シートを持つ。

    fcl_list_o_FI／fcl_list_x_FIはそれぞれ「HILSコード」ヘッダの直後の行にデータを1行持つ。
    `*_has_defect=True`にすると、そのシートのデータ行がA列のみ値ありB/C列が空（データ不備）に
    なる。
    """
    workbook = Workbook()
    motohyou_sheet = workbook.active
    motohyou_sheet.title = "元表"
    motohyou_sheet.cell(row=1, column=1, value="項目A")
    motohyou_sheet.cell(row=1, column=2, value="項目B")
    motohyou_sheet.cell(row=2, column=1, value="値1")
    motohyou_sheet.cell(row=2, column=2, value="値2")

    for sheet_name, has_defect in (("fcl_list_o_FI", ok_sheet_has_defect), ("fcl_list_x_FI", ng_sheet_has_defect)):
        sheet = workbook.create_sheet(sheet_name)
        sheet.cell(row=1, column=1, value=sheet_name)
        sheet.cell(row=2, column=1, value="HILSコード")
        if has_defect:
            sheet.cell(row=3, column=1, value="H001")  # B・C列が空のまま（データ不備）
        else:
            sheet.cell(row=3, column=1, value="H001")
            sheet.cell(row=3, column=2, value="設定値B")
            sheet.cell(row=3, column=3, value="設定値C")

    workbook.save(path)


def generate_sharepoint_reference_fixture(path, matches_motohyou=True):
    """
    definition_file_check.compare_motohyou_sheetsが比較する「もう1つの固定参照ファイル」
    （SharePoint上の`FI-FailSafeMatrix元表_SP.xlsx`相当）の「FI FSmatrix」シートを生成する。

    `matches_motohyou=True`なら`generate_master_definition_fixture`の「元表」シートと
    セル単位で一致する内容にする（OK判定の確認用）。Falseなら内容を変えて不一致にする
    （NG判定の確認用）。
    """
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "FI FSmatrix"

    if matches_motohyou:
        sheet.cell(row=1, column=1, value="項目A")
        sheet.cell(row=1, column=2, value="項目B")
        sheet.cell(row=2, column=1, value="値1")
        sheet.cell(row=2, column=2, value="値2")
    else:
        sheet.cell(row=1, column=1, value="項目A")
        sheet.cell(row=1, column=2, value="項目B")
        sheet.cell(row=2, column=1, value="値1")
        sheet.cell(row=2, column=2, value="値2_不一致")

    workbook.save(path)

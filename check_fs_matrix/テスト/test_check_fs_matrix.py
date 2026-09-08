"""
②(check_fs_matrix)の判定ロジック（matrix_check.py／gst_command_check.py／
definition_file_check.py）を、generate_fixtures.pyが生成する最小フィクスチャで検証する
テストスクリプト（2026-09-02新規）。

【2026-09-02改修】①②のSharePoint連携コード削除に伴い、input_check_main.py（sort_key／
merge_check2_line含む）自体を削除したため、これらに対するテストも削除した。列マージ・ソート
ロジックは方式B統合実行ツール（"..\ツール本体\"）のunified_result.pyに一本化されている。

①(check_mao)の`テスト\\test_can_mao_check.py`と同じ方針で、pytest等の追加ライブラリには
依存せず、標準のassert文で検証する。

実行方法:
    python test_check_fs_matrix.py
"""

import sys
import tempfile
from pathlib import Path

# ツール本体（matrix_check.py等）とこのフォルダ（generate_fixtures.py）の両方をインポートできる
# ようにする。
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import check2_report  # noqa: E402
import definition_file_check  # noqa: E402
import matrix_check  # noqa: E402
import gst_command_check  # noqa: E402
from generate_fixtures import (  # noqa: E402
    NG_SYMBOL_INVALID,
    generate_gst_command_fixture,
    generate_master_definition_fixture,
    generate_matrix_fixture,
    generate_no_matrix_sheet_fixture,
    generate_no_sid_list_sheet_fixture,
    generate_sharepoint_reference_fixture,
)


def test_check_matrix_sheet_ok():
    """FSマトリクスシート・部番/診断識別コード表が正しい場合にOKを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_matrix_fixture(path)
        result, message = matrix_check.check_matrix_sheet(str(path))
        assert result == "OK", f"OKになるはず（実際: {result} {message}）"
    print("test_check_matrix_sheet_ok: OK")


def test_check_matrix_sheet_parts_wording():
    """結果ログ用の2行（シートの有無／適用欄のデータ不備の有無）が、
    肯定形（ある・ない）の文で返ることを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_matrix_fixture(path)
        parts = matrix_check.check_matrix_sheet_parts(str(path))

        assert [result for result, _ in parts] == ["OK", "OK"], f"2行ともOKのはず（実際: {parts}）"
        assert parts[0][1].startswith("ツールが想定している「"), parts[0][1]
        assert parts[0][1].endswith("シートがある。"), parts[0][1]
        assert parts[1][1].endswith(f"{matrix_check.OX_FIELD_LABEL}にデータ不備がない。"), parts[1][1]
    print("test_check_matrix_sheet_parts_wording: OK")


def test_check_matrix_sheet_missing_sheet():
    """想定シート名がいずれも存在しない場合にNGを返すことを検証する。
    このときは適用欄を検査できないので、行は1行（シートの有無）だけになる。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_no_matrix_sheet_fixture(path)
        result, message = matrix_check.check_matrix_sheet(str(path))
        assert result == "NG"
        assert "シートがない。" in message, f"否定形の文になるはず（実際: {message}）"

        parts = matrix_check.check_matrix_sheet_parts(str(path))
        assert len(parts) == 1, f"適用欄の行は返さないはず（実際: {parts}）"
    print("test_check_matrix_sheet_missing_sheet: OK")


def test_check_matrix_sheet_bad_symbol():
    """部番・診断識別コード表に○×以外の値がある場合にNGを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_matrix_fixture(path, table_symbol=NG_SYMBOL_INVALID)
        result, message = matrix_check.check_matrix_sheet(str(path))
        assert result == "NG"
        assert "○/× ではありません" in message
    print("test_check_matrix_sheet_bad_symbol: OK")


def test_check_matrix_sheet_for_all_files_aggregates():
    """複数ファイルのうち1つでもNGがあれば全体NG、全てOKなら全体OKになることを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        ok_path = Path(tmp_dir) / "ok.xlsx"
        ng_path = Path(tmp_dir) / "ng.xlsx"
        generate_matrix_fixture(ok_path)
        generate_no_matrix_sheet_fixture(ng_path)

        result, message = matrix_check.check_matrix_sheet_for_all_files([str(ok_path)])
        assert result == "OK", f"単独OKファイルのみならOKのはず（実際: {result} {message}）"

        result, message = matrix_check.check_matrix_sheet_for_all_files([str(ok_path), str(ng_path)])
        assert result == "NG", f"1つでもNGがあれば全体NGのはず（実際: {result} {message}）"
        assert "ng.xlsx" in message
    print("test_check_matrix_sheet_for_all_files_aggregates: OK")


def test_check_sid_list_sheet_ok():
    """SID一覧シート・J1979判定表が正しい場合にOKを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "CommandSupportGST_test.xlsx"
        generate_gst_command_fixture(path)
        result, message = gst_command_check.check_sid_list_sheet(str(path))
        assert result == "OK", f"OKになるはず（実際: {result} {message}）"
    print("test_check_sid_list_sheet_ok: OK")


def test_check_sid_list_sheet_parts_wording():
    """結果ログ用の2行（シートの有無／適用欄のデータ不備の有無）が、
    肯定形（ある・ない）の文で返ることを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "CommandSupportGST_test.xlsx"
        generate_gst_command_fixture(path)
        parts = gst_command_check.check_sid_list_sheet_parts(str(path))

        assert [result for result, _ in parts] == ["OK", "OK"], f"2行ともOKのはず（実際: {parts}）"
        assert parts[0][1] == f"ツールが想定している「{gst_command_check.SID_LIST_SHEET_NAME}」シートがある。", parts[0][1]
        assert parts[1][1] == (
            f"「{gst_command_check.SID_LIST_SHEET_NAME}」シートの"
            f"{gst_command_check.OX_FIELD_LABEL}にデータ不備がない。"
        ), parts[1][1]
    print("test_check_sid_list_sheet_parts_wording: OK")


def test_check_sid_list_sheet_missing():
    """SID一覧シートが存在しない場合にNGを返すことを検証する。
    このときは適用欄を検査できないので、行は1行（シートの有無）だけになる。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "CommandSupportGST_test.xlsx"
        generate_no_sid_list_sheet_fixture(path)
        result, message = gst_command_check.check_sid_list_sheet(str(path))
        assert result == "NG"
        assert "シートがない。" in message, f"否定形の文になるはず（実際: {message}）"

        parts = gst_command_check.check_sid_list_sheet_parts(str(path))
        assert len(parts) == 1, f"適用欄の行は返さないはず（実際: {parts}）"
    print("test_check_sid_list_sheet_missing: OK")


def test_check_gst_command_files_skips_when_no_attachment():
    """添付が無い（空リスト）場合はOK・メッセージ無しになることを検証する
    （gst_command_check.pyのdocstring記載の「サイレントスキップ」仕様）。"""
    result, message = gst_command_check.check_gst_command_files([])
    assert (result, message) == ("OK", "")
    print("test_check_gst_command_files_skips_when_no_attachment: OK")


def test_check_gst_command_files_ignores_non_matching_names():
    """ファイル名がCommandSupportGSTパターンに一致しないファイルは無視されることを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "無関係なファイル.xlsx"
        generate_no_sid_list_sheet_fixture(path)  # NGになる内容だが、名前が一致しないので無視されるはず
        result, message = gst_command_check.check_gst_command_files([str(path)])
        assert (result, message) == ("OK", "")
    print("test_check_gst_command_files_ignores_non_matching_names: OK")


def test_compare_motohyou_sheets_ok():
    """マスタ側「元表」とSharePoint側「FI FSmatrix」の内容が一致する場合にOKを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "master.xlsm"
        sp_path = Path(tmp_dir) / "sp_reference.xlsx"
        generate_master_definition_fixture(master_path)
        generate_sharepoint_reference_fixture(sp_path, matches_motohyou=True)

        result, message = definition_file_check.compare_motohyou_sheets(str(master_path), str(sp_path))
        assert result == "OK", f"内容一致ならOKのはず（実際: {result} {message}）"
        # 結果ログの文（ワークブック名は拡張子なし・肯定形で終わる）。
        assert message == (
            "「master」の「元表」シートと"
            f"{definition_file_check.MOTOHYOU_SP_DESCRIPTION}内容が一致している。"
        ), message
    print("test_compare_motohyou_sheets_ok: OK")


def test_compare_motohyou_sheets_mismatch():
    """内容が不一致の場合にNGを返し、どのセルが違うのかをメッセージに示すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "master.xlsm"
        sp_path = Path(tmp_dir) / "sp_reference.xlsx"
        generate_master_definition_fixture(master_path)
        generate_sharepoint_reference_fixture(sp_path, matches_motohyou=False)

        result, message = definition_file_check.compare_motohyou_sheets(str(master_path), str(sp_path))
        assert result == "NG"
        assert "一致していない。" in message, f"否定形の文になるはず（実際: {message}）"
        # フィクスチャの差分はB2セル（値2 ↔ 値2_不一致）。
        assert "B2" in message, f"不一致セルの番地が含まれるはず（実際: {message}）"
        assert "値2_不一致" in message, f"不一致セルの値が含まれるはず（実際: {message}）"
    print("test_compare_motohyou_sheets_mismatch: OK")


def test_check_o_x_sheet_data_ok_and_ng():
    """fcl_list_o_FI／fcl_list_x_FIのデータ不備検出（HILSコード直後の行でB・C列が両方空）を検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        clean_path = Path(tmp_dir) / "clean.xlsm"
        defect_path = Path(tmp_dir) / "defect.xlsm"
        generate_master_definition_fixture(clean_path, ok_sheet_has_defect=False, ng_sheet_has_defect=False)
        generate_master_definition_fixture(defect_path, ok_sheet_has_defect=True, ng_sheet_has_defect=False)

        result, message = definition_file_check.check_o_x_sheet_data(str(clean_path), definition_file_check.OK_SHEET_NAME)
        assert result == "OK", f"データ不備が無ければOKのはず（実際: {result} {message}）"
        # 結果ログの文（ワークブック名は拡張子なし・肯定形で終わる）。
        assert message == f"「clean」の「{definition_file_check.OK_SHEET_NAME}」にデータ不備がない。", message

        result, message = definition_file_check.check_o_x_sheet_data(str(defect_path), definition_file_check.OK_SHEET_NAME)
        assert result == "NG", f"データ不備があればNGのはず（実際: {result} {message}）"
        assert "データ不備がある。" in message, f"否定形の文になるはず（実際: {message}）"
        # NGメッセージはExcel上の位置（A列のセル番地とHILSコードの値）を示すこと。
        # フィクスチャの不備行は3行目・HILSコード=H001。
        assert "A3" in message, f"不備セルの番地が含まれるはず（実際: {message}）"
        assert "H001" in message, f"不備行のHILSコードが含まれるはず（実際: {message}）"
    print("test_check_o_x_sheet_data_ok_and_ng: OK")


def test_check_o_x_sheet_data_reports_every_defect_row_up_to_limit():
    """不備が複数ある場合、最初の1件で打ち切らず全行を走査し、
    MAX_REPORTED_DEFECT_ROWS件までセル番地を列挙して超過分は件数で示すことを検証する。"""
    limit = definition_file_check.MAX_REPORTED_DEFECT_ROWS
    defect_count = limit + 2
    with tempfile.TemporaryDirectory() as tmp_dir:
        defect_path = Path(tmp_dir) / "defect_many.xlsm"
        generate_master_definition_fixture(defect_path, ok_sheet_has_defect=True, defect_row_count=defect_count)

        result, message = definition_file_check.check_o_x_sheet_data(str(defect_path), definition_file_check.OK_SHEET_NAME)
        assert result == "NG"
        # 不備行は3行目から連続。上限までは番地が載り、上限を超えた行は載らない。
        assert f"A{3 + limit - 1}" in message, f"上限内の不備セルは列挙されるはず（実際: {message}）"
        assert f"A{3 + limit}" not in message, f"上限を超えた不備セルは列挙しないはず（実際: {message}）"
        assert f"計{defect_count}件" in message, f"総件数が含まれるはず（実際: {message}）"
    print("test_check_o_x_sheet_data_reports_every_defect_row_up_to_limit: OK")


def test_check_definition_file_parts_end_to_end():
    """check_definition_file_partsが○×データ不備チェック2件だけを返し、[元表]を返さない
    ことをend-to-endで検証する（2026-09-04で元表比較を無効化したため）。

    check_definition_file_partsは内部でモジュール定数SERVER_REFERENCE_FILE_PATHを直接参照する
    実装（本番用ファイルへの変更は②担当者に委ねる方針のため、2026-09-02改修で引数化を取消し
    元の固定パス定数方式に戻した）。そのためテストでは同定数を一時的にフィクスチャのパスへ
    差し替える。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "master.xlsm"
        generate_master_definition_fixture(master_path)

        original_path = definition_file_check.SERVER_REFERENCE_FILE_PATH
        definition_file_check.SERVER_REFERENCE_FILE_PATH = str(master_path)
        try:
            parts = definition_file_check.check_definition_file_parts()
        finally:
            definition_file_check.SERVER_REFERENCE_FILE_PATH = original_path

        assert [label for label, _, _ in parts] == ["fcl_list_o_FI", "fcl_list_x_FI"], (
            f"元表比較は無効化済みなので○×2件のみ返るはず（実際: {parts}）"
        )
        assert all(result == "OK" for _, result, _ in parts), f"全項目OKになるはず（実際: {parts}）"
    print("test_check_definition_file_parts_end_to_end: OK")


# ---- 元表比較の無効化（2026-09-04）に伴い、以下2件のテストをコメントアウト -------------------
# どちらも「元表比較の結果によって○×シートのチェックを打ち切るか」を検証するテストで、
# 打ち切りの根拠だった元表比較そのものを止めたため成立しない。元表比較を再開する場合は
# `definition_file_check.check_definition_file_parts`のコメントアウトを戻すのと合わせて、
# この2件と`run_test()`内の呼び出しも戻すこと。
# なお`compare_motohyou_sheets`自体は残してあり、`test_compare_motohyou_sheets_ok`／
# `test_compare_motohyou_sheets_mismatch`が引き続き比較ロジックを守っている。
#
# def test_check_definition_file_parts_stops_on_motohyou_mismatch():
#     """元表が不一致の場合、[fcl_list_o_FI]／[fcl_list_x_FI]のチェックを実行せず
#     [元表]の1件だけを返して打ち切ることを検証する。
#
#     フィクスチャの○×シートはデータ不備が無い（＝打ち切らなければOK2件が返る）状態なので、
#     1件しか返らないことが「実行していない」ことの確認になる。"""
#     with tempfile.TemporaryDirectory() as tmp_dir:
#         master_path = Path(tmp_dir) / "master.xlsm"
#         sp_path = Path(tmp_dir) / "sp_reference.xlsx"
#         generate_master_definition_fixture(master_path)
#         generate_sharepoint_reference_fixture(sp_path, matches_motohyou=False)
#
#         original_path = definition_file_check.SERVER_REFERENCE_FILE_PATH
#         definition_file_check.SERVER_REFERENCE_FILE_PATH = str(master_path)
#         try:
#             parts = definition_file_check.check_definition_file_parts(str(sp_path))
#         finally:
#             definition_file_check.SERVER_REFERENCE_FILE_PATH = original_path
#
#         assert [label for label, _, _ in parts] == ["元表"], f"[元表]の1件だけになるはず（実際: {parts}）"
#         assert parts[0][1] == "NG"
#         assert "一致していない。" in parts[0][2], parts[0][2]
#     print("test_check_definition_file_parts_stops_on_motohyou_mismatch: OK")
#
#
# def test_check_definition_file_parts_continues_when_sheet_unreadable():
#     """元表がNGでも「不一致」以外の理由（SharePoint側ファイルを開けない）の場合は
#     打ち切らず、○×シートのチェックを従来どおり実行することを検証する
#     （unified_main.runの取得失敗時の前提）。"""
#     with tempfile.TemporaryDirectory() as tmp_dir:
#         master_path = Path(tmp_dir) / "master.xlsm"
#         generate_master_definition_fixture(master_path)
#         missing_sp_path = Path(tmp_dir) / "_未取得.xlsx"  # 存在しないパス
#
#         original_path = definition_file_check.SERVER_REFERENCE_FILE_PATH
#         definition_file_check.SERVER_REFERENCE_FILE_PATH = str(master_path)
#         try:
#             parts = definition_file_check.check_definition_file_parts(str(missing_sp_path))
#         finally:
#             definition_file_check.SERVER_REFERENCE_FILE_PATH = original_path
#
#         assert [label for label, _, _ in parts] == ["元表", "fcl_list_o_FI", "fcl_list_x_FI"], f"3件返るはず（実際: {parts}）"
#         assert parts[0][1] == "NG"
#         assert all(result == "OK" for _, result, _ in parts[1:]), f"○×シートは評価されOKのはず（実際: {parts}）"
#     print("test_check_definition_file_parts_continues_when_sheet_unreadable: OK")
# ------------------------------------------------------------------------------------------


def test_build_report_body_groups_lines_under_headings():
    """結果ログ本文が[定義ファイル管理Excel]／[FI FSマトリクスファイル]／[GSTコマンド装備表]の
    3グループに分かれ、各行が箇条書き記号付きの1行になることを検証する。
    行が空のグループは見出しごと出さない（GST添付なしの案件）。"""
    groups = [
        (check2_report.GROUP_DEFINITION_FILE, [("OK", "定義ファイルの行1"), ("NG", "定義ファイルの行2")]),
        (check2_report.GROUP_FS_MATRIX, [("OK", "FSマトリクスの行")]),
        (check2_report.GROUP_GST_COMMAND, []),
    ]
    body = check2_report.build_body(groups, timestamp="2026-09-03 15:00:00")

    assert body.splitlines() == [
        "実行日時: 2026-09-03 15:00:00",
        "",
        f"[{check2_report.GROUP_DEFINITION_FILE}]",
        "・定義ファイルの行1",
        "・定義ファイルの行2",
        "",
        f"[{check2_report.GROUP_FS_MATRIX}]",
        "・FSマトリクスの行",
    ], body
    print("test_build_report_body_groups_lines_under_headings: OK")


def test_build_report_body_end_to_end():
    """フィクスチャ3種（定義ファイル管理Excel／FSマトリクス／GSTコマンド装備表）を全てOKの
    状態でチェックし、結果ログ本文が想定の書式どおりに並ぶことをend-to-endで検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "XPX定義ファイル管理_5.22_FI_ELEC.xlsm"
        matrix_path = Path(tmp_dir) / "SYSTEM_MATRIX.xlsx"
        gst_path = Path(tmp_dir) / "CommandSupportGST_test.xlsx"
        generate_master_definition_fixture(master_path)
        # 実在シート名はMATRIX_SHEET_NAME_VARIANTSの範囲で揺れる（結果ログには見つかった
        # 実名が出る）。ここは本番の提出物と同じ「FI FSmatrix」で生成する。
        generate_matrix_fixture(matrix_path, sheet_name=matrix_check.CANONICAL_SHEET_NAME)
        generate_gst_command_fixture(gst_path)

        original_path = definition_file_check.SERVER_REFERENCE_FILE_PATH
        definition_file_check.SERVER_REFERENCE_FILE_PATH = str(master_path)
        try:
            definition_parts = definition_file_check.check_definition_file_parts()
        finally:
            definition_file_check.SERVER_REFERENCE_FILE_PATH = original_path

        groups = [
            (
                check2_report.GROUP_DEFINITION_FILE,
                [(result, message) for _label, result, message in definition_parts],
            ),
            (check2_report.GROUP_FS_MATRIX, matrix_check.check_matrix_sheet_parts_for_all_files([str(matrix_path)])),
            (check2_report.GROUP_GST_COMMAND, gst_command_check.check_gst_command_files_parts([str(gst_path)])),
        ]
        body = check2_report.build_body(groups, timestamp="2026-09-03 15:00:00")

        assert body.splitlines() == [
            "実行日時: 2026-09-03 15:00:00",
            "",
            "[定義ファイル管理Excel]",
            # 元表比較は2026-09-04に無効化したため、[元表]の行は出ない。
            "・「XPX定義ファイル管理_5.22_FI_ELEC」の「fcl_list_o_FI」にデータ不備がない。",
            "・「XPX定義ファイル管理_5.22_FI_ELEC」の「fcl_list_x_FI」にデータ不備がない。",
            "",
            "[FI FSマトリクスファイル]",
            "・ツールが想定している「FI FSmatrix」シートがある。",
            "・「FI FSmatrix」シートの適用（〇/×）欄にデータ不備がない。",
            "",
            "[GSTコマンド装備表]",
            "・ツールが想定している「SID一覧」シートがある。",
            "・「SID一覧」シートの適用（〇/×）欄にデータ不備がない。",
        ], body
    print("test_build_report_body_end_to_end: OK")


def run_test():
    """本ファイルの全テストを実行する。"""
    test_check_matrix_sheet_ok()
    test_check_matrix_sheet_parts_wording()
    test_check_matrix_sheet_missing_sheet()
    test_check_matrix_sheet_bad_symbol()
    test_check_matrix_sheet_for_all_files_aggregates()
    test_check_sid_list_sheet_ok()
    test_check_sid_list_sheet_parts_wording()
    test_check_sid_list_sheet_missing()
    test_check_gst_command_files_skips_when_no_attachment()
    test_check_gst_command_files_ignores_non_matching_names()
    test_compare_motohyou_sheets_ok()
    test_compare_motohyou_sheets_mismatch()
    test_check_o_x_sheet_data_ok_and_ng()
    test_check_o_x_sheet_data_reports_every_defect_row_up_to_limit()
    test_check_definition_file_parts_end_to_end()
    # 元表比較の無効化（2026-09-04）でコメントアウト。再開時に戻すこと。
    # test_check_definition_file_parts_stops_on_motohyou_mismatch()
    # test_check_definition_file_parts_continues_when_sheet_unreadable()
    test_build_report_body_groups_lines_under_headings()
    test_build_report_body_end_to_end()
    print("すべてのテストに成功しました。")


if __name__ == "__main__":
    run_test()

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


def test_check_matrix_sheet_missing_sheet():
    """想定シート名がいずれも存在しない場合にNGを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_no_matrix_sheet_fixture(path)
        result, message = matrix_check.check_matrix_sheet(str(path))
        assert result == "NG"
        assert "見つからない" in message
    print("test_check_matrix_sheet_missing_sheet: OK")


def test_check_matrix_sheet_bad_symbol():
    """部番・診断識別コード表に○×以外の値がある場合にNGを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "matrix.xlsx"
        generate_matrix_fixture(path, table_symbol=NG_SYMBOL_INVALID)
        result, message = matrix_check.check_matrix_sheet(str(path))
        assert result == "NG"
        assert "〇×ではない値" in message
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


def test_check_sid_list_sheet_missing():
    """SID一覧シートが存在しない場合にNGを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        path = Path(tmp_dir) / "CommandSupportGST_test.xlsx"
        generate_no_sid_list_sheet_fixture(path)
        result, message = gst_command_check.check_sid_list_sheet(str(path))
        assert result == "NG"
        assert "見つからない" in message
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
    print("test_compare_motohyou_sheets_ok: OK")


def test_compare_motohyou_sheets_mismatch():
    """内容が不一致の場合にNGを返すことを検証する。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "master.xlsm"
        sp_path = Path(tmp_dir) / "sp_reference.xlsx"
        generate_master_definition_fixture(master_path)
        generate_sharepoint_reference_fixture(sp_path, matches_motohyou=False)

        result, message = definition_file_check.compare_motohyou_sheets(str(master_path), str(sp_path))
        assert result == "NG"
        assert "一致しません" in message
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

        result, message = definition_file_check.check_o_x_sheet_data(str(defect_path), definition_file_check.OK_SHEET_NAME)
        assert result == "NG", f"データ不備があればNGのはず（実際: {result} {message}）"
    print("test_check_o_x_sheet_data_ok_and_ng: OK")


def test_check_definition_file_parts_end_to_end():
    """check_definition_file_parts（元表比較＋○×データ不備チェック2種）が3項目の
    (label, result, message)タプルを返すことをend-to-endで検証する。

    check_definition_file_partsは内部でモジュール定数SERVER_REFERENCE_FILE_PATHを直接参照する
    実装（本番用ファイルへの変更は②担当者に委ねる方針のため、2026-09-02改修で引数化を取消し
    元の固定パス定数方式に戻した）。そのためテストでは同定数を一時的にフィクスチャのパスへ
    差し替える。"""
    with tempfile.TemporaryDirectory() as tmp_dir:
        master_path = Path(tmp_dir) / "master.xlsm"
        sp_path = Path(tmp_dir) / "sp_reference.xlsx"
        generate_master_definition_fixture(master_path)
        generate_sharepoint_reference_fixture(sp_path, matches_motohyou=True)

        original_path = definition_file_check.SERVER_REFERENCE_FILE_PATH
        definition_file_check.SERVER_REFERENCE_FILE_PATH = str(master_path)
        try:
            parts = definition_file_check.check_definition_file_parts(str(sp_path))
        finally:
            definition_file_check.SERVER_REFERENCE_FILE_PATH = original_path

        assert [label for label, _, _ in parts] == ["元表", "fcl_list_o_FI", "fcl_list_x_FI"]
        assert all(result == "OK" for _, result, _ in parts), f"全項目OKになるはず（実際: {parts}）"
    print("test_check_definition_file_parts_end_to_end: OK")


def run_test():
    """本ファイルの全テストを実行する。"""
    test_check_matrix_sheet_ok()
    test_check_matrix_sheet_missing_sheet()
    test_check_matrix_sheet_bad_symbol()
    test_check_matrix_sheet_for_all_files_aggregates()
    test_check_sid_list_sheet_ok()
    test_check_sid_list_sheet_missing()
    test_check_gst_command_files_skips_when_no_attachment()
    test_check_gst_command_files_ignores_non_matching_names()
    test_compare_motohyou_sheets_ok()
    test_compare_motohyou_sheets_mismatch()
    test_check_o_x_sheet_data_ok_and_ng()
    test_check_definition_file_parts_end_to_end()
    print("すべてのテストに成功しました。")


if __name__ == "__main__":
    run_test()

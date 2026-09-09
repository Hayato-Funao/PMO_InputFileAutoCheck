"""
判定ロジック（compare.ComparisonRow.judgment）・結果出力（result.build_check_result等）・
案件フォルダ走査（case_scan.scan_case_folder）を、generate_fixtures.pyが生成する
最小フィクスチャで検証するテストスクリプト。

【2026-09-01改修】チェック対象をCANテーブルからCANマトリクスへ変更したことに伴い、①の判定に
関わるテスト（判定ロジック・結果出力・案件フォルダ走査）の期待値・フィクスチャをマトリクス基準へ
更新し、マトリクスのヘッダ位置ばらつき検出・バス使用率行/ゲートウェイ経路除外のテストを追加した。
CANテーブル関連のテスト（`_detect_columns`等。当面未使用の`can_info.py`のCANテーブル関数に対する
レガシーテスト）はそのまま残している。

pytest等の追加ライブラリには依存せず、標準のassert文で検証する
（本ツール本体が標準ライブラリのみで動作する方針に合わせる）。

実行方法:
	python test_can_mao_check.py
"""

import shutil
import sys
import tempfile
from pathlib import Path

# ツール本体（main.py等）とこのフォルダ（generate_fixtures.py）の両方をインポートできるようにする
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from can_info import (  # noqa: E402
	extract_matrix_can_id_records,
	find_matrix_header_location,
	select_matrix_sheet,
)
from can_table_legacy import (  # noqa: E402
	_detect_columns,
	extract_table_can_id_records,
	find_can_table_header_row,
	select_can_table_sheet,
)
from case_scan import (  # noqa: E402
	derive_case_id,
	is_valid_case_id_format,
	resolve_input_folder_path,
	scan_case_folder,
)
from compare import compare_can_id_sources  # noqa: E402
from generate_fixtures import (  # noqa: E402
	generate_3daa_cdc_style_matrix_fixture,
	generate_a2l_fixture,
	generate_ambiguous_multi_sheet_table_fixture,
	generate_case_folder_fixture,
	generate_case_folder_fixture_with_a2l_fallback,
	generate_fi_icm_case_folder_fixture,
	generate_fixtures,
	generate_mao_fixture,
	generate_matrix_fixture,
	generate_multi_sheet_matrix_fixture,
	generate_multi_sheet_table_fixture,
	generate_table_fixture,
	zip_directory_contents,
)
from mao_can_id import extract_a2l_can_id_records, extract_mao_can_id_records  # noqa: E402
from result import (  # noqa: E402
	MaskCheckResult,
	build_case_result,
	build_check_result,
	format_case_result_message,
	sort_key,
)

# generate_fixtures.pyのdocstring記載の期待判定（内部仕様書4.5節。MAO＋CANマトリクスの2ソース版）
EXPECTED_JUDGMENT_BY_CAN_ID = {
	256: "OK",  # 100h: MAO+マトリクス
	512: "NG",  # 200h: MAOのみ（マトリクス欠け）
	768: "OK",  # 300h: MAO+マトリクス
	1024: "NG",  # 400h: MAOのみ（マトリクス欠け）
	1280: "－",  # 500h: マトリクスのみ（MAOに無い）
}


def test_judgment_and_check_result():
	"""判定ロジック（2ソース版）とチェック結果の組み立てを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture_paths = generate_fixtures(temporary_directory)

		mao_records = extract_mao_can_id_records(fixture_paths["mao"])
		matrix_records = extract_matrix_can_id_records(fixture_paths["matrix"])

		comparison_rows = compare_can_id_sources(mao_records, matrix_records)
		judgment_by_can_id = {row.can_id_decimal: row.judgment for row in comparison_rows}

		for can_id_decimal, expected_judgment in EXPECTED_JUDGMENT_BY_CAN_ID.items():
			actual_judgment = judgment_by_can_id.get(can_id_decimal)
			assert actual_judgment == expected_judgment, (
				f"CAN ID {can_id_decimal} の判定が想定と異なる"
				f"（期待:{expected_judgment} 実際:{actual_judgment}）"
			)

		check_result = build_check_result(comparison_rows)
		assert check_result.overall_judgment == "NG", "NGが存在するため全体判定はNGになるはず"
		assert len(check_result.ng_details) == 2, f"NG明細は2件のはず（実際:{len(check_result.ng_details)}件）"

		# CAN ID件数の集計値（総数5件: OK2/NG2/－1。2026-08-31追補）
		assert check_result.total_count == 5, f"総数は5件のはず（実際:{check_result.total_count}件）"
		assert check_result.ok_count == 2, f"OK件数は2件のはず（実際:{check_result.ok_count}件）"
		assert check_result.dash_count == 1, f"－件数は1件のはず（実際:{check_result.dash_count}件）"

		ng_can_id_decimals = {detail.can_id_decimal for detail in check_result.ng_details}
		assert ng_can_id_decimals == {512, 1024}, f"NGとなるCAN IDが想定と異なる（実際:{sorted(ng_can_id_decimals)}）"

		# 欠けている出典は、CANテーブルを対象外としたため常に["CANマトリクス"]のみ（2026-09-01改修）
		for detail in check_result.ng_details:
			assert detail.missing_sources == ["CANマトリクス"], (
				f"CAN ID {detail.can_id_decimal} の欠けている出典が想定と異なる（実際:{detail.missing_sources}）"
			)

		print("test_judgment_and_check_result: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_extract_matrix_can_id_records_excludes_bus_summary_and_gateway_route():
	"""
	CANマトリクスのバス使用率集計行（CAN ID列="F1"）・ゲートウェイ経路表記（"F1->F4,F5"）が
	いずれもCAN IDとして読み飛ばされることを検証する（2026-09-01新規）。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		matrix_path = temporary_directory / "fixture_Matrix.xlsx"
		generate_matrix_fixture(
			matrix_path,
			include_bus_summary_row=True,
			include_gateway_route_row=True,
		)

		records = extract_matrix_can_id_records(matrix_path)
		can_id_decimals = {record.can_id_decimal for record in records}

		assert can_id_decimals == {256, 768, 1280}, (
			f"バス使用率行・GW経路行は除外され、100h/300h/500hのみ抽出されるはず（実際:{sorted(can_id_decimals)}）"
		)

		print("test_extract_matrix_can_id_records_excludes_bus_summary_and_gateway_route: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_find_matrix_header_location_detects_various_positions():
	"""
	CANマトリクスのヘッダ（`CAN ID (HEX)`列）が、機種によるヘッダ行・列位置のばらつき
	（旧仕様のA列固定・R1〜R3ヘッダ固定では対応できないケース）でも動的に検出できることを
	検証する（2026-09-01新規。フォーマットのばらつきへの対応確認）。
	"""
	# 旧仕様どおり：ヘッダR1・A列
	rows_header_at_1_column_a = [["CAN ID (HEX)", "GW Direction"], ["100", "F1->F4"]]
	assert find_matrix_header_location(rows_header_at_1_column_a) == (1, 1)

	# ばらつき例：ヘッダR3・C列（旧仕様の固定値から外れる機種を想定）
	rows_header_at_3_column_c = [
		[None, None, None],
		[None, None, None],
		[None, None, "CAN ID (HEX)"],
		[None, None, "100"],
	]
	assert find_matrix_header_location(rows_header_at_3_column_c) == (3, 3)

	# ヘッダが見つからない場合はNone
	rows_without_header = [["Message Name", "Length"] for _ in range(5)]
	assert find_matrix_header_location(rows_without_header) is None

	print("test_find_matrix_header_location_detects_various_positions: OK")


def test_extract_matrix_can_id_records_supports_3daa_cdc_style():
	"""
	3DAA(CDC)形式（ファイル名は"Matrix"を含むが内部列構成がID・Transmitter・Receiversという
	CANテーブルに近い構造。実データ202608_006_01で確認）でも、Transmitter/Receiversヘッダの
	フォールバック検出によりCAN IDを抽出できることを検証する（2026-09-01追補）。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		matrix_path = temporary_directory / "Matrix_3daa_cdc.xlsx"
		generate_3daa_cdc_style_matrix_fixture(matrix_path)

		records = extract_matrix_can_id_records(matrix_path)
		can_id_decimals = {record.can_id_decimal for record in records}

		assert can_id_decimals == {256, 768}, (
			f"ID/Transmitter/Receivers構成でも100h/300hが抽出されるはず（実際:{sorted(can_id_decimals)}）"
		)

		print("test_extract_matrix_can_id_records_supports_3daa_cdc_style: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_select_matrix_sheet_returns_none_when_ambiguous():
	"""
	1ブックに`CAN ID (HEX)`ヘッダを持つシートが複数ある場合、CANテーブルのような絞り込み
	キーワードが無いため、誤ったシートを選ばずNone（判定不可）を返すことを検証する（2026-09-01新規）。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		matrix_path = temporary_directory / "ambiguous_Matrix.xlsx"
		generate_multi_sheet_matrix_fixture(matrix_path)

		assert select_matrix_sheet(matrix_path) is None

		print("test_select_matrix_sheet_returns_none_when_ambiguous: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_basic():
	"""正常な案件フォルダ構造で、MAOが無いマスクを除外し"Matrix"を含むファイルを採用できることを検証する
	（2026-09-01改修。旧仕様とは逆に、CANテーブル形式のdecoyが除外されることを確認する）。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture(temporary_directory)

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			assert scan_result.work_directory is None, "zipが無いのでwork_directoryはNoneのはず"
			assert len(scan_result.structural_ng_messages) == 0, (
				f"構造的NGは無いはず（実際:{scan_result.structural_ng_messages}）"
			)
			assert len(scan_result.mask_entries) == 1, (
				f"MAOがあるマスクのみ1件のはず（実際:{len(scan_result.mask_entries)}件）"
			)
			assert scan_result.mask_entries[0].label == "Mask1_HEX(A)"
			assert scan_result.matrix_path == fixture["matrix_path"], (
				"CANテーブル形式のdecoyを除外し、matrix_No1.xlsxのみがCANマトリクスとして採用されるはず"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_basic: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_duplicate_mao_is_structural_ng():
	"""1マスクフォルダ内にMAOが2個ある場合、そのマスクは構造的NGとして記録され判定対象から外れる。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture(temporary_directory)
		generate_mao_fixture(fixture["mask1_dir"] / "fixture2.mao")

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			assert len(scan_result.mask_entries) == 0, "重複MAOのマスクは判定対象から除外されるはず"
			assert any("MAO" in message for message in scan_result.structural_ng_messages), (
				f"MAO重複の構造的NGが記録されるはず（実際:{scan_result.structural_ng_messages}）"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_duplicate_mao_is_structural_ng: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_duplicate_matrix_is_structural_ng():
	"""02_CAN関連にファイル名"Matrix"を含む候補が2個以上ある場合、構造的NGとしてマトリクス未特定になる
	（2026-09-01改修。旧`test_case_scan_duplicate_table_is_structural_ng`のマトリクス版）。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture(temporary_directory)
		generate_matrix_fixture(fixture["matrix_path"].parent / "matrix_No2.xlsx")

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			assert scan_result.matrix_path is None, "CANマトリクス候補が2個になるため特定できないはず"
			assert any("複数存在" in message for message in scan_result.structural_ng_messages), (
				f"マトリクス重複の構造的NGが記録されるはず（実際:{scan_result.structural_ng_messages}）"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_duplicate_matrix_is_structural_ng: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_extracts_zip():
	"""案件フォルダがzipで共有された場合、自動的に解凍してから走査できることを検証する。"""
	source_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_src_"))
	outer_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_outer_"))
	try:
		fixture = generate_case_folder_fixture(source_directory)
		zip_directory_contents(fixture["input_root"], outer_directory / "案件一式.zip")

		scan_result = scan_case_folder(outer_directory)
		try:
			assert scan_result.work_directory is not None, "zipを展開したのでwork_directoryが作られるはず"
			assert len(scan_result.mask_entries) == 1, (
				f"zip展開後もMAOがあるマスクを検出できるはず（実際:{len(scan_result.mask_entries)}件）"
			)
			assert scan_result.matrix_path is not None, "zip展開後もCANマトリクスを特定できるはず"
			assert len(scan_result.structural_ng_messages) == 0
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_extracts_zip: OK")
	finally:
		shutil.rmtree(source_directory, ignore_errors=True)
		shutil.rmtree(outer_directory, ignore_errors=True)


def test_build_case_result_aggregates_masks_and_structural_ng():
	"""複数マスクの結果と構造的NGを1つのCaseCheckResultへ集約する挙動を検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture_paths = generate_fixtures(temporary_directory)
		mao_records = extract_mao_can_id_records(fixture_paths["mao"])
		matrix_records = extract_matrix_can_id_records(fixture_paths["matrix"])
		comparison_rows = compare_can_id_sources(mao_records, matrix_records)
		check_result = build_check_result(comparison_rows)

		mask_results = [
			MaskCheckResult("Mask1_HEX(A)", check_result=check_result),
			MaskCheckResult("Mask2_HEX(B)", skipped_reason="CANマトリクスが特定できないため判定不可"),
		]
		case_result = build_case_result(mask_results, structural_ng_messages=["ダミーの構造的NG"])

		assert case_result.overall_judgment == "NG", "マスクNGまたは構造的NGがあれば全体NGのはず"
		assert len(case_result.mask_results) == 2
		assert len(case_result.structural_ng_messages) == 1

		print("test_build_case_result_aggregates_masks_and_structural_ng: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_format_case_result_message_shows_count_line():
	"""CAN ID件数の行（総数/OK/NG/－）がマスク単位のメッセージに含まれることを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture_paths = generate_fixtures(temporary_directory)
		mao_records = extract_mao_can_id_records(fixture_paths["mao"])
		matrix_records = extract_matrix_can_id_records(fixture_paths["matrix"])
		comparison_rows = compare_can_id_sources(mao_records, matrix_records)
		check_result = build_check_result(comparison_rows)

		message = format_case_result_message(
			build_case_result([MaskCheckResult("Mask1_HEX(A)", check_result=check_result)], [])
		)
		assert "CAN ID件数: 総数5件 / OK 2件 / NG 2件 / － 1件" in message, (
			f"件数行が想定と異なる（実際のメッセージ:\n{message}）"
		)

		print("test_format_case_result_message_shows_count_line: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_format_case_result_message_omits_heading_for_single_mask():
	"""マスクが1件だけの場合は「=== マスク: ... ===」見出しを出さないことを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture_paths = generate_fixtures(temporary_directory)
		mao_records = extract_mao_can_id_records(fixture_paths["mao"])
		matrix_records = extract_matrix_can_id_records(fixture_paths["matrix"])
		comparison_rows = compare_can_id_sources(mao_records, matrix_records)
		check_result = build_check_result(comparison_rows)

		single_mask_message = format_case_result_message(
			build_case_result([MaskCheckResult("01_HEX関連", check_result=check_result)], [])
		)
		assert "=== マスク:" not in single_mask_message, (
			f"マスクが1件のみの場合は見出しを出さないはず（実際のメッセージ:\n{single_mask_message}）"
		)
		assert "CAN ID件数:" in single_mask_message, "件数行自体は単一マスクでも表示されるはず"

		multi_mask_message = format_case_result_message(
			build_case_result(
				[
					MaskCheckResult("Mask1_HEX(A)", check_result=check_result),
					MaskCheckResult("Mask2_HEX(B)", check_result=check_result),
				],
				[],
			)
		)
		assert "=== マスク: Mask1_HEX(A) ===" in multi_mask_message, "マスクが複数なら見出しを出すはず"
		assert "=== マスク: Mask2_HEX(B) ===" in multi_mask_message, "マスクが複数なら見出しを出すはず"

		print("test_format_case_result_message_omits_heading_for_single_mask: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_derive_case_id_from_year_month_structure():
	"""実データと同型（{年度}年度/{月}月/{案件ID}/01_INPUT/...）から案件IDを抽出できることを検証する。"""
	path = Path("2026年度") / "08月" / "202608_005_01" / "01_INPUT" / "01_HEX関連" / "Mask1_HEX(A)"
	assert derive_case_id(path) == "202608_005_01"
	assert is_valid_case_id_format(derive_case_id(path)) is True
	print("test_derive_case_id_from_year_month_structure: OK")


def test_derive_case_id_from_input_or_case_root():
	"""
	01_INPUT自身を渡した場合・案件ルートを渡した場合のいずれでも同じ案件IDになることを検証する。

	案件ルート（01_INPUTを含まないパス）は、2026-08-31追補のフォールバック（フォルダ名が
	案件ID書式に一致する要素を探す）で判別できることを検証する。
	"""
	case_root = Path("2026年度") / "08月" / "202608_005_01"
	assert derive_case_id(case_root / "01_INPUT") == "202608_005_01"
	assert derive_case_id(case_root) == "202608_005_01", "01_INPUTが無くてもフォルダ名から判別できるはず"
	print("test_derive_case_id_from_input_or_case_root: OK")


def test_derive_case_id_fallback_for_branch_case_without_01_input():
	"""
	枝番_02以降のように01_INPUTを持たない従属案件でも、フォルダ名から案件IDを判別できることを
	検証する（本番想定データで多数確認された構成。2026-08-31追補）。
	"""
	path = Path("2026年度") / "06月" / "202606_001_02"
	assert derive_case_id(path) == "202606_001_02"
	assert is_valid_case_id_format(derive_case_id(path)) is True
	print("test_derive_case_id_fallback_for_branch_case_without_01_input: OK")


def test_derive_case_id_returns_none_without_01_input():
	"""パスに01_INPUTが含まれない場合はNoneを返すことを検証する。"""
	assert derive_case_id(Path("何らかのフォルダ") / "サブフォルダ") is None
	print("test_derive_case_id_returns_none_without_01_input: OK")


def test_derive_case_id_flags_unexpected_format():
	"""想定外フォーマットの名前でも候補自体は返し、is_valid_case_id_formatがFalseになることを検証する。"""
	path = Path("テスト案件") / "01_INPUT" / "01_HEX関連"
	case_id = derive_case_id(path)
	assert case_id == "テスト案件"
	assert is_valid_case_id_format(case_id) is False
	print("test_derive_case_id_flags_unexpected_format: OK")


def test_find_can_table_header_row_detects_various_positions():
	"""
	ヘッダ行（A列=Transmitter・B列=Receivers）の検出が、R21（大多数の機種）・R1（3DAA(CDC)。
	実データ確認済み）のいずれでも機能し、見出しが無いシートではNoneを返すことを検証する。
	"""
	rows_with_header_at_21 = [[None, None] for _ in range(20)] + [["Transmitter", "Receivers"]]
	assert find_can_table_header_row(rows_with_header_at_21) == 21

	rows_with_header_at_1 = [["Transmitter", "Receivers", "Signal Name"]]
	assert find_can_table_header_row(rows_with_header_at_1) == 1

	rows_without_header = [["変更日", "変更内容"] for _ in range(30)]
	assert find_can_table_header_row(rows_without_header) is None

	print("test_find_can_table_header_row_detects_various_positions: OK")


def test_detect_columns_uses_header_names_and_handles_missing_id_format():
	"""
	列位置をヘッダの見出し文字列で動的に特定できること、および3DAA(CDC)の実データで確認された
	「ID列が8列目でない」「ID-Format列が存在しない」ケースでも、固定列(H列/I列)を誤読せずに
	正しく処理できることを検証する（2026-08-31追補。列固定読みだと発生する誤読の再現・回避確認）。
	"""
	# シート5型（3DAA(CDC)実データ）: Signal Name(変更後)列が挿入され、ID=9列/ID-Format=10列にずれる
	header_shifted = [
		"Transmitter",
		"Receivers",
		"Signal Name",
		"Signal Name(変更後)",
		"PDU ID",
		"Message",
		"Frame Version",
		"Frame Version2",
		"ID",
		"ID-Format",
	]
	columns_shifted = _detect_columns(header_shifted)
	assert columns_shifted["can_id"] == 9, "ID列は見出し名から9列目と特定されるはず"
	assert columns_shifted["id_format"] == 10, "ID-Format列は見出し名から10列目と特定されるはず"

	# シート8型（3DAA(CDC)実データ）: ID-Format列が存在せず、9列目はTx Method
	header_missing_id_format = [
		"Transmitter",
		"Receivers",
		"Signal Name",
		"Signal Version",
		"PDU ID",
		"Message",
		"Frame Version",
		"ID",
		"Tx Method",
	]
	columns_missing = _detect_columns(header_missing_id_format)
	assert columns_missing["can_id"] == 8
	assert columns_missing["id_format"] is None, "ID-Format列が無い場合はTx Method列を誤読せずNoneのはず"

	print("test_detect_columns_uses_header_names_and_handles_missing_id_format: OK")


def test_select_can_table_sheet_picks_unique_communication_sheet():
	"""
	1ブックに複数のTransmitter/Receiversヘッダシートがある場合（3DAA(CDC)類似構成）、シート名に
	"can"を含む本体シートを一意に選べることを検証する（変更履歴シート・作業用シートは除外される）。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		table_path = temporary_directory / "cdc_like.xlsx"
		generate_multi_sheet_table_fixture(table_path)

		selected = select_can_table_sheet(table_path)
		assert selected is not None, "本体シートを一意に選べるはず"
		assert selected[0] == "xFCAN1", f"本体シートはxFCAN1が選ばれるはず（実際:{selected[0]}）"
		assert selected[1] == 1

		print("test_select_can_table_sheet_picks_unique_communication_sheet: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_select_can_table_sheet_returns_none_when_ambiguous():
	"""シート名で一意に絞り込めない場合、誤ったシートを選ばずNone（判定不可）を返すことを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		table_path = temporary_directory / "ambiguous.xlsx"
		generate_ambiguous_multi_sheet_table_fixture(table_path)

		assert select_can_table_sheet(table_path) is None

		print("test_select_can_table_sheet_returns_none_when_ambiguous: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_extract_a2l_can_id_records_reads_can_id_but_not_protocol_ids():
	"""
	A2L代替取得の抽出ロジックを検証する（2026-09追補）。MEASUREMENT／CHARACTERISTIC説明文の
	"CAN ID=100h"（空白あり）・"CANID=300h"（空白なし）はいずれも拾い、"/begin CAN"ブロックや
	"XCP_ON_CAN"の"CAN_ID_MASTER 0x..."（"="を使わない別表記）は拾わないことを確認する。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		a2l_path = temporary_directory / "fixture.a2l"
		generate_a2l_fixture(a2l_path)

		records = extract_a2l_can_id_records(a2l_path)
		can_id_decimals = {record.can_id_decimal for record in records}

		assert can_id_decimals == {256, 768}, (
			f"MEASUREMENT/CHARACTERISTICのCAN ID=100h/300hのみ抽出されるはず"
			f"（実際:{sorted(can_id_decimals)}）。CAN_ID_MASTER等の診断/計測通信IDが"
			f"混入していないか確認すること"
		)

		print("test_extract_a2l_can_id_records_reads_can_id_but_not_protocol_ids: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_extract_a2l_can_id_records_supports_cp932_fallback():
	"""A2LがUTF-8でデコードできない（cp932等）場合でも、フォールバックで抽出できることを確認する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		a2l_path = temporary_directory / "fixture_cp932.a2l"
		generate_a2l_fixture(a2l_path, encoding="cp932")

		records = extract_a2l_can_id_records(a2l_path)
		can_id_decimals = {record.can_id_decimal for record in records}

		assert can_id_decimals == {256, 768}, (
			f"cp932エンコーディングでも抽出できるはず（実際:{sorted(can_id_decimals)}）"
		)

		print("test_extract_a2l_can_id_records_supports_cp932_fallback: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_uses_a2l_when_mao_missing():
	"""
	MAOが無くA2Lのみのマスクフォルダでは、A2Lを代替取得元として採用することを検証する
	（2026-09追補。MAOがあるマスクは従来どおりMAOを採用する）。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture_with_a2l_fallback(temporary_directory)

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			assert len(scan_result.structural_ng_messages) == 0, (
				f"構造的NGは無いはず（実際:{scan_result.structural_ng_messages}）"
			)
			entries_by_label = {entry.label: entry for entry in scan_result.mask_entries}
			assert set(entries_by_label) == {"Mask1_HEX(A)", "Mask2_HEX(B)"}, (
				f"MAOのマスク・A2Lのマスクの両方が検出されるはず（実際:{sorted(entries_by_label)}）"
			)
			assert entries_by_label["Mask1_HEX(A)"].source_kind == "MAO"
			assert entries_by_label["Mask2_HEX(B)"].source_kind == "A2L", (
				"MAOが無いマスクではA2Lが代替取得元になるはず"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_uses_a2l_when_mao_missing: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_prefers_mao_when_both_present():
	"""1マスクフォルダにMAOとA2Lの両方がある場合、MAOが優先され取得元になることを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture(temporary_directory)
		generate_a2l_fixture(fixture["mask1_dir"] / "fixture.a2l")

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			assert len(scan_result.mask_entries) == 1
			assert scan_result.mask_entries[0].source_kind == "MAO", (
				"MAOとA2Lが両方ある場合はMAOが優先されるはず"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_prefers_mao_when_both_present: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_duplicate_a2l_is_structural_ng():
	"""MAOが無くA2Lが複数存在する場合、A2L重複として構造的NGになることを検証する。"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_case_folder_fixture_with_a2l_fallback(temporary_directory)
		generate_a2l_fixture(fixture["mask2_dir"] / "fixture2.a2l")

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			labels = {entry.label for entry in scan_result.mask_entries}
			assert "Mask2_HEX(B)" not in labels, "A2L重複のマスクは判定対象から除外されるはず"
			assert any("A2L" in message for message in scan_result.structural_ng_messages), (
				f"A2L重複の構造的NGが記録されるはず（実際:{scan_result.structural_ng_messages}）"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_duplicate_a2l_is_structural_ng: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_case_scan_fi_icm_labels_do_not_collide():
	"""
	FI/ICM配下にそれぞれ同名のMask1_HEX(A)フォルダがある構成（202606_005_01・202608_006_01で
	実データ確認済み）で、マスクラベルが相対パス化され重複しないことを検証する。
	"""
	temporary_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_test_"))
	try:
		fixture = generate_fi_icm_case_folder_fixture(temporary_directory)

		scan_result = scan_case_folder(fixture["input_root"])
		try:
			labels = sorted(entry.label for entry in scan_result.mask_entries)
			assert labels == ["FI/Mask1_HEX(A)", "ICM/Mask1_HEX(A)"], (
				f"FI/ICMでラベルが重複せず区別されるはず（実際:{labels}）"
			)
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)

		print("test_case_scan_fi_icm_labels_do_not_collide: OK")
	finally:
		shutil.rmtree(temporary_directory, ignore_errors=True)


def test_resolve_input_folder_path_from_input_itself():
	"""01_INPUT自身を渡した場合、同じパスがそのまま返ることを検証する。"""
	path = Path("2026年度") / "08月" / "202608_005_01" / "01_INPUT"
	assert resolve_input_folder_path(path) == path
	print("test_resolve_input_folder_path_from_input_itself: OK")


def test_resolve_input_folder_path_from_case_root():
	"""案件ルート（01_INPUTを含まないパス）からは組み立てられず、Noneを返すことを検証する
	（01_INPUT自体がパスの構成要素に無いため）。"""
	case_root = Path("2026年度") / "08月" / "202608_005_01"
	assert resolve_input_folder_path(case_root) is None
	print("test_resolve_input_folder_path_from_case_root: OK")


def test_resolve_input_folder_path_from_deeper_folder():
	"""01_HEX関連等、01_INPUTより深いフォルダを渡した場合でも、01_INPUTまでのパスを組み立てられる
	ことを検証する。"""
	path = Path("2026年度") / "08月" / "202608_005_01" / "01_INPUT" / "01_HEX関連" / "Mask1_HEX(A)"
	expected = Path("2026年度") / "08月" / "202608_005_01" / "01_INPUT"
	assert resolve_input_folder_path(path) == expected
	print("test_resolve_input_folder_path_from_deeper_folder: OK")


def test_resolve_input_folder_path_returns_none_without_01_input():
	"""パスに01_INPUTが含まれない場合はNoneを返すことを検証する（枝番案件等）。"""
	assert resolve_input_folder_path(Path("2026年度") / "06月" / "202606_001_02") is None
	print("test_resolve_input_folder_path_returns_none_without_01_input: OK")


def test_build_case_result_returns_taishougai_when_no_masks():
	"""
	マスク（照合対象のMAO）が1件も無い場合、判定OK/NGではなく「対象外」になることを検証する
	（枝番案件・空マスク・zip納品でMAO0件等。2026-08-31追補）。
	"""
	case_result = build_case_result([], ["01_HEX関連フォルダが見つからない", "02_CAN関連フォルダが見つからない"])
	assert case_result.overall_judgment == "対象外"

	message = format_case_result_message(case_result)
	assert "対象外" in message, f"対象外である旨がメッセージに含まれるはず（実際:\n{message}）"
	assert "照合対象のMAOファイルが見つかりませんでした" in message

	print("test_build_case_result_returns_taishougai_when_no_masks: OK")


def test_sort_key_orders_by_circled_number():
	"""
	result.sort_keyが、丸数字プレフィックス（チェック①〜⑳）を番号順に並べるキーを返し、
	想定外の行（プレフィックス不一致）を末尾へ回すことを検証する（2026-09-02新規。
	②(check_fs_matrix)とのマージに伴う表示順固定化）。
	"""
	lines = ["チェック③（...）：OK", "想定外の行", "チェック①（...）：NG", "チェック②（...）：OK"]
	sorted_lines = sorted(lines, key=sort_key)
	assert sorted_lines == [
		"チェック①（...）：NG",
		"チェック②（...）：OK",
		"チェック③（...）：OK",
		"想定外の行",
	], f"番号順（想定外行は末尾）になっていない（実際:\n{sorted_lines}）"

	print("test_sort_key_orders_by_circled_number: OK")


def run_test():
	"""本ファイルの全テストを実行する。"""
	test_judgment_and_check_result()
	test_extract_matrix_can_id_records_excludes_bus_summary_and_gateway_route()
	test_find_matrix_header_location_detects_various_positions()
	test_extract_matrix_can_id_records_supports_3daa_cdc_style()
	test_select_matrix_sheet_returns_none_when_ambiguous()
	test_case_scan_basic()
	test_case_scan_duplicate_mao_is_structural_ng()
	test_case_scan_duplicate_matrix_is_structural_ng()
	test_case_scan_extracts_zip()
	test_build_case_result_aggregates_masks_and_structural_ng()
	test_format_case_result_message_shows_count_line()
	test_format_case_result_message_omits_heading_for_single_mask()
	test_derive_case_id_from_year_month_structure()
	test_derive_case_id_from_input_or_case_root()
	test_derive_case_id_fallback_for_branch_case_without_01_input()
	test_derive_case_id_returns_none_without_01_input()
	test_derive_case_id_flags_unexpected_format()
	test_find_can_table_header_row_detects_various_positions()
	test_detect_columns_uses_header_names_and_handles_missing_id_format()
	test_select_can_table_sheet_picks_unique_communication_sheet()
	test_select_can_table_sheet_returns_none_when_ambiguous()
	test_extract_a2l_can_id_records_reads_can_id_but_not_protocol_ids()
	test_extract_a2l_can_id_records_supports_cp932_fallback()
	test_case_scan_uses_a2l_when_mao_missing()
	test_case_scan_prefers_mao_when_both_present()
	test_case_scan_duplicate_a2l_is_structural_ng()
	test_case_scan_fi_icm_labels_do_not_collide()
	test_resolve_input_folder_path_from_input_itself()
	test_resolve_input_folder_path_from_case_root()
	test_resolve_input_folder_path_from_deeper_folder()
	test_resolve_input_folder_path_returns_none_without_01_input()
	test_build_case_result_returns_taishougai_when_no_masks()
	test_sort_key_orders_by_circled_number()
	print("すべてのテストに成功しました。")


if __name__ == "__main__":
	run_test()

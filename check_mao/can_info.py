"""CANマトリクス(.xlsx等)からCAN ID情報を抽出するモジュール。

【2026-09-01改修】チェック対象をCANテーブルからCANマトリクスへ変更した（船尾さんの指示）。
以降、①(check_mao)の判定に使うのは本ファイルの関数群（`extract_matrix_can_id_records`等）で
ある。2026-08-31改修時点のCANテーブル用実装（`extract_table_can_id_records`等）は、2026-09
追補で可読性改善のため`can_table_legacy.py`へ分離した（チェック対象の変遷の経緯を記録として
保持する目的。将来CANテーブルを再度対象に含める可能性があるため削除しない。詳細説明・改修履歴も
同ファイル側に記載）。

CANマトリクス（FCAN_ALL_THRU_Matrix等）:
	- CAN ID列の見出しは`CAN ID (HEX)`（3桁16進、"h"無し）。旧実装ではヘッダはR1〜R3・A列固定と
	  していたが、**機種によってヘッダ行・列位置にばらつきがある**ことを踏まえ、2026-09-01改修で
	  CANテーブルと同様に見出し文字列からの動的検出に変更した（`find_matrix_header_location`等を
	  参照）。
	- ECU別のTx/Rxノード列（"T(F2)","R(F5)"等）は存在するが、列位置・ECU名の記載行が機種
	  （3DM/3EL/3GZ/3QJ等）ごとに異なることを実データ調査で確認したため、本モジュールはこの列を
	  読み取らない（判定に必要なのは「CAN IDが存在するか」のみで、ECUごとの対応関係は判定ロジック上
	  不要。「4.5 判定ロジック」参照）。
	- シート末尾には、CAN ID列にバス名（"F1"〜"F12"）を置いたバス使用率(%)の集計行がある
	  （3DM・3GZの実データで確認済み）ほか、ゲートウェイ経路表記（"F1->F4,F5"等。"GW Direction"列）が
	  CAN ID列以外の列に混在することがある。バス使用率行・非16進値はいずれも読み飛ばす。
"""

import re

from normalize import normalize_can_id
from xlsx_reader import list_sheet_names, read_sheet

HEADER_TRANSMITTER_TEXT = "Transmitter"
HEADER_RECEIVERS_TEXT = "Receivers"


def is_valid_can_id_hex(hex_text):
	"""
	セルの値がCAN IDの16進表記として妥当かを判定する。

	CANテーブルには数値以外の記載（見出し・注記等）が混在する可能性があるため、
	16進数として解釈できる値だけをCAN IDとして扱う（この除外は非CAN-ID文字列の除外であり、
	「対象外CAN ID」の判別とは別物。対象外CAN IDの判別は今回のスコープ外）。
	"""
	if hex_text is None:
		return False
	cleaned = hex_text.strip()
	if not cleaned:
		return False
	if cleaned.lower().endswith("h"):
		cleaned = cleaned[:-1]
	try:
		int(cleaned, 16)
	except ValueError:
		return False
	return True


def _cell_text(row, column_index):
	"""1始まり列番号でセル値を取得する。文字列値は前後の空白を除去し、範囲外・Noneはそのままとする。"""
	if column_index is None:
		return None
	if column_index - 1 < 0 or column_index - 1 >= len(row):
		return None
	value = row[column_index - 1]
	if isinstance(value, str):
		return value.strip()
	return value


# =====================================================================================
# CANマトリクス関連（2026-09-01改修で①(check_mao)の判定対象に採用。以降の関数群が本体）
# =====================================================================================

# CANマトリクスのヘッダ探索範囲（実データで確認済みの旧位置はR1〜R3。機種によるヘッダ行の
# ばらつきに対応するため、CANテーブル側と同様に動的検出とし、余裕を持った行数まで探索する）
MATRIX_HEADER_SEARCH_ROW_LIMIT = 10

# CANマトリクスのCAN ID列見出し文字列（実データで確認済みの表記）。大文字小文字・前後の空白の
# 違いを吸収して比較する（`find_matrix_header_location`参照）。実データ(3DM
# `FCAN_ALL_THRU_Matrix_v1.0.0.xlsx`)では、セル内で"CAN ID"と"(HEX)"の間に改行が入る
# 表記（"CAN ID\n(HEX)"）で確認されており、これも同一の見出しとして一致させる必要がある
# （2026-09-01改修。フォーマットのばらつきの実例）。
MATRIX_CAN_ID_HEADER_TEXT = "CAN ID (HEX)"

# 3DAA(CDC)形式のCANマトリクス用フォールバック見出し（2026-09-01追補）。実データ202608_006_01の
# `Matrix_3DAA_FHEV_PADAS(CDC)_F_ALL_04.03.00.xlsx`では、ファイル名は"Matrix"を含むが内部の
# 「Matrix」シートの列構成が`CAN ID (HEX)`単独列ではなく、`ID`・`Transmitter`・`Receivers`という
# CANテーブルに近い列構成（列位置はB=ID/C=Transmitter/D=Receivers等、機種により異なる）になっている
# ことを確認した。`CAN ID (HEX)`見出しが見つからない場合、`Transmitter`の直後の列に`Receivers`が
# あるヘッダ行を探し、そのヘッダ行の`ID`という見出しを持つ列をCAN ID列として採用する。
MATRIX_ID_HEADER_TEXT = "ID"


# OOXML（Excel）が制御文字をシリアライズする際のエスケープ表記（例："_x000D_"=CR）。実データ
# （3AC_002の`BusMatrix_...xlsx`）で、この表記が復号されず文字列としてそのまま残っている
# ヘッダセル（"CAN ID_x000D_\n(HEX)"）を確認した（2026-09-03追補、HILS版で確認・本ファイルへ移植）。
# 空白正規化の前に除去する。
_EXCEL_CONTROL_CHAR_ESCAPE_PATTERN = re.compile(r"_x[0-9A-Fa-f]{4}_")


def _normalize_header_text(text):
	"""
	見出し文字列の空白（改行・タブ・連続スペースを含む）を単一の半角スペースへ正規化し、
	前後の空白を除去したうえで小文字化する。

	実データでは、セル内改行を含む表記（"CAN ID\\n(HEX)"）が確認されており、これを
	"CAN ID (HEX)"と同一視するために使う（`find_matrix_header_location`参照）。2026-09-03追補：
	OOXMLの制御文字エスケープ表記（"_x000D_"等）が復号されず文字列に残っているケースも確認され
	たため、空白正規化の前にこれを除去する。
	"""
	text = _EXCEL_CONTROL_CHAR_ESCAPE_PATTERN.sub(" ", text)
	return re.sub(r"\s+", " ", text.strip()).lower()

# ヘッダが検出できない場合の後方互換フォールバック（旧仕様。A列固定・R1〜R3がヘッダ、R4がデータ開始）
MATRIX_FALLBACK_CAN_ID_COLUMN = 1
MATRIX_FALLBACK_DATA_START_ROW = 4

# CANマトリクスのシート末尾には、CAN ID列にバス名("F1"〜"F12")を置いたバス使用率(%)の
# 集計行があり、これはCAN IDデータではない（3DM・3GZの実データで確認済み）。
# "F1"〜"F12"は16進数として解釈できてしまう（例:"F1"=0xF1）ため、is_valid_can_id_hex
# だけでは除外できず、明示的に除外する。3EL・3QJは"PFCAN1(F1)"等の記法で16進として
# 解釈できないため元々対象外（本パターンには一致しない）。
BUS_SUMMARY_TOKEN_PATTERN = re.compile(r"^F(1[0-2]|[1-9])$", re.IGNORECASE)


class MatrixCanIdRecord:
	"""CANマトリクスの1行(1 CAN ID)から抽出した情報を保持するレコード。"""

	def __init__(self, row_number, can_id_hex_text, can_id_decimal):
		self.row_number = row_number
		self.can_id_hex_text = can_id_hex_text
		self.can_id_decimal = can_id_decimal


def _find_transmitter_receivers_pair(rows):
	"""
	シート内で、見出し"Transmitter"の直後の列に見出し"Receivers"がある行を探す
	（3DAA(CDC)形式のフォールバック用。2026-09-01追補）。

	CANテーブル側の`find_can_table_header_row`はA列=Transmitter・B列=Receivers固定だが、
	3DAA(CDC)のMatrixシートでは列位置がB=ID/C=Transmitter/D=Receivers等、A列/B列に固定されない
	ため、全列を走査してTransmitter・Receiversが隣接する位置を探す。

	戻り値:
		(ヘッダ行番号, Transmitter列番号)のタプル（いずれも1始まり）。見つからない場合はNone。
	"""
	limit = min(MATRIX_HEADER_SEARCH_ROW_LIMIT, len(rows))
	transmitter_text = _normalize_header_text(HEADER_TRANSMITTER_TEXT)
	receivers_text = _normalize_header_text(HEADER_RECEIVERS_TEXT)
	for row_index in range(1, limit + 1):
		row = rows[row_index - 1]
		for column_index, value in enumerate(row, start=1):
			if not isinstance(value, str) or _normalize_header_text(value) != transmitter_text:
				continue
			if column_index >= len(row):
				continue
			next_value = row[column_index]
			if isinstance(next_value, str) and _normalize_header_text(next_value) == receivers_text:
				return row_index, column_index
	return None


def _find_id_column(header_row):
	"""ヘッダ行の中から、見出し文字列`ID`（空白・大文字小文字ゆれを無視）を持つ列を探す。"""
	target_text = _normalize_header_text(MATRIX_ID_HEADER_TEXT)
	for column_index, value in enumerate(header_row, start=1):
		if isinstance(value, str) and _normalize_header_text(value) == target_text:
			return column_index
	return None


def find_matrix_header_location(rows):
	"""
	CANマトリクスのヘッダ（`CAN ID (HEX)`列の見出し）を、シート先頭から
	`MATRIX_HEADER_SEARCH_ROW_LIMIT`行以内・全列から探す。

	旧実装はヘッダ行R1〜R3・A列固定の前提だったが、機種によりヘッダ行・列位置にばらつきが
	あることを踏まえ、CANテーブル側（`find_can_table_header_row`）と同様に見出し文字列からの
	動的検出へ変更した（2026-09-01改修）。

	`CAN ID (HEX)`という見出しが見つからない場合、3DAA(CDC)形式（`Transmitter`・`Receivers`
	ヘッダを持つCANテーブルに近い構造。実データ202608_006_01の
	`Matrix_3DAA_FHEV_PADAS(CDC)_F_ALL_04.03.00.xlsx`で確認）のヘッダをフォールバックとして
	試す。この場合、ヘッダ行の中から見出し`ID`を持つ列をCAN ID列として採用する
	（2026-09-01追補）。

	引数:
		rows: `xlsx_reader.read_sheet`が返す二次元リスト（1シート分）

	戻り値:
		(ヘッダ行番号, CAN ID列番号)のタプル（いずれも1始まり）。見つからない場合はNone。
	"""
	limit = min(MATRIX_HEADER_SEARCH_ROW_LIMIT, len(rows))
	target_text = _normalize_header_text(MATRIX_CAN_ID_HEADER_TEXT)
	for row_index in range(1, limit + 1):
		row = rows[row_index - 1]
		for column_index, value in enumerate(row, start=1):
			if isinstance(value, str) and _normalize_header_text(value) == target_text:
				return row_index, column_index

	transmitter_pair = _find_transmitter_receivers_pair(rows)
	if transmitter_pair is not None:
		header_row, _ = transmitter_pair
		id_column = _find_id_column(rows[header_row - 1])
		if id_column is not None:
			return header_row, id_column

	return None


def find_matrix_sheets(path):
	"""
	ブック内の全シートを走査し、`CAN ID (HEX)`ヘッダを持つシートを列挙する
	（CANテーブル側`find_can_table_sheets`と同じ考え方。2026-09-01新規）。

	引数:
		path: 走査するxlsx/xlsm/xlsファイルのパス

	戻り値:
		(シート名, ヘッダ行番号, CAN ID列番号)のタプルのリスト（ブック内のシート宣言順）。
		ブック自体、または個々のシートの読込に失敗した場合はスキップする（例外を上位へ伝播させない）。
	"""
	try:
		sheet_names = list_sheet_names(path)
	except Exception:
		return []

	found = []
	for sheet_name in sheet_names:
		try:
			rows = read_sheet(path, sheet_name)
		except Exception:
			continue
		header_location = find_matrix_header_location(rows)
		if header_location is not None:
			header_row, can_id_column = header_location
			found.append((sheet_name, header_row, can_id_column))
	return found


def looks_like_matrix(path):
	"""
	ファイルの内容（`CAN ID (HEX)`ヘッダを持つシートがちょうど1つあるか）からCANマトリクスらしさを
	判定する（CANテーブル側`looks_like_can_table`と同じ考え方。ファイル名の"Matrix"判定だけでは
	別構造の補助ファイルを誤って対象に含めてしまう可能性があるため、内容ベースの判別を用いる）。

	2026-09-03修正（HILS版で確認・本ファイルへ移植）：判定基準を「1件以上ヒット」から
	「ちょうど1件ヒット」へ変更した。実データ（3AC_001の使用禁止decoyファイル
	`Matrix_..._01.02.00 (1).xlsx`）で、1ファイル内に2シート（`Matrix`本体と、その列を絞り込んだ
	派生シート`受信DUMMYのみ削除IDリスト`）が3DAA(CDC)フォールバック検出に一致し、
	`select_matrix_sheet`は（自己矛盾のため）`None`を返すにもかかわらず、旧基準（`> 0`）では
	「候補あり」と誤判定されていた。これにより`case_scan._find_matrix_path`のファイル単位の
	候補選定で、本来1件しかない正しいCANマトリクスと合わせて候補2件に見え、「複数存在するため
	判定不能」の誤ったNGになっていた。`select_matrix_sheet`と同じ「一意に特定できるか」を基準に
	揃えることで、ファイル単体でも自己矛盾するdecoyファイルを候補から正しく除外する。
	"""
	return len(find_matrix_sheets(path)) == 1


def select_matrix_sheet(path):
	"""
	ブック内でCANマトリクス本体とみなすシートを1つ選ぶ（CANテーブル側`select_can_table_sheet`と
	同じ考え方）。

	候補（`CAN ID (HEX)`ヘッダを持つシート）が1つだけならそのまま採用する。候補が0個、または
	複数ある場合は、誤ったシートを選ばずNoneを返す（呼び出し側で「判定不可・要手動確認」として
	扱う）。CANテーブルの3DAA(CDC)のような複数候補の絞り込みキーワードは、CANマトリクスの
	実データでは確認されていないため、複数候補時の絞り込みは行わない。

	戻り値:
		(シート名, ヘッダ行番号, CAN ID列番号) または None
	"""
	candidates = find_matrix_sheets(path)
	if len(candidates) == 1:
		return candidates[0]
	return None


def extract_matrix_can_id_records(matrix_file_path, sheet_name=None, header_row=None, can_id_column=None):
	"""
	CANマトリクス(.xlsx等)を読み込み、`CAN ID (HEX)`列からCAN ID情報を抽出する。

	引数:
		matrix_file_path: CANマトリクスのパス
		sheet_name, header_row, can_id_column: `select_matrix_sheet`等で事前に確定した
			シート名・ヘッダ行番号・CAN ID列番号（`case_scan.py`がスキャン時に特定した値を
			そのまま渡す想定。呼び出し側の再計算・二重読込を避けるため）。3つとも未指定の場合は、
			この関数自身が`select_matrix_sheet`で特定する（`テスト\\run_sample.py`等、事前スキャンを
			経由しない単体確認からの呼び出しに対応するため）。
			ヘッダが検出できない場合（想定外の構成）は、後方互換のため旧固定レイアウト
			（A列・R4データ開始）へフォールバックする。

	戻り値:
		MatrixCanIdRecordのリスト
	"""
	if sheet_name is None and header_row is None and can_id_column is None:
		selected = select_matrix_sheet(matrix_file_path)
		if selected is not None:
			sheet_name, header_row, can_id_column = selected

	rows = read_sheet(matrix_file_path, sheet_name)

	if header_row is None or can_id_column is None:
		data_start_row = MATRIX_FALLBACK_DATA_START_ROW
		can_id_column = MATRIX_FALLBACK_CAN_ID_COLUMN
	else:
		data_start_row = header_row + 1

	records = []
	for row_index in range(data_start_row, len(rows) + 1):
		row = rows[row_index - 1]
		can_id_hex_text = _cell_text(row, can_id_column)
		if not is_valid_can_id_hex(can_id_hex_text):
			# ゲートウェイ経路表記等、CAN IDでない値は読み飛ばす
			continue
		if isinstance(can_id_hex_text, str) and BUS_SUMMARY_TOKEN_PATTERN.match(can_id_hex_text.strip()):
			# シート末尾のバス使用率集計行（バス名"F1"〜"F12"）はCAN IDではないため読み飛ばす
			continue

		can_id_decimal = normalize_can_id(can_id_hex_text)
		records.append(MatrixCanIdRecord(row_index, can_id_hex_text, can_id_decimal))

	return records


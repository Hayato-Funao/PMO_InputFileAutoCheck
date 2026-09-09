"""CANテーブル(.xlsm/.xlsx/.xls)からCAN ID情報を抽出するレガシーモジュール。

2026-08-31改修時点の実装（当時はCANテーブル単独でMAOとの整合を判定していた）。2026-09-01改修で
①(check_mao)の判定対象をCANマトリクスへ変更したため（`can_info.py`参照）、以降は本モジュールの
関数群は判定に使われていない。チェック対象の変遷の経緯を記録として保持する目的、および将来CAN
テーブルを再度対象に含める可能性があるため削除せず、`can_info.py`から本ファイルへ分離した
（2026-09追補・可読性改善のためのモジュール分割。ロジック自体の変更はない）。

CANテーブル（FCAN_ALL_THRU等）:
	- A列=Transmitter、B列=Receivers（いずれもECU名をカンマ区切りで格納）
	- C列=Signal Name、H列=ID（CAN ID,HEX）、I列=ID-Format（Standard/Extended）
	- ヘッダ行は多くの機種でR21だが、機種によって異なる（3DAA(CDC)ではR1）ため、
	  ヘッダ行・列位置はいずれも実データの見出し文字列から動的に検出する（後述、2026-08-31改修）。
	- .xlsm/.xlsxに加え、旧バイナリ形式(.xls)の機種にも対応する（read_sheet側でxlrdを使用）。

2026-08-31改修（本番想定36案件の実データ調査を踏まえた変更。当時はCANテーブル単独判定だった）:
	- CANマトリクスは対象外とし、CANテーブル単独でMAOとの整合を判定する（従前どおり）。
	  対象ファイルの判別は`case_scan.py`が担う。ファイル名に"Matrix"を含むものをCANマトリクス
	  として除外するが、それだけでは`_GW_INOUT`のような別構造の補助ファイル（Matrix名を含まない）
	  を除外できないケースが実データで確認された（202606_004）。
	- そのため、**ファイルの内容（Transmitter/Receiversという見出しを持つシートがあるか）でCAN
	  テーブルかどうかを判別する**関数（`looks_like_can_table`）を追加した。
	- さらに、3DAA(CDC)フォーマットのCANテーブルは(a)先頭シートが「変更履歴」でCANテーブル本体が
	  別シートにある、(b)1ブック内にTransmitter/Receiversヘッダを持つシートが複数（本体＋変更後・
	  削除リスト等の作業用シート）存在する、(c)シートによって列位置が異なる（ID列がH列でない場合や
	  ID-Format列が存在しない場合がある）ことが実データ調査で判明した。これに対応するため、
	  全シート走査によるヘッダ検出（`find_can_table_sheets`）、シート名によるCANテーブル本体の
	  選定（`select_can_table_sheet`）、見出し文字列による列の動的特定（`_detect_columns`）を追加した。
	  （実データ確認: 本番想定データ`202608_006_01`の`3DAA_FHEV_PADAS(CDC)_F_ALL_04.03.00.xlsx`）
"""

from normalize import normalize_can_id
from xlsx_reader import list_sheet_names, read_sheet
from can_info import HEADER_RECEIVERS_TEXT, HEADER_TRANSMITTER_TEXT, _cell_text, is_valid_can_id_hex

# ヘッダ行を探す範囲（ブックの先頭からこの行数以内。実データで確認済みの最大値はR21、
# 3DAA(CDC)はR1。将来の機種でこれより下にヘッダがあった場合は見直すこと）
HEADER_SEARCH_ROW_LIMIT = 30

HEADER_SIGNAL_NAME_TEXT = "Signal Name"
HEADER_CAN_ID_TEXT = "ID"
HEADER_ID_FORMAT_TEXT = "ID-Format"

# ヘッダの見出し文字列で列位置を特定できない場合の既定列（3M0U等、実データの大多数で一致する配置）
DEFAULT_TRANSMITTER_COLUMN = 1
DEFAULT_RECEIVERS_COLUMN = 2
DEFAULT_SIGNAL_NAME_COLUMN = 3
DEFAULT_CAN_ID_COLUMN = 8
DEFAULT_ID_FORMAT_COLUMN = 9

# ヘッダ行が全く検出できない場合の後方互換フォールバック（旧仕様。R21がヘッダ、R22がデータ開始）
FALLBACK_HEADER_ROW = 21

# 1ブックにTransmitter/Receiversヘッダを持つシートが複数ある場合、CANテーブル本体を選ぶための
# シート名キーワード（大小文字無視の部分一致）。実データでは本体シート名に"xFCAN1"等、通信バスを
# 示す名称が使われ、作業用シート（変更後・削除リスト等）は日本語名でこれらを含まない。
COMMUNICATION_SHEET_NAME_MARKERS = ("can", "communication")


class TableCanIdRecord:
	"""CANテーブルの1行(1シグナル)から抽出した情報を保持するレコード。"""

	def __init__(self, row_number, can_id_hex_text, can_id_decimal, transmitter, receivers, signal_name, id_format):
		self.row_number = row_number
		self.can_id_hex_text = can_id_hex_text
		self.can_id_decimal = can_id_decimal
		self.transmitter = transmitter
		self.receivers = receivers
		self.signal_name = signal_name
		self.id_format = id_format


def find_can_table_header_row(rows):
	"""
	CANテーブルのヘッダ行（A列=Transmitter・B列=Receivers）を、シートの先頭から
	`HEADER_SEARCH_ROW_LIMIT`行以内で探す。

	引数:
		rows: `xlsx_reader.read_sheet`が返す二次元リスト（1シート分）

	戻り値:
		ヘッダ行の行番号（1始まり）。見つからない場合はNone。
	"""
	limit = min(HEADER_SEARCH_ROW_LIMIT, len(rows))
	for row_index in range(1, limit + 1):
		row = rows[row_index - 1]
		if _cell_text(row, 1) == HEADER_TRANSMITTER_TEXT and _cell_text(row, 2) == HEADER_RECEIVERS_TEXT:
			return row_index
	return None


def _detect_columns(header_row):
	"""
	ヘッダ行の値から、各項目（Transmitter/Receivers/Signal Name/ID/ID-Format）の列位置を
	見出し文字列の完全一致で特定する。

	見出し名で見つからない必須項目（Transmitter・Receivers・Signal Name・ID）は、実データの
	大多数で一致する既定列へフォールバックする。ID-Format（任意項目。3DAA(CDC)の一部シートには
	存在しないことが実データで確認されている）は見つからなければNoneとし、別の列（Tx Method等）
	を誤って読まないようにする（2026-08-31改修。列位置固定のH列/I列読みでは、3DAA(CDC)のシートに
	よって列がずれる・ID-Format列が存在しないケースで誤読することが判明したため）。

	引数:
		header_row: ヘッダ行のセル値配列（`read_sheet`の行そのもの）

	戻り値:
		{"transmitter": 列番号, "receivers": 列番号, "signal_name": 列番号,
		 "can_id": 列番号, "id_format": 列番号またはNone}
	"""
	text_to_column = {}
	for column_index, value in enumerate(header_row, start=1):
		if isinstance(value, str) and value.strip():
			text_to_column.setdefault(value.strip(), column_index)

	return {
		"transmitter": text_to_column.get(HEADER_TRANSMITTER_TEXT, DEFAULT_TRANSMITTER_COLUMN),
		"receivers": text_to_column.get(HEADER_RECEIVERS_TEXT, DEFAULT_RECEIVERS_COLUMN),
		"signal_name": text_to_column.get(HEADER_SIGNAL_NAME_TEXT, DEFAULT_SIGNAL_NAME_COLUMN),
		"can_id": text_to_column.get(HEADER_CAN_ID_TEXT, DEFAULT_CAN_ID_COLUMN),
		"id_format": text_to_column.get(HEADER_ID_FORMAT_TEXT),  # 見つからなければNone（誤読防止）
	}


def find_can_table_sheets(path):
	"""
	ブック内の全シートを走査し、Transmitter/Receiversヘッダ行を持つシートを列挙する
	（2026-08-31新規。3DAA(CDC)のように先頭シートがCANテーブル本体でない機種に対応するため）。

	引数:
		path: 走査するxlsx/xlsm/xlsファイルのパス

	戻り値:
		(シート名, ヘッダ行番号)のタプルのリスト（ブック内のシート宣言順）。ブック自体、または
		個々のシートの読込に失敗した場合はスキップする（例外を上位へ伝播させない）。
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
		header_row = find_can_table_header_row(rows)
		if header_row is not None:
			found.append((sheet_name, header_row))
	return found


def looks_like_can_table(path):
	"""
	ファイルの内容（Transmitter/Receiversヘッダを持つシートの有無）からCANテーブルらしさを判定する
	（2026-08-31新規。ファイル名に"Matrix"を含まない補助ファイル`_GW_INOUT`等を、ファイル名だけでは
	除外できないため、内容ベースの判別を追加した）。
	"""
	return len(find_can_table_sheets(path)) > 0


def select_can_table_sheet(path):
	"""
	ブック内でCANテーブル本体とみなすシートを1つ選ぶ（2026-08-31新規）。

	候補（Transmitter/Receiversヘッダを持つシート）が1つだけならそのまま採用する（実データの
	大多数はこのケース）。複数ある場合（3DAA(CDC)のように、本体シートに加えて変更後・削除リスト等の
	作業用シートも同じヘッダを持つ機種）は、シート名に"CAN"または"Communication"（大小文字無視）を
	含むものへ絞り込み、一意に定まればそれを本体として採用する。候補が0個、または絞り込んでも
	複数残る場合は、誤ったシートを選ばずNoneを返す（呼び出し側で「判定不可・要手動確認」として扱う）。

	戻り値:
		(シート名, ヘッダ行番号) または None
	"""
	candidates = find_can_table_sheets(path)
	if len(candidates) == 0:
		return None
	if len(candidates) == 1:
		return candidates[0]

	narrowed = [
		candidate
		for candidate in candidates
		if any(marker in candidate[0].lower() for marker in COMMUNICATION_SHEET_NAME_MARKERS)
	]
	if len(narrowed) == 1:
		return narrowed[0]
	return None


def extract_table_can_id_records(table_file_path, sheet_name=None, header_row=None):
	"""
	CANテーブル(.xlsm/.xlsx/.xls)を読み込み、シグナル単位のCAN ID情報を抽出する。

	引数:
		table_file_path: CANテーブルのパス
		sheet_name, header_row: `select_can_table_sheet`等で事前に確定したシート名・ヘッダ行番号
			（`case_scan.py`がスキャン時に特定した値をそのまま渡す想定。呼び出し側の再計算・
			二重読込を避けるため）。両方None（未指定）の場合は、この関数自身が
			`select_can_table_sheet`で特定する（`テスト\\run_sample.py`等、事前スキャンを経由しない
			単体確認からの呼び出しに対応するため）。
			ヘッダが検出できない場合（想定外の構成）は、後方互換のため先頭シート・固定レイアウト
			（R21ヘッダ・R22データ開始・列1/2/3/8/9）へフォールバックする。

	戻り値:
		TableCanIdRecordのリスト
	"""
	if sheet_name is None and header_row is None:
		selected = select_can_table_sheet(table_file_path)
		if selected is not None:
			sheet_name, header_row = selected

	if header_row is None:
		rows = read_sheet(table_file_path, sheet_name)
		data_start_row = FALLBACK_HEADER_ROW + 1
		columns = {
			"transmitter": DEFAULT_TRANSMITTER_COLUMN,
			"receivers": DEFAULT_RECEIVERS_COLUMN,
			"signal_name": DEFAULT_SIGNAL_NAME_COLUMN,
			"can_id": DEFAULT_CAN_ID_COLUMN,
			"id_format": DEFAULT_ID_FORMAT_COLUMN,
		}
	else:
		rows = read_sheet(table_file_path, sheet_name)
		header_values = rows[header_row - 1] if header_row - 1 < len(rows) else []
		columns = _detect_columns(header_values)
		data_start_row = header_row + 1

	records = []
	for row_index in range(data_start_row, len(rows) + 1):
		row = rows[row_index - 1]
		can_id_hex_text = _cell_text(row, columns["can_id"])
		if not is_valid_can_id_hex(can_id_hex_text):
			continue

		can_id_decimal = normalize_can_id(can_id_hex_text)
		records.append(
			TableCanIdRecord(
				row_index,
				can_id_hex_text,
				can_id_decimal,
				_cell_text(row, columns["transmitter"]),
				_cell_text(row, columns["receivers"]),
				_cell_text(row, columns["signal_name"]),
				_cell_text(row, columns["id_format"]),
			)
		)

	return records

"""MAOファイル(.mao)からCAN ID情報を抽出するモジュール。

MAOはCCPラベル定義をShift-JIS(cp932)テキストで持つファイルで、各行が
1つのCCPラベルレコードに対応する。CAN IDは「説明文」に相当する部分に
"CAN ID=074h"という形式で書かれている（実データ調査で確認済み、197行）。

区切り文字（表示上"?"に見える1バイト文字）の実体はレコードごとに解釈が
必要だが、CAN ID抽出自体は区切り文字に依存しない行単位の正規表現で行う
（区切り文字の解釈を誤ってもCAN ID抽出には影響しない設計）。

実データで確認済みの表記ゆれ:
	- 基本形: "CAN ID=074h"
	- "h"欠落: "CAN ID=14E"
	- 前置語混在: "IMA-CAN ID=338h" / "スターター制御用CAN ID=387h" / "1回目F-CAN ID=195h"
	- 拡張ID(8桁): "CAN ID=0CD9AA4Dh"
	- 配列ラベル: "CNTCAN1B8C[0]"のように末尾に添字が付き、同一CAN IDを複数行が参照する
"""

import re

from normalize import normalize_can_id

# ラベル名は行頭から、英数字・アンダースコア・角括弧が続く範囲とする
LABEL_NAME_PATTERN = re.compile(r"^[A-Za-z0-9_\[\]]+")

# "CAN ID=074h"のような表記を、前置語の有無を問わず抜き出す正規表現。
# "h"は付かない実データ（"CAN ID=14E"）もあるため任意(?)とする。
CAN_ID_PATTERN = re.compile(r"CAN\s*ID\s*=\s*([0-9A-Fa-f]{1,8})h?", re.IGNORECASE)

# ラベル名末尾の送信/受信サフィックス判定（配列添字"[0]"等が付く場合も許容）
DIRECTION_SUFFIX_PATTERN = re.compile(r"(TX|RX)(\[\d+\])?$")


class MaoCanIdRecord:
	"""MAOの1行から抽出したCAN ID情報を保持するレコード。"""

	def __init__(self, line_number, label_name, can_id_hex_text, can_id_decimal, direction):
		self.line_number = line_number
		self.label_name = label_name
		self.can_id_hex_text = can_id_hex_text
		self.can_id_decimal = can_id_decimal
		self.direction = direction  # "TX" / "RX" / None(判別不能)


def detect_direction(label_name):
	"""ラベル名末尾のTX/RX表記から送信・受信方向を判定する。判別できない場合はNoneを返す。"""
	matched = DIRECTION_SUFFIX_PATTERN.search(label_name)
	if matched is None:
		return None
	return matched.group(1)


def extract_mao_can_id_records(mao_file_path):
	"""
	MAOファイルを読み込み、CAN IDが記載された全行からレコードを抽出する。

	引数:
		mao_file_path: 読み込むMAOファイル(.mao)のパス

	戻り値:
		MaoCanIdRecordのリスト（出現順）。1つのCAN IDに複数ラベルが対応する場合、
		レコードは重複排除せずすべて保持する（重複排除はbuild_unique_can_id_mapで行う）。
	"""
	records = []

	# MAOはShift-JIS(cp932)。想定外の文字が混在してもファイル全体の読み取りを
	# 止めないよう、置換文字での読み進めを許容する。
	with open(mao_file_path, encoding="cp932", errors="replace") as mao_file:
		for line_number, line in enumerate(mao_file, start=1):
			matched = CAN_ID_PATTERN.search(line)
			if matched is None:
				continue

			label_matched = LABEL_NAME_PATTERN.match(line)
			label_name = label_matched.group(0) if label_matched is not None else ""

			can_id_hex_text = matched.group(1)
			can_id_decimal = normalize_can_id(can_id_hex_text)
			direction = detect_direction(label_name)

			records.append(MaoCanIdRecord(line_number, label_name, can_id_hex_text, can_id_decimal, direction))

	return records


def build_unique_can_id_map(records):
	"""
	抽出済みレコードから、正規化CAN ID(10進)をキーに紐づくレコード一覧をまとめる。

	戻り値:
		{正規化CAN ID(10進): [対応するMaoCanIdRecord, ...]} の辞書
	"""
	can_id_to_records = {}
	for record in records:
		if record.can_id_decimal is None:
			continue
		can_id_to_records.setdefault(record.can_id_decimal, []).append(record)
	return can_id_to_records

"""CAN IDの表記ゆれ（"h"の有無・桁数・大文字小文字・前置語）を吸収し、
3ソース（MAO・CANマトリクス・CANテーブル）を同じ基準で突合できるようにする正規化処理。

正規化の方針（実データ調査で確定した事実に基づく）:
	- 表記例: "074h" / "074" / "14E"(h欠落) / "IMA-CAN ID=338h"(前置語混在) / "0CD9AA4Dh"(拡張8桁)
	- いずれも16進数として解釈し、10進整数に変換したうえで比較する
	- 前ゼロ・大文字小文字の違いはint()変換時に自然に吸収される

注意（今回のスコープ外）:
	- 拡張ID(29bit)はDBC側でbit31フラグが立った表記になることがあるが、今回はMAO・
	  CANマトリクス・CANテーブルのみを扱うため、bit31マスクは行わない。DBCを扱う場合は
	  別途マスク処理を追加すること。
	- 【2026-09追補】A2Lファイル（MAO代替取得用。`mao_can_id.extract_a2l_can_id_records`）は
	  bit31マスクが不要であることを確認済み。マスクが必要になるのはA2Lの`/begin CAN`〜
	  `/end CAN`ブロック（計測/診断ツール通信ID、bit31フラグ付き表記）を読む場合であり、
	  本ツールが読むMEASUREMENT／CHARACTERISTIC説明文中の"CAN ID=074h"（車両CAN ID）は
	  bit31フラグの表記ではないため、そのまま`normalize_can_id`で正規化してよい。
"""


def normalize_can_id(hex_text):
	"""
	16進表記のCAN ID文字列を10進整数に正規化する。

	引数:
		hex_text: 16進表記のCAN ID（例:"074h", "074", "14E", "0CD9AA4D"）。
		          末尾の"h"・前後の空白は許容する。

	戻り値:
		10進整数のCAN ID。変換できない場合はNone。
	"""
	if hex_text is None:
		return None

	# 末尾の"h"（大文字小文字問わず）と前後の空白を取り除く
	cleaned = hex_text.strip()
	if cleaned.lower().endswith("h"):
		cleaned = cleaned[:-1]
	cleaned = cleaned.strip()

	if not cleaned:
		return None

	try:
		return int(cleaned, 16)
	except ValueError:
		return None


def format_can_id_hex(decimal_value, digit_count=3):
	"""10進整数のCAN IDを、桁数をそろえた大文字16進表記（末尾"h"付き）に変換する。"""
	if decimal_value is None:
		return ""
	hex_digits = format(decimal_value, "X")
	if len(hex_digits) < digit_count:
		hex_digits = hex_digits.rjust(digit_count, "0")
	return hex_digits + "h"

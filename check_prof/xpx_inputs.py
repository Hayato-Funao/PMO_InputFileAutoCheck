"""
XPX INPUT不備チェック用の入力ファイルパーサ群。

対応ファイル:
  - Intel HEX  … 書き込み範囲 / トレーラのメタ情報 / 任意アドレスのバイト列
  - .epd       … 書き込み許可ROMセクタ / 診断CAN ID / CPUTYPE
  - .a2l       … コードフラッシュ領域定義 / 診断CAN ID

判定ロジックは持たない(checks.py 側が担当)。
"""

import re
from pathlib import Path

# 1レコードが跨げるアドレス境界(Intel HEXのアドレスフィールドは16bit)
ADDRESS_PAGE_SIZE = 0x10000

# DOSのEOFマーカー。フラッシュツール生成HEXではこの後にメタ情報が続く
EOF_MARKER = b"\x1a"

# Intel HEXレコードタイプ
RECORD_DATA = 0x00
RECORD_EOF = 0x01
RECORD_EXT_SEGMENT_ADDR = 0x02
RECORD_START_SEGMENT_ADDR = 0x03
RECORD_EXT_LINEAR_ADDR = 0x04
RECORD_START_LINEAR_ADDR = 0x05

# A2Lの拡張CAN IDはbit31を立てて表現される(実IDは29bit)
CAN_EXT_ID_MASK = 0x1FFFFFFF


class HexParseError(Exception):
	"""Intel HEXファイルを解析できない場合の例外

	「ファイルが読めない」ことを示す。ソフトの欠陥(NG)とは区別して扱うこと。
	"""
	pass


class EpdParseError(Exception):
	"""epdファイルを解析できない場合の例外"""
	pass


class A2lParseError(Exception):
	"""A2Lファイルを解析できない場合の例外"""
	pass


# =====================================================================
# Intel HEX
# =====================================================================

def _decode_hex_line(raw_line: bytes, line_no: int) -> str:
	"""HEXレコード行をASCIIとしてデコードする"""
	try:
		return raw_line.decode("ascii")
	except UnicodeDecodeError:
		raise HexParseError(f"{line_no}行目: ASCII以外の文字が含まれています")


def _iter_hex_lines(stream):
	"""バイナリストリームからHEXレコード行だけを(行番号, 文字列)で取り出す。

	フラッシュツールが生成するHEXは、EOFレコードの後にDOSのEOFマーカー(0x1A)で
	始まるメタ情報(COM:/DBF:/HIS: 等、Shift_JISの日本語を含む場合がある)を
	付与する。0x1A以降はHEXレコードではないため、そこで打ち切る。
	"""
	for line_no, raw_line in enumerate(stream, start=1):
		raw_line = raw_line.strip(b"\r\n")

		marker = raw_line.find(EOF_MARKER)
		if marker >= 0:
			raw_line = raw_line[:marker].strip()
			if raw_line:
				yield line_no, _decode_hex_line(raw_line, line_no)
			return

		raw_line = raw_line.strip()
		if not raw_line:
			continue
		yield line_no, _decode_hex_line(raw_line, line_no)


def _parse_hex_record(line: str, line_no: int) -> tuple:
	"""1行のIntel HEXレコードを(byte_count, addr, record_type, data)に分解する。

	検証はバイト数整合とチェックサムのみ。これは「次のフィールド位置を
	決定できない」「壊れたデータで誤ったアドレス範囲を算出してしまう」の
	2点を防ぐために不可欠なため。
	EOFレコードの有無やレコード種別の網羅性は検証しない(ツールチェーンの
	出力形式に関する仮定を持ち込まないため)。
	"""
	if not line.startswith(":"):
		raise HexParseError(f"{line_no}行目: ':'で始まらない行です")

	body = line[1:]
	if len(body) % 2 != 0:
		raise HexParseError(f"{line_no}行目: レコード長が奇数です")
	if len(body) < 10:  # byte_count(1)+addr(2)+type(1)+checksum(1) = 5バイト
		raise HexParseError(f"{line_no}行目: レコードが短すぎます")

	try:
		raw = bytes.fromhex(body)
	except ValueError:
		raise HexParseError(f"{line_no}行目: 16進数以外の文字が含まれています")

	byte_count = raw[0]
	if len(raw) != 5 + byte_count:
		raise HexParseError(
			f"{line_no}行目: バイト数フィールド({byte_count})と"
			f"実データ長({len(raw) - 5})が一致しません"
		)
	if sum(raw) & 0xFF != 0:
		raise HexParseError(f"{line_no}行目: チェックサム不一致")

	addr = (raw[1] << 8) | raw[2]
	return byte_count, addr, raw[3], raw[4:4 + byte_count]


def iter_hex_data(hex_path: str):
	"""HEXのデータレコードを(絶対アドレス, データ)で順に返す。

	レコード種別 04(拡張リニアアドレス) と 02(拡張セグメントアドレス) で
	上位アドレスを更新し、03/05(開始アドレス指定)は書き込み範囲に影響しない
	ため読み飛ばす。
	"""
	base = 0
	with open(hex_path, "rb") as f:
		for line_no, line in _iter_hex_lines(f):
			byte_count, addr, record_type, data = _parse_hex_record(line, line_no)

			if record_type == RECORD_DATA:
				if addr + byte_count > ADDRESS_PAGE_SIZE:
					# 64KB境界を跨ぐレコードはツール間で解釈が異なるため扱わない
					raise HexParseError(
						f"{line_no}行目: データレコードが64KB境界を跨いでいます"
						f"(addr=0x{addr:04X}, len={byte_count})"
					)
				yield base + addr, data
			elif record_type == RECORD_EXT_LINEAR_ADDR:
				if byte_count != 2:
					raise HexParseError(f"{line_no}行目: 04レコードの長さが不正です")
				base = ((data[0] << 8) | data[1]) << 16
			elif record_type == RECORD_EXT_SEGMENT_ADDR:
				if byte_count != 2:
					raise HexParseError(f"{line_no}行目: 02レコードの長さが不正です")
				base = ((data[0] << 8) | data[1]) << 4
			elif record_type == RECORD_EOF:
				return
			# 03 / 05 / その他は書き込み範囲に影響しないため無視する


def load_hex_write_ranges(hex_path: str) -> list:
	"""HEXが実際に書き込むアドレス範囲を[start, end)のリストで返す(マージ済み)"""
	ranges = []
	cur_start = None
	cur_len = 0

	for addr, data in iter_hex_data(hex_path):
		if cur_start is None:
			cur_start, cur_len = addr, len(data)
		elif addr == cur_start + cur_len:
			cur_len += len(data)
		else:
			ranges.append((cur_start, cur_start + cur_len))
			cur_start, cur_len = addr, len(data)

	if cur_start is not None:
		ranges.append((cur_start, cur_start + cur_len))
	if not ranges:
		raise HexParseError(f"データレコードが1件もありません: {hex_path}")

	return merge_ranges(ranges)


def read_hex_bytes(hex_path: str, address: int, length: int):
	"""HEXの指定アドレスからlengthバイトを取り出す。

	boot Ver.確認ツール(Excel)と異なり、レコード先頭に一致しないアドレスでも
	読み出せる(例: 0xA00048BC は16byte境界に無いためExcel版では取得不可)。
	全バイトが揃わない場合は None を返す。
	"""
	buf = {}
	end = address + length
	for addr, data in iter_hex_data(hex_path):
		if addr >= end or addr + len(data) <= address:
			continue
		for i, byte in enumerate(data):
			if address <= addr + i < end:
				buf[addr + i] = byte
	if len(buf) != length:
		return None
	return bytes(buf[address + i] for i in range(length))


def load_hex_trailer(hex_path: str) -> dict:
	"""HEXのトレーラ(0x1A以降のメタ情報)を辞書で返す。

	フラッシュツールが付与する COM:/DBF:/MAF:/2ND:/PVR:/DVR: 等のフィールドを
	抽出する。DBF/MAFはビルドID(A2L等との対応確認に使える)。
	トレーラが無い場合は空辞書。
	"""
	data = Path(hex_path).read_bytes()
	pos = data.find(EOF_MARKER)
	if pos < 0:
		return {}
	text = data[pos + 1:].decode("shift_jis", errors="replace")
	fields = {}
	for key, value in re.findall(r"([A-Z0-9]{3}):((?:(?![A-Z0-9]{3}:).)*)", text):
		fields[key] = value.strip()
	return fields


# =====================================================================
# .epd (フラッシュツールのプロファイル、Shift_JIS XML属性)
# =====================================================================

def _read_epd(epd_path: str) -> str:
	return Path(epd_path).read_text(encoding="shift_jis", errors="replace")


def load_epd_rom_ranges(epd_path: str) -> list:
	"""epdのROMxx_ADDRESS/ROMxx_SIZEから書き込み許可範囲を返す(マージ済み)。

	- ROM番号は連番でなくても、桁数(ROM00 / ROM100)が混在していても読める。
	- SIZE=0のエントリはプレースホルダとみなして無視する。
	"""
	content = _read_epd(epd_path)
	addresses = {int(n): int(v) for n, v in re.findall(r'ROM(\d+)_ADDRESS="(\d+)"', content)}
	sizes = {int(n): int(v) for n, v in re.findall(r'ROM(\d+)_SIZE="(\d+)"', content)}

	if not addresses:
		raise EpdParseError(f"epdにROMxx_ADDRESS定義が見つかりません: {epd_path}")

	only_address = sorted(set(addresses) - set(sizes))
	only_size = sorted(set(sizes) - set(addresses))
	if only_address or only_size:
		raise EpdParseError(
			f"epdのROM定義が不整合です(ADDRESSのみ: {only_address}, SIZEのみ: {only_size})"
		)

	ranges = [(addresses[i], addresses[i] + sizes[i])
	          for i in sorted(addresses) if sizes[i] > 0]
	if not ranges:
		raise EpdParseError(f"epdに有効な(SIZE>0の)ROMセクタがありません: {epd_path}")
	return merge_ranges(ranges)


def load_epd_can_ids(epd_path: str) -> set:
	"""epdのCANIDR/CANIDS(10進表記)を29bitの診断CAN IDの集合として返す"""
	content = _read_epd(epd_path)
	ids = set()
	for key in ("CANIDR", "CANIDS"):
		m = re.search(rf'{key}="(\d+)"', content)
		if m:
			ids.add(int(m.group(1)) & CAN_EXT_ID_MASK)
	if not ids:
		raise EpdParseError(f"epdにCANIDR/CANIDS定義が見つかりません: {epd_path}")
	return ids


def load_epd_cpu_type(epd_path: str):
	"""epdのCPUTYPE属性を返す(無ければNone)"""
	m = re.search(r'CPUTYPE="([^"]*)"', _read_epd(epd_path))
	return m.group(1) if m else None


# =====================================================================
# .a2l (ASAP2)
# =====================================================================

def load_a2l_code_segments(a2l_path: str) -> list:
	"""A2Lの`CODE FLASH INTERN`宣言からコード領域を[start, end)で返す(マージ済み)。

	書式: <名前> "<説明>" CODE FLASH INTERN <開始アドレス> <サイズ> ...
	"""
	pattern = re.compile(rb"CODE\s+FLASH\s+INTERN\s+0x([0-9A-Fa-f]+)\s+0x([0-9A-Fa-f]+)")
	segments = []
	with open(a2l_path, "rb") as f:
		for line in f:
			m = pattern.search(line)
			if m:
				start = int(m.group(1), 16)
				segments.append((start, start + int(m.group(2), 16)))
	if not segments:
		raise A2lParseError(f"A2Lに'CODE FLASH INTERN'宣言が見つかりません: {a2l_path}")
	return merge_ranges(segments)


def load_a2l_module(a2l_path: str):
	"""A2Lの`/begin MODULE <名前>`のモジュール名を返す(無ければNone)。

	実測ではデバイス型番が入る(TC277 / TC387 / TC275 / TC1782 / RH850C1M 等)。
	A2Lのファイル名(ビルドID)はファイル内容に現れないため、内容から
	デバイス種別を判定できる唯一の手掛かり。
	"""
	pattern = re.compile(rb"/begin\s+MODULE\s+(\S+)")
	with open(a2l_path, "rb") as f:
		for line in f:
			m = pattern.search(line)
			if m:
				return m.group(1).decode("shift_jis", errors="replace")
	return None


def normalize_device(name: str):
	"""デバイス型番表記を比較可能な形に正規化する。

	'Infineon TC387_暫定' → 'TC387' / 'TC277' → 'TC277'
	ベンダ名を除去し、先頭の英数字の連なりだけを取り出して大文字化する。
	"""
	if not name:
		return None
	text = re.sub(r"(?i)\b(infineon|renesas|nxp|st)\b", " ", name).strip()
	m = re.match(r"([0-9A-Za-z]+)", text)
	return m.group(1).upper() if m else None


def load_a2l_can_ids(a2l_path: str) -> set:
	"""A2LのCANセクション内 /begin ADDRESS ～ /end ADDRESS のCAN IDを返す。

	拡張IDはbit31を立てて記述される(0x98DAF110 → 実ID 0x18DAF110)ため
	29bitでマスクして返す。
	"""
	ids = set()
	in_can = False
	in_address = False
	with open(a2l_path, "rb") as f:
		for raw in f:
			line = raw.split(b"/*")[0]  # 行末コメントを除去
			if b"/begin CAN" in raw:
				in_can = True
			elif b"/end CAN" in raw:
				in_can = False
			if not in_can:
				continue
			if b"/begin ADDRESS" in raw:
				in_address = True
				continue
			if b"/end ADDRESS" in raw:
				in_address = False
				continue
			if in_address:
				m = re.search(rb"0x([0-9A-Fa-f]+)", line)
				if m:
					ids.add(int(m.group(1), 16) & CAN_EXT_ID_MASK)
	if not ids:
		raise A2lParseError(f"A2LにCANセクションのADDRESS定義が見つかりません: {a2l_path}")
	return ids


# =====================================================================
# 範囲演算のヘルパー
# =====================================================================

def merge_ranges(ranges: list) -> list:
	"""重複・隣接する範囲をマージする"""
	merged = []
	for start, end in sorted(ranges):
		if merged and start <= merged[-1][1]:
			merged[-1] = (merged[-1][0], max(merged[-1][1], end))
		else:
			merged.append((start, end))
	return merged


def subtract_ranges(target: list, covering: list) -> list:
	"""targetのうちcoveringに含まれない部分を返す(集合差)"""
	result = []
	for t_start, t_end in target:
		covered = t_start
		for c_start, c_end in covering:
			if c_end <= covered or c_start >= t_end:
				continue
			if c_start > covered:
				result.append((covered, min(c_start, t_end)))
			covered = max(covered, c_end)
		if covered < t_end:
			result.append((covered, t_end))
	return result


def intersect_size(a: list, b: list) -> int:
	"""2つの範囲リストの交差サイズ(バイト)"""
	return sum(max(0, min(a_end, b_end) - max(a_start, b_start))
	           for a_start, a_end in a for b_start, b_end in b)


def format_ranges(ranges: list) -> str:
	"""範囲リストを人が読める1行文字列にする"""
	if not ranges:
		return "(なし)"
	return ", ".join(f"0x{s:08X}-0x{e - 1:08X}({format_size(e - s)})" for s, e in ranges)


def format_size(size: int) -> str:
	"""バイト数をKB/MB付きの読みやすい文字列にする"""
	if size >= 1024 * 1024 and size % (1024 * 1024) == 0:
		return f"{size // (1024 * 1024)}MB"
	if size >= 1024 and size % 1024 == 0:
		return f"{size // 1024}KB"
	return f"{size}B"


# =====================================================================
# 診断CAN IDからのECUアドレス解読
# =====================================================================

# ISO 15765-2 の物理アドレス指定: 0x18DA<宛先><送信元>。テスタ側は 0xF1
DIAG_ID_PREFIX = 0x18DA0000
DIAG_TESTER_ADDRESS = 0xF1

# ECUアドレスと種別の対応(資料スライド3の記載より)
ECU_TYPE_BY_ADDRESS = {
	0x10: "FI",
	0x0E: "FI",
	0x07: "ICM",
}


def ecu_addresses_from_can_ids(can_ids) -> set:
	"""診断CAN IDの集合からECUアドレスを取り出す。

	0x18DAxxyy の xx=宛先 / yy=送信元。テスタ(0xF1)でない側がECUアドレス。
	診断IDの形式でないものは無視する。
	"""
	found = set()
	for can_id in can_ids:
		if (can_id & 0xFFFF0000) != DIAG_ID_PREFIX:
			continue
		target = (can_id >> 8) & 0xFF
		source = can_id & 0xFF
		if source == DIAG_TESTER_ADDRESS:
			found.add(target)      # テスタ → ECU(要求)
		elif target == DIAG_TESTER_ADDRESS:
			found.add(source)      # ECU → テスタ(応答)
	return found


def format_ecu_addresses(can_ids) -> str:
	"""ECUアドレスを '0x10 (FI)' 形式の読みやすい文字列にする"""
	addresses = ecu_addresses_from_can_ids(can_ids)
	if not addresses:
		return "(診断IDの形式ではないため不明)"
	return ", ".join(
		f"0x{a:02X}" + (f" ({ECU_TYPE_BY_ADDRESS[a]})" if a in ECU_TYPE_BY_ADDRESS else " (種別未登録)")
		for a in sorted(addresses))

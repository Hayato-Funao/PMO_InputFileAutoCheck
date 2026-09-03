"""
XPX INPUT不備チェック ③(HEX / .epd / .a2l の整合性確認)。

既存2ツールの確認内容を取り込んでいる:
  - AutoFlashHex(MATLAB) CombinationCheck.AddressCheck
      → check_a2l_epd_start()   (比較先を .cnf から .epd に置き換え、
                                 文字列比較ではなく数値比較、かつfail-closed)
  - AutoFlashHex(MATLAB) CombinationCheck.CANIDCheck
      → check_a2l_epd_canid()
  - boot Ver.自動確認ツール(Excel VBA) CompareBootVer
      → check_boot_version()    (レコード先頭に限らず任意アドレスを読める)

各チェックはCheckResultを返す。STATUS_NGのみが「ソフト側の不備」を意味し、
STATUS_ERRORは「入力が読めない/情報が足りない」を意味する(区別して扱う)。
"""

import re
from pathlib import Path

from xpx_inputs import (
	A2lParseError,
	EpdParseError,
	HexParseError,
	format_ranges,
	format_size,
	intersect_size,
	load_a2l_can_ids,
	load_a2l_code_segments,
	load_a2l_module,
	load_epd_can_ids,
	load_epd_cpu_type,
	load_epd_rom_ranges,
	load_hex_trailer,
	load_hex_write_ranges,
	format_ecu_addresses,
	normalize_device,
	read_hex_bytes,
	subtract_ranges,
)

STATUS_OK = "OK"
STATUS_NG = "NG"
STATUS_WARN = "警告"
STATUS_ERROR = "確認不能"   # 入力が読めない/参照情報が無い
STATUS_SKIP = "対象外"      # 任意入力が与えられていない

# boot Ver.のアドレス表。boot Ver.自動確認ツール(Excel)の
# 「boot_Verアドレスリスト」シートから移設したもの。
# ※このツールの所有者が維持している一次データであり、機種追加時は同期が必要。
BOOT_VER_ADDRESSES = {
	"BZ1_CS非対応機種": 0xA0001DB0,
	"BZ1_CS対応機種": 0xA001F080,
	"HP1-ICM": 0xA00048BC,
	"FI_CX1V_CXKV": 0xA0000780,
	"ICM_HP1": 0xA00048BC,
	"MOT_MG4Y": 0xA00013B0,
}

# boot Ver.として読み出すバイト数(Excel版はレコード1件=16byteを取得していた)
BOOT_VER_LENGTH = 16


class CheckResult:
	"""1チェックの結果"""

	def __init__(self, key: str, title: str, status: str, message: str, detail=None):
		self.key = key
		self.title = title
		self.status = status
		self.message = message
		self.detail = detail or []

	@property
	def is_ng(self) -> bool:
		return self.status == STATUS_NG

	@property
	def is_error(self) -> bool:
		return self.status == STATUS_ERROR

	def __repr__(self):
		return f"<{self.key} {self.status}: {self.message}>"


# =====================================================================
# ① A2L ↔ epd 先頭アドレス一致
# =====================================================================

def check_a2l_epd_start(a2l_path: str, epd_path: str) -> CheckResult:
	"""A2Lのコード領域先頭アドレスと、epdの許可範囲先頭が一致するか。

	A2Lが宣言する「アプリ領域の開始」と、フラッシュツールが書き込みを許可する
	範囲の開始は同一でなければならない。不一致はソフトとプロファイルの
	組み合わせ違い(例: BZ1のソフトにTC387用プロファイル)を意味する。
	"""
	key, title = "①", "A2L↔epd 先頭アドレス一致"
	try:
		segments = load_a2l_code_segments(a2l_path)
		allowed = load_epd_rom_ranges(epd_path)
	except (A2lParseError, EpdParseError) as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	a2l_start = segments[0][0]
	epd_start = allowed[0][0]
	detail = [f"A2Lコード領域先頭 : 0x{a2l_start:08X}",
	          f"epd許可範囲先頭   : 0x{epd_start:08X}"]
	if a2l_start == epd_start:
		return CheckResult(key, title, STATUS_OK,
		                   f"一致 (0x{a2l_start:08X})", detail)
	return CheckResult(key, title, STATUS_NG,
	                   f"不一致 (A2L=0x{a2l_start:08X} / epd=0x{epd_start:08X})"
	                   " → ソフトとプロファイルの組み合わせ違いの可能性", detail)


# =====================================================================
# 参考1 A2L ↔ epd 診断CAN ID一致(仕様外・任意)
# =====================================================================

def check_a2l_epd_canid(a2l_path: str, epd_path: str) -> CheckResult:
	"""A2LのCANセクションのIDと、epdのCANIDR/CANIDSが一致するか。

	送受信の対応付けを取り違えないよう、集合として比較する。
	"""
	key, title = "参考1", "A2L↔epd 診断CAN ID一致"
	try:
		a2l_ids = load_a2l_can_ids(a2l_path)
		epd_ids = load_epd_can_ids(epd_path)
	except (A2lParseError, EpdParseError) as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	fmt = lambda s: ", ".join(f"0x{i:08X}" for i in sorted(s))
	detail = [f"A2L : {fmt(a2l_ids)}   ECUアドレス {format_ecu_addresses(a2l_ids)}",
	          f"epd : {fmt(epd_ids)}   ECUアドレス {format_ecu_addresses(epd_ids)}"]
	if a2l_ids == epd_ids:
		return CheckResult(key, title, STATUS_OK,
		                   f"一致 (ECUアドレス {format_ecu_addresses(a2l_ids)})", detail)
	return CheckResult(key, title, STATUS_NG,
	                   f"不一致 (A2L={format_ecu_addresses(a2l_ids)} / "
	                   f"epd={format_ecu_addresses(epd_ids)})", detail)


# =====================================================================
# 参考2 A2L ↔ epd デバイス種別 / ② HEX ↔ A2L ビルドID一致
# =====================================================================

def check_a2l_epd_device(a2l_path: str, epd_path: str) -> CheckResult:
	"""A2Lの`/begin MODULE`のデバイス型番と、epdの`CPUTYPE`が一致するか。

	どちらもファイル内部の情報なので、ファイル名の付け替えに影響されない。
	デバイス世代違い(TC1782 / TC275 / RH850 を TC277用プロファイルに当てる等)を
	検出する。
	"""
	key, title = "参考2", "A2L↔epd デバイス種別一致"
	try:
		module = load_a2l_module(a2l_path)
		cpu_type = load_epd_cpu_type(epd_path)
	except OSError as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	if not module:
		return CheckResult(key, title, STATUS_ERROR,
		                   "A2Lに'/begin MODULE'が見つかりません")
	if not cpu_type:
		return CheckResult(key, title, STATUS_ERROR,
		                   "epdに'CPUTYPE'が見つかりません")

	a2l_dev = normalize_device(module)
	epd_dev = normalize_device(cpu_type)
	detail = [f"A2L MODULE  : {module} → {a2l_dev}",
	          f"epd CPUTYPE : {cpu_type} → {epd_dev}"]
	if a2l_dev and a2l_dev == epd_dev:
		return CheckResult(key, title, STATUS_OK, f"一致 ({a2l_dev})", detail)
	return CheckResult(key, title, STATUS_NG,
	                   f"不一致 (A2L={a2l_dev} / epd={epd_dev})"
	                   " → デバイス種別が異なるプロファイル", detail)


def check_hex_a2l_pairing(hex_path: str, a2l_path: str) -> CheckResult:
	"""HEXトレーラのDBF名と、A2Lのファイル名(拡張子なし)が一致するか。

	HEXのトレーラはそのHEXをビルドした定義ファイル名を保持している
	(例: DBF:CX1G5DB0_r2.dbo ↔ CX1G5DB0_r2.a2l)。

	注意: ビルドIDはA2Lの**内容には現れない**(実測18件で確認)ため、
	照合相手はA2Lのファイル名しかない。A2Lを改名すると不一致になる点に注意。
	"""
	key, title = "②", "HEX↔A2L ビルドID一致"
	try:
		trailer = load_hex_trailer(hex_path)
	except OSError as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	dbf = trailer.get("DBF", "")
	if not dbf:
		return CheckResult(key, title, STATUS_ERROR,
		                   "HEXトレーラにDBF名がありません(照合できません)")

	build_id = re.sub(r"\.dbo$", "", dbf, flags=re.IGNORECASE)
	a2l_stem = Path(a2l_path).stem
	detail = [f"HEXトレーラ DBF : {dbf}", f"A2Lファイル名    : {Path(a2l_path).name}"]
	if build_id.upper() == a2l_stem.upper():
		return CheckResult(key, title, STATUS_OK, f"一致 ({build_id})", detail)
	return CheckResult(key, title, STATUS_NG,
	                   f"不一致 (HEXは {build_id} 由来 / A2Lは {a2l_stem})"
	                   " → その HEX に対応しない A2L", detail)


# =====================================================================
# ③ A2L内の全コード領域にデータあり
# =====================================================================

def check_a2l_coverage(hex_path: str, a2l_path: str) -> CheckResult:
	"""A2Lが宣言するコード領域すべてにHEXのデータが存在するか。

	宣言されたコード領域にデータが無い = イメージが不完全。
	"""
	key, title = "③", "A2L内の全コード領域にデータあり"
	try:
		segments = load_a2l_code_segments(a2l_path)
		write_ranges = load_hex_write_ranges(hex_path)
	except (A2lParseError, HexParseError) as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	missing = subtract_ranges(segments, write_ranges)
	detail = [f"A2L宣言 : {format_ranges(segments)}",
	          f"HEX     : {format_ranges(write_ranges)}"]
	if not missing:
		return CheckResult(key, title, STATUS_OK, "宣言領域すべてにデータあり", detail)
	total = sum(e - s for s, e in missing)
	return CheckResult(key, title, STATUS_NG,
	                   f"データ未供給 {format_size(total)}: {format_ranges(missing)}",
	                   detail)


# =====================================================================
# 参考3 HEX ↔ epd アドレス空間の重なり(A2L無し時の代替)
# =====================================================================

def check_hex_epd_overlap(hex_path: str, epd_path: str) -> CheckResult:
	"""HEXの書き込み範囲とepdの許可範囲が重なるか(アドレス空間の妥当性)。

	①∧④が成立すれば数学的に必ず成立するため、通常は実施しない。
	A2Lが提出されなかった場合の唯一の代替手段としてのみ使う。
	"""
	key, title = "参考3", "HEX↔epd アドレス空間の重なり(A2L無し時の代替)"
	try:
		write_ranges = load_hex_write_ranges(hex_path)
		allowed = load_epd_rom_ranges(epd_path)
	except (HexParseError, EpdParseError) as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	overlap = intersect_size(write_ranges, allowed)
	epd_start = allowed[0][0]
	start_covered = any(s <= epd_start < e for s, e in write_ranges)
	detail = [f"HEX     : {format_ranges(write_ranges)}",
	          f"epd許可 : {format_ranges(allowed)}",
	          f"重なり  : {format_size(overlap)}"]
	if overlap > 0 and start_covered:
		return CheckResult(key, title, STATUS_OK,
		                   f"重なり {format_size(overlap)}", detail)
	reason = "重なりなし" if overlap == 0 else f"epd先頭 0x{epd_start:08X} にHEXデータなし"
	return CheckResult(key, title, STATUS_NG,
	                   f"{reason} → 別デバイス/別アドレス空間のHEXの可能性", detail)


# =====================================================================
# 参考4 boot Ver.照合(要求時のみ)
# =====================================================================

def read_boot_version(hex_path: str, model: str):
	"""指定機種のboot Ver.アドレスからバイト列を読み、16進文字列で返す"""
	if model not in BOOT_VER_ADDRESSES:
		raise KeyError(model)
	raw = read_hex_bytes(hex_path, BOOT_VER_ADDRESSES[model], BOOT_VER_LENGTH)
	return raw.hex().upper() if raw else None


def check_boot_version(hex_path: str, model: str, expected: str = None,
                       reference_hex: str = None) -> CheckResult:
	"""boot Ver.を読み出し、期待値または基準HEXと一致するか確認する。

	Excel版のCompareBootVerと同等だが、レコード先頭に一致しないアドレス
	(例: 0xA00048BC)でも読み出せる点が異なる。
	expected も reference_hex も無い場合は読み出し値の報告のみ。
	"""
	key, title = "参考4", "boot Ver.照合"
	if not model:
		return CheckResult(key, title, STATUS_SKIP, "機種が指定されていません")
	if model not in BOOT_VER_ADDRESSES:
		return CheckResult(key, title, STATUS_ERROR,
		                   f"boot Ver.アドレス未登録の機種です: {model}")
	try:
		actual = read_boot_version(hex_path, model)
	except HexParseError as e:
		return CheckResult(key, title, STATUS_ERROR, str(e))

	address = BOOT_VER_ADDRESSES[model]
	detail = [f"機種 : {model}", f"アドレス : 0x{address:08X}",
	          f"読出値 : {actual or '(取得不可)'}"]
	if actual is None:
		return CheckResult(key, title, STATUS_ERROR,
		                   f"0x{address:08X} にデータがありません(取得不可)", detail)

	if reference_hex:
		try:
			ref = read_boot_version(reference_hex, model)
		except HexParseError as e:
			return CheckResult(key, title, STATUS_ERROR, str(e), detail)
		detail.append(f"基準HEX読出値 : {ref or '(取得不可)'}")
		if ref is None:
			return CheckResult(key, title, STATUS_ERROR,
			                   "基準HEXから取得できません", detail)
		if ref == actual:
			return CheckResult(key, title, STATUS_OK, f"基準HEXと一致 ({actual})", detail)
		return CheckResult(key, title, STATUS_NG,
		                   f"基準HEXと不一致 ({actual} ≠ {ref})", detail)

	if expected:
		want = expected.replace(" ", "").upper()
		if actual == want:
			return CheckResult(key, title, STATUS_OK, f"期待値と一致 ({actual})", detail)
		return CheckResult(key, title, STATUS_NG,
		                   f"期待値と不一致 ({actual} ≠ {want})", detail)

	return CheckResult(key, title, STATUS_WARN,
	                   f"読出のみ(期待値未指定): {actual}", detail)


# =====================================================================
# 実行
# =====================================================================

def run_all(hex_path: str, epd_path: str, a2l_path: str = None,
            model: str = None, expected_boot_ver: str = None,
            reference_hex: str = None, extra: bool = False) -> list:
	"""チェックを実行し、CheckResultのリストを返す。

	判定対象(NG)は仕様どおり ①②③ の3項目のみ。
	a2l_path が無い場合は黙って通さず STATUS_ERROR(確認不能)とする(fail-closed)。

	extra=True のとき、仕様外の参考チェック(診断CAN ID / デバイス種別)も
	実施する。参考チェックは判定(終了コード)には影響しない。
	"""
	results = []
	if a2l_path:
		results.append(check_a2l_epd_start(a2l_path, epd_path))       # ①
		results.append(check_hex_a2l_pairing(hex_path, a2l_path))     # ②
		results.append(check_a2l_coverage(hex_path, a2l_path))        # ③
		if extra:
			results.append(_as_reference(check_a2l_epd_canid(a2l_path, epd_path)))
			results.append(_as_reference(check_a2l_epd_device(a2l_path, epd_path)))
	else:
		for k, t in (("①", "A2L↔epd 先頭アドレス一致"),
		             ("②", "HEX↔A2L ビルドID一致"),
		             ("③", "A2L内の全コード領域にデータあり")):
			results.append(CheckResult(k, t, STATUS_ERROR, "A2Lが指定されていません"))
		# A2Lが無いときの唯一の手掛かりとしてアドレス空間の重なりだけ見る
		results.append(_as_reference(check_hex_epd_overlap(hex_path, epd_path)))

	# boot Ver.照合は判定対象外(要求時のみ)。boot Ver.確認ツール(Excel)相当。
	if model or expected_boot_ver or reference_hex:
		results.append(_as_reference(
			check_boot_version(hex_path, model, expected_boot_ver, reference_hex)))
	return results


def _as_reference(result: CheckResult) -> CheckResult:
	"""参考チェックの結果を判定に影響しない状態へ落とす。

	NG相当でも終了コードを変えないよう STATUS_WARN に置き換える。
	"""
	if result.status == STATUS_NG:
		result.status = STATUS_WARN
		result.message = "(参考) " + result.message
	return result


# =====================================================================
# 情報表示(判定には影響しない)
# =====================================================================

def describe_inputs(hex_path: str, epd_path: str, a2l_path: str = None) -> list:
	"""入力ファイルから読み取れる素性を行のリストで返す。

	ECUアドレス(診断CAN IDに埋め込まれた物理アドレス)やデバイス型番など、
	判定はしないが確認に有用な情報を表示するためのもの。
	"""
	lines = []
	try:
		lines.append(f"ECUアドレス(epd) : {format_ecu_addresses(load_epd_can_ids(epd_path))}")
	except (EpdParseError, OSError):
		lines.append("ECUアドレス(epd) : (取得不可)")
	try:
		cpu = load_epd_cpu_type(epd_path)
		lines.append(f"デバイス(epd)     : {cpu or '(取得不可)'}")
	except OSError:
		pass
	if a2l_path:
		try:
			lines.append(f"ECUアドレス(A2L) : {format_ecu_addresses(load_a2l_can_ids(a2l_path))}")
		except (A2lParseError, OSError):
			lines.append("ECUアドレス(A2L) : (取得不可)")
		try:
			mod = load_a2l_module(a2l_path)
			lines.append(f"デバイス(A2L)     : {mod or '(取得不可)'}")
		except OSError:
			pass
	try:
		trailer = load_hex_trailer(hex_path)
		if trailer.get("DBF"):
			lines.append(f"ビルドID(HEX)     : {trailer['DBF']}")
		for k, label in (("PVR", "プログラムVer.アドレス"), ("DVR", "データVer.アドレス")):
			if trailer.get(k):
				lines.append(f"{label}(HEX) : 0x{trailer[k]}")
	except OSError:
		pass
	return lines

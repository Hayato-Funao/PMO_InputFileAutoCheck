"""
チェック③(HEX/A2Lチェック)のテスト。標準ライブラリのunittestのみを使用する。

実行方法(どちらでも可):
	python -m unittest test_xpx_checks -v      # check_prof直下から
	python テスト\test_xpx_checks.py           # テストフォルダ内から直接

①②のテスト(`check_mao\テスト\test_can_mao_check.py`等)は素のassert+手書きの
`run_test()`方式だが、③はunittestを使う。unittestも標準ライブラリのみで動くため
「外部ライブラリに依存しない」という本ツールの方針には反しない。末尾に①②と
同じ`run_test()`の起動口も用意してある。
"""

import os
import sys
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

# テストフォルダ内から直接実行された場合でも親フォルダのモジュールをimportできるようにする
# (①`check_mao\テスト\test_can_mao_check.py`と同じ作法)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import prof_check  # noqa: E402
import prof_scan  # noqa: E402
import xpx_checks as C  # noqa: E402
from xpx_inputs import (
	A2lParseError,
	EpdParseError,
	HexParseError,
	intersect_size,
	load_a2l_can_ids,
	load_a2l_code_segments,
	load_epd_can_ids,
	load_epd_rom_ranges,
	load_hex_trailer,
	load_hex_write_ranges,
	merge_ranges,
	read_hex_bytes,
	subtract_ranges,
)


def hex_record(record_type: int, addr: int, data: bytes) -> str:
	"""Intel HEXレコードを1行生成する(チェックサム自動計算)"""
	raw = bytes([len(data), (addr >> 8) & 0xFF, addr & 0xFF, record_type]) + data
	return ":" + (raw + bytes([(-sum(raw)) & 0xFF])).hex().upper()


def ext_linear(upper: int) -> str:
	return hex_record(0x04, 0, bytes([(upper >> 8) & 0xFF, upper & 0xFF]))


def ext_segment(seg: int) -> str:
	return hex_record(0x02, 0, bytes([(seg >> 8) & 0xFF, seg & 0xFF]))


EOF_RECORD = ":00000001FF"

A2L_TEMPLATE = """\
/begin PROJECT _TEST ""
/begin MODULE {module} ""
/begin MEMORY_LAYOUT
\tcode1 "Main1" CODE FLASH INTERN {start} {size} -1 -1 -1 -1 -1
/end MEMORY_LAYOUT
/begin CAN
\t500000
\t/begin ADDRESS\t/* 拡張IDを使用する場合はbit31をTrueにする */
\t\t{can_rx}\t/* CAN_ID ECU -> INCA */
\t\t{can_tx}\t/* CAN_ID INCA -> ECU */
\t/end ADDRESS
/end CAN
"""


class Fixtures(unittest.TestCase):
	"""一時ファイルを生成するヘルパー"""

	def setUp(self):
		self._tmp = TemporaryDirectory()
		self.dir = Path(self._tmp.name)
		self.addCleanup(self._tmp.cleanup)

	def hexfile(self, lines, name="test.hex", trailer=None):
		path = self.dir / name
		blob = ("\r\n".join(lines) + "\r\n").encode("ascii")
		if trailer:
			blob += b"\x1a" + trailer.encode("shift_jis")
		path.write_bytes(blob)
		return str(path)

	def full_hex(self, upper=0xA000, low=0x0000, nbytes=64, name="test.hex", trailer=None):
		lines = [ext_linear(upper)]
		for off in range(0, nbytes, 16):
			lines.append(hex_record(0x00, low + off, bytes(range(16))))
		lines.append(EOF_RECORD)
		return self.hexfile(lines, name, trailer)

	def epdfile(self, roms, canidr=0x18DA10F1, canids=0x18DAF110, name="test.epd"):
		attrs = " ".join(f'ROM{i:02d}_ADDRESS="{a}" ROM{i:02d}_SIZE="{s}"'
		                 for i, (a, s) in enumerate(roms))
		xml = ('<?xml version="1.0" encoding="Shift_JIS" ?>\n'
		       f'<profile CPUTYPE="Infineon TC277" CANIDR="{canidr}" '
		       f'CANIDS="{canids}" {attrs} />\n')
		path = self.dir / name
		path.write_text(xml, encoding="shift_jis")
		return str(path)

	def a2lfile(self, start=0xA0010000, size=0xF0000,
	            can_rx=0x98DA10F1, can_tx=0x98DAF110, name="BUILD_r2.a2l",
	            module="TC277"):
		path = self.dir / name
		# 実物のA2Lはコメントに日本語を含み Shift_JIS で保存されている
		path.write_text(A2L_TEMPLATE.format(
			start=f"0x{start:08X}", size=f"0x{size:X}", module=module,
			can_rx=f"0x{can_rx:08X}", can_tx=f"0x{can_tx:08X}"), encoding="shift_jis")
		return str(path)


# =====================================================================
# パーサ
# =====================================================================

class TestRangeHelpers(unittest.TestCase):

	def test_merge(self):
		self.assertEqual(merge_ranges([(0x100, 0x200), (0x200, 0x300), (0x180, 0x1C0)]),
		                 [(0x100, 0x300)])

	def test_subtract(self):
		self.assertEqual(subtract_ranges([(0, 0x400)], [(0, 0x100), (0x300, 0x400)]),
		                 [(0x100, 0x300)])

	def test_subtract_fully_covered(self):
		self.assertEqual(subtract_ranges([(0x100, 0x200)], [(0, 0x400)]), [])

	def test_intersect_size(self):
		self.assertEqual(intersect_size([(0, 0x200)], [(0x100, 0x400)]), 0x100)
		self.assertEqual(intersect_size([(0, 0x100)], [(0x200, 0x300)]), 0)


class TestHexParser(Fixtures):

	def test_write_ranges(self):
		p = self.full_hex(nbytes=32)
		self.assertEqual(load_hex_write_ranges(p), [(0xA0000000, 0xA0000020)])

	def test_extended_segment_record(self):
		p = self.hexfile([ext_segment(0x1000), hex_record(0x00, 0x20, bytes(16)), EOF_RECORD])
		self.assertEqual(load_hex_write_ranges(p), [(0x00010020, 0x00010030)])

	def test_start_address_records_ignored(self):
		p = self.hexfile([hex_record(0x05, 0, bytes([0xA0, 0, 0, 0])),
		                  ext_linear(0xA000), hex_record(0x00, 0, bytes(16)), EOF_RECORD])
		self.assertEqual(load_hex_write_ranges(p), [(0xA0000000, 0xA0000010)])

	def test_checksum_error(self):
		good = hex_record(0x00, 0, bytes(16))
		broken = good[:-2] + ("00" if good[-2:] != "00" else "01")
		p = self.hexfile([ext_linear(0xA000), broken, EOF_RECORD])
		with self.assertRaises(HexParseError):
			load_hex_write_ranges(p)

	def test_byte_count_mismatch(self):
		p = self.hexfile([ext_linear(0xA000), ":10000000000000000000FF", EOF_RECORD])
		with self.assertRaises(HexParseError):
			load_hex_write_ranges(p)

	def test_trailer_with_shift_jis(self):
		"""0x1A以降のトレーラ(日本語を含む)で解析が止まること"""
		p = self.full_hex(trailer=" COM:          DBF:BUILD_r2.dbo    HIS:HILS用ALL.his ")
		self.assertEqual(load_hex_write_ranges(p), [(0xA0000000, 0xA0000040)])

	def test_missing_eof_is_tolerated(self):
		"""EOFレコード欠落は解析エラーにしない(ツールチェーン依存の仮定を持たない)"""
		p = self.hexfile([ext_linear(0xA000), hex_record(0x00, 0, bytes(16))])
		self.assertEqual(load_hex_write_ranges(p), [(0xA0000000, 0xA0000010)])

	def test_trailer_fields(self):
		p = self.full_hex(trailer=" COM:  DBF:BUILD_r2.dbo   MAF:BUILD_r2.MAO   PVR:A0011FB0 ")
		t = load_hex_trailer(p)
		self.assertEqual(t["DBF"], "BUILD_r2.dbo")
		self.assertEqual(t["MAF"], "BUILD_r2.MAO")
		self.assertEqual(t["PVR"], "A0011FB0")

	def test_read_hex_bytes_unaligned(self):
		"""レコード先頭でないアドレスからも読めること(Excel版との差分)"""
		p = self.full_hex(nbytes=32)
		self.assertEqual(read_hex_bytes(p, 0xA000000C, 4), bytes([12, 13, 14, 15]))

	def test_read_hex_bytes_across_records(self):
		p = self.full_hex(nbytes=32)
		self.assertEqual(read_hex_bytes(p, 0xA000000E, 4), bytes([14, 15, 0, 1]))

	def test_read_hex_bytes_missing_returns_none(self):
		p = self.full_hex(nbytes=16)
		self.assertIsNone(read_hex_bytes(p, 0xA0000010, 4))


class TestEpdParser(Fixtures):

	def test_rom_ranges_merge_and_skip_zero(self):
		p = self.epdfile([(0, 0), (0xA0010000, 0x4000), (0xA0014000, 0x4000)])
		self.assertEqual(load_epd_rom_ranges(p), [(0xA0010000, 0xA0018000)])

	def test_can_ids_masked(self):
		p = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(load_epd_can_ids(p), {0x18DA10F1, 0x18DAF110})

	def test_no_rom_raises(self):
		path = self.dir / "bad.epd"
		path.write_text('<?xml version="1.0" ?>\n<profile />\n', encoding="shift_jis")
		with self.assertRaises(EpdParseError):
			load_epd_rom_ranges(str(path))


class TestA2lParser(Fixtures):

	def test_code_segments(self):
		p = self.a2lfile(start=0xA0010000, size=0xF0000)
		self.assertEqual(load_a2l_code_segments(p), [(0xA0010000, 0xA0100000)])

	def test_can_ids_bit31_masked(self):
		p = self.a2lfile(can_rx=0x98DA10F1, can_tx=0x98DAF110)
		self.assertEqual(load_a2l_can_ids(p), {0x18DA10F1, 0x18DAF110})

	def test_no_code_flash_raises(self):
		path = self.dir / "bad.a2l"
		path.write_text("/begin PROJECT\n/end PROJECT\n", encoding="ascii")
		with self.assertRaises(A2lParseError):
			load_a2l_code_segments(str(path))


# =====================================================================
# チェック
# =====================================================================

class TestChecks(Fixtures):

	def test_start_address_match(self):
		a2l = self.a2lfile(start=0xA0010000)
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_a2l_epd_start(a2l, epd).status, C.STATUS_OK)

	def test_start_address_mismatch(self):
		a2l = self.a2lfile(start=0xA0010000)
		epd = self.epdfile([(0xA0080000, 0x4000)])
		r = C.check_a2l_epd_start(a2l, epd)
		self.assertEqual(r.status, C.STATUS_NG)
		self.assertIn("0xA0080000", r.message)

	def test_start_address_unparsable_is_error_not_ok(self):
		"""AutoFlashHexはfail-openだが、こちらは確認不能として扱う"""
		bad = self.dir / "bad.a2l"
		bad.write_text("nothing here\n", encoding="ascii")
		epd = self.epdfile([(0xA0010000, 0x4000)])
		r = C.check_a2l_epd_start(str(bad), epd)
		self.assertEqual(r.status, C.STATUS_ERROR)
		self.assertNotEqual(r.status, C.STATUS_OK)

	def test_canid_match(self):
		a2l = self.a2lfile()
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_a2l_epd_canid(a2l, epd).status, C.STATUS_OK)

	def test_canid_mismatch(self):
		a2l = self.a2lfile(can_rx=0x98DA07F1, can_tx=0x98DAF107)
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_a2l_epd_canid(a2l, epd).status, C.STATUS_NG)

	def test_pairing_match(self):
		h = self.full_hex(trailer=" DBF:BUILD_r2.dbo ")
		a2l = self.a2lfile(name="BUILD_r2.a2l")
		self.assertEqual(C.check_hex_a2l_pairing(h, a2l).status, C.STATUS_OK)

	def test_pairing_mismatch_is_ng(self):
		"""仕様②: HEXに対応しないA2Lは NG"""
		h = self.full_hex(trailer=" DBF:OTHER_r2.dbo ")
		a2l = self.a2lfile(name="BUILD_r2.a2l")
		r = C.check_hex_a2l_pairing(h, a2l)
		self.assertEqual(r.status, C.STATUS_NG)

	def test_device_match(self):
		a2l = self.a2lfile(module="TC277")
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_a2l_epd_device(a2l, epd).status, C.STATUS_OK)

	def test_device_mismatch_generation(self):
		"""TC1782(旧世代)をTC277用プロファイルに当てたケース"""
		a2l = self.a2lfile(module="TC1782")
		epd = self.epdfile([(0xA0010000, 0x4000)])
		r = C.check_a2l_epd_device(a2l, epd)
		self.assertEqual(r.status, C.STATUS_NG)
		self.assertIn("TC1782", r.message)

	def test_device_mismatch_vendor(self):
		a2l = self.a2lfile(module="RH850C1M")
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_a2l_epd_device(a2l, epd).status, C.STATUS_NG)

	def test_device_normalizes_epd_suffix(self):
		"""epdの'Infineon TC387_暫定'とA2Lの'TC387'を一致とみなせること"""
		from xpx_inputs import normalize_device
		self.assertEqual(normalize_device("Infineon TC387_暫定"), "TC387")
		self.assertEqual(normalize_device("TC387"), "TC387")
		self.assertEqual(normalize_device("Infineon TC277"), "TC277")

	def test_pairing_no_trailer_is_error(self):
		h = self.full_hex()
		a2l = self.a2lfile(name="BUILD_r2.a2l")
		self.assertEqual(C.check_hex_a2l_pairing(h, a2l).status, C.STATUS_ERROR)

	def test_coverage_ok(self):
		h = self.full_hex(upper=0xA001, low=0x0000, nbytes=64)
		a2l = self.a2lfile(start=0xA0010000, size=0x40)
		self.assertEqual(C.check_a2l_coverage(h, a2l).status, C.STATUS_OK)

	def test_coverage_missing_is_ng(self):
		h = self.full_hex(upper=0xA001, low=0x0000, nbytes=32)
		a2l = self.a2lfile(start=0xA0010000, size=0x1000)
		r = C.check_a2l_coverage(h, a2l)
		self.assertEqual(r.status, C.STATUS_NG)

	def test_overlap_ok(self):
		h = self.full_hex(upper=0xA001, nbytes=64)
		epd = self.epdfile([(0xA0010000, 0x40)])
		self.assertEqual(C.check_hex_epd_overlap(h, epd).status, C.STATUS_OK)

	def test_overlap_none_is_ng(self):
		h = self.full_hex(upper=0x0000, low=0x0000, nbytes=64)
		epd = self.epdfile([(0xA0010000, 0x4000)])
		self.assertEqual(C.check_hex_epd_overlap(h, epd).status, C.STATUS_NG)

	def test_hex_below_epd_start_is_not_ng(self):
		"""撤回した旧ロジックの回帰テスト:
		HEXがepd許可範囲より低位に書き込んでもNGにしてはならない
		(全ての正常HEXが該当するため)。
		"""
		lines = [ext_linear(0xA000)]
		for off in range(0, 0x40, 16):
			lines.append(hex_record(0x00, off, bytes(16)))
		lines.append(ext_linear(0xA001))
		for off in range(0, 0x40, 16):
			lines.append(hex_record(0x00, off, bytes(16)))
		lines.append(EOF_RECORD)
		h = self.hexfile(lines)
		epd = self.epdfile([(0xA0010000, 0x40)])
		self.assertEqual(C.check_hex_epd_overlap(h, epd).status, C.STATUS_OK)

	def test_partial_image_is_not_ng(self):
		"""回帰テスト: 部分イメージ(epdが消去する範囲にHEXデータが無い)を
		NGにしてはならない。実在の D7805(2560KB) が該当する。
		"""
		h = self.full_hex(upper=0xA001, nbytes=32, trailer=" DBF:BUILD_r2.dbo ")
		epd = self.epdfile([(0xA0010000, 0x4000)])
		a2l = self.a2lfile(start=0xA0010000, size=0x20, name="BUILD_r2.a2l")
		results = C.run_all(h, epd, a2l)
		self.assertFalse([r for r in results if r.is_ng],
		                 [f"{r.key}:{r.message}" for r in results if r.is_ng])

	def test_boot_version_expected_match(self):
		h = self.full_hex(upper=0xA000, low=0x1DB0, nbytes=16)
		want = bytes(range(16)).hex().upper()
		r = C.check_boot_version(h, "BZ1_CS非対応機種", expected=want)
		self.assertEqual(r.status, C.STATUS_OK)

	def test_boot_version_expected_mismatch(self):
		h = self.full_hex(upper=0xA000, low=0x1DB0, nbytes=16)
		r = C.check_boot_version(h, "BZ1_CS非対応機種", expected="00" * 16)
		self.assertEqual(r.status, C.STATUS_NG)

	def test_boot_version_two_file_compare(self):
		a = self.full_hex(upper=0xA000, low=0x1DB0, nbytes=16, name="a.hex")
		b = self.full_hex(upper=0xA000, low=0x1DB0, nbytes=16, name="b.hex")
		r = C.check_boot_version(a, "BZ1_CS非対応機種", reference_hex=b)
		self.assertEqual(r.status, C.STATUS_OK)

	def test_boot_version_unknown_model_is_error(self):
		h = self.full_hex()
		self.assertEqual(C.check_boot_version(h, "存在しない機種").status, C.STATUS_ERROR)

	def test_boot_version_absent_data_is_error(self):
		h = self.full_hex(upper=0xA001, nbytes=16)
		r = C.check_boot_version(h, "BZ1_CS非対応機種")
		self.assertEqual(r.status, C.STATUS_ERROR)


class TestEcuAddress(unittest.TestCase):
	"""診断CAN IDからのECUアドレス解読"""

	def test_decodes_request_and_response(self):
		from xpx_inputs import ecu_addresses_from_can_ids
		self.assertEqual(ecu_addresses_from_can_ids({0x18DA10F1, 0x18DAF110}), {0x10})
		self.assertEqual(ecu_addresses_from_can_ids({0x18DA07F1, 0x18DAF107}), {0x07})
		self.assertEqual(ecu_addresses_from_can_ids({0x18DA0EF1, 0x18DAF10E}), {0x0E})

	def test_labels_ecu_type(self):
		from xpx_inputs import format_ecu_addresses
		self.assertEqual(format_ecu_addresses({0x18DA10F1}), "0x10 (FI)")
		self.assertEqual(format_ecu_addresses({0x18DA07F1}), "0x07 (ICM)")
		self.assertIn("種別未登録", format_ecu_addresses({0x18DA99F1}))

	def test_ignores_non_diagnostic_ids(self):
		from xpx_inputs import ecu_addresses_from_can_ids, format_ecu_addresses
		self.assertEqual(ecu_addresses_from_can_ids({0x12345678}), set())
		self.assertIn("不明", format_ecu_addresses({0x12345678}))


class TestRunAll(Fixtures):

	def test_spec_has_exactly_three_graded_checks(self):
		"""仕様どおり NG判定対象は ①②③ の3項目のみであること"""
		h = self.full_hex(upper=0xA001, nbytes=64, trailer=" DBF:BUILD_r2.dbo ")
		epd = self.epdfile([(0xA0010000, 0x40)])
		a2l = self.a2lfile(start=0xA0010000, size=0x40, name="BUILD_r2.a2l")
		results = C.run_all(h, epd, a2l)
		self.assertEqual([r.key for r in results], ["①", "②", "③"])

	def test_reference_checks_never_change_verdict(self):
		"""参考チェック(CAN ID/デバイス種別)が不一致でも判定はNGにならない"""
		h = self.full_hex(upper=0xA001, nbytes=64, trailer=" DBF:BUILD_r2.dbo ")
		# epdのCAN IDとデバイスをA2Lと食い違わせる
		epd = self.epdfile([(0xA0010000, 0x40)], canidr=0x18DA07F1, canids=0x18DAF107)
		a2l = self.a2lfile(start=0xA0010000, size=0x40, name="BUILD_r2.a2l",
		                   module="TC1782")
		results = C.run_all(h, epd, a2l, extra=True)
		refs = [r for r in results if r.key.startswith("参考")]
		self.assertEqual(len(refs), 2)
		self.assertTrue(all(r.status == C.STATUS_WARN for r in refs),
		                [f"{r.key}:{r.status}" for r in refs])
		self.assertFalse([r for r in results if r.is_ng])

	def test_missing_a2l_is_error_not_ok(self):
		"""A2L未指定時、A2Lを要するチェックを黙って通さないこと(fail-closed)"""
		h = self.full_hex(upper=0xA001, nbytes=64)
		epd = self.epdfile([(0xA0010000, 0x40)])
		results = C.run_all(h, epd, a2l_path=None)
		a2l_checks = [r for r in results if r.key in ("①", "②", "③")]
		self.assertEqual(len(a2l_checks), 3)
		self.assertTrue(all(r.status == C.STATUS_ERROR for r in a2l_checks))

	def test_all_ok_triple(self):
		h = self.full_hex(upper=0xA001, low=0x0000, nbytes=64,
		                  name="img.hex", trailer=" DBF:BUILD_r2.dbo ")
		epd = self.epdfile([(0xA0010000, 0x40)])
		a2l = self.a2lfile(start=0xA0010000, size=0x40, name="BUILD_r2.a2l")
		results = C.run_all(h, epd, a2l)
		self.assertFalse([r for r in results if r.is_ng], [r.key for r in results if r.is_ng])
		self.assertFalse([r for r in results if r.is_error])


# =====================================================================
# 統合アダプタ(prof_scan / prof_check)
# =====================================================================

class ProfFixtures(Fixtures):
	"""案件フォルダ(01_INPUT)を模した構成を組み立てる"""

	def case_dir(self, with_hex=True, with_epd=True, with_a2l=True,
	             extra_hex=False, a2l_name="BUILD_r2.a2l"):
		root = self.dir / "01_INPUT" / "01_HEX関連" / "mask1"
		root.mkdir(parents=True, exist_ok=True)
		if with_hex:
			src = Path(self.full_hex(upper=0xA001, nbytes=64, name="img.hex",
			                        trailer=" DBF:BUILD_r2.dbo "))
			(root / "img.hex").write_bytes(src.read_bytes())
		if extra_hex:
			# ファイル名が違えば重複排除キー(名前, サイズ)が別になるため2個と数えられる
			(root / "another.hex").write_bytes((root / "img.hex").read_bytes())
		if with_epd:
			src = Path(self.epdfile([(0xA0010000, 0x40)], name="prof.epd"))
			(root / "prof.epd").write_bytes(src.read_bytes())
		if with_a2l:
			src = Path(self.a2lfile(start=0xA0010000, size=0x40, name=a2l_name))
			(root / a2l_name).write_bytes(src.read_bytes())
		return str(self.dir / "01_INPUT")


class TestProfScan(ProfFixtures):

	def test_finds_all_three(self):
		root = self.case_dir()
		scan = prof_scan.scan_prof_inputs(root)
		self.assertTrue(scan.is_complete)
		self.assertEqual(scan.structural_ng_messages, [])

	def test_duplicate_hex_is_structural_ng(self):
		"""資料スライド2: 同種ファイルが複数なら判定不能(NG)"""
		root = self.case_dir(extra_hex=True)
		scan = prof_scan.scan_prof_inputs(root)
		self.assertFalse(scan.is_complete)
		self.assertIsNone(scan.hex_path)
		self.assertTrue(any("複数存在する" in m for m in scan.structural_ng_messages),
		                scan.structural_ng_messages)

	def test_same_file_in_two_roots_is_not_duplicate(self):
		"""回帰テスト: 01_INPUTとzip展開先に同じHEXがあっても2個と数えない"""
		root = self.case_dir()
		unzipped = self.dir / "unzipped"
		unzipped.mkdir()
		original = Path(root) / "01_HEX関連" / "mask1" / "img.hex"
		(unzipped / "img.hex").write_bytes(original.read_bytes())
		scan = prof_scan.scan_prof_inputs(root, str(unzipped))
		self.assertTrue(scan.is_complete, scan.structural_ng_messages)

	def test_finds_file_only_in_extra_root(self):
		"""zip展開先にしか無いファイルも拾えること"""
		root = self.case_dir(with_hex=False)
		unzipped = self.dir / "unzipped"
		unzipped.mkdir()
		src = Path(self.full_hex(upper=0xA001, nbytes=64, name="zipped.hex",
		                        trailer=" DBF:BUILD_r2.dbo "))
		(unzipped / "zipped.hex").write_bytes(src.read_bytes())
		scan = prof_scan.scan_prof_inputs(root, [str(unzipped)])
		self.assertTrue(scan.is_complete, scan.structural_ng_messages)

	def test_excel_lock_file_ignored(self):
		root = self.case_dir()
		mask = Path(root) / "01_HEX関連" / "mask1"
		(mask / "~$prof.epd").write_bytes(b"lock")
		scan = prof_scan.scan_prof_inputs(root)
		self.assertTrue(scan.is_complete, scan.structural_ng_messages)

	def test_missing_file_reported(self):
		root = self.case_dir(with_a2l=False)
		scan = prof_scan.scan_prof_inputs(root)
		self.assertFalse(scan.is_complete)
		self.assertTrue(any("A2L" in m and "提出されていない" in m
		                    for m in scan.structural_ng_messages),
		                scan.structural_ng_messages)

	def test_case_insensitive_suffix(self):
		root = self.case_dir(with_hex=False)
		mask = Path(root) / "01_HEX関連" / "mask1"
		src = Path(self.full_hex(upper=0xA001, nbytes=64, name="upper.HEX",
		                        trailer=" DBF:BUILD_r2.dbo "))
		(mask / "UPPER.HEX").write_bytes(src.read_bytes())
		scan = prof_scan.scan_prof_inputs(root)
		self.assertTrue(scan.is_complete, scan.structural_ng_messages)


class TestProfCheck(ProfFixtures):

	def test_ok_case(self):
		result, reason, detail = prof_check.run_check3(self.case_dir())
		self.assertEqual(result, prof_check.RESULT_OK)
		self.assertIsNone(reason)
		self.assertIn("=== OK ===", detail)

	def test_result_is_only_two_valued(self):
		"""確認不能・対象外を返さず OK/NG の2値であること(2026-09-02決定)"""
		for kwargs in ({}, {"with_a2l": False}, {"with_hex": False}, {"extra_hex": True}):
			result, _reason, _detail = prof_check.run_check3(self.case_dir(**kwargs))
			self.assertIn(result, (prof_check.RESULT_OK, prof_check.RESULT_NG), kwargs)
			self.setUp()  # 次のケース用にフォルダを作り直す

	def test_missing_a2l_is_ng_with_reason(self):
		"""A2L未提出は「確認不能」ではなくNG。理由が返ること"""
		result, reason, detail = prof_check.run_check3(self.case_dir(with_a2l=False))
		self.assertEqual(result, prof_check.RESULT_NG)
		self.assertIn("A2L", reason)
		self.assertIn("=== NG ===", detail)

	def test_detail_has_expected_sections(self):
		"""詳細ファイルの書式が②に揃っていること"""
		_result, _reason, detail = prof_check.run_check3(self.case_dir())
		for section in ("実行日時:", "=== 入力ファイル ===", "=== OK ===", "=== NG ==="):
			self.assertIn(section, detail)

	def test_check_ng_reports_reason(self):
		"""②(ビルドID不一致)でNGになり、理由が返ること"""
		result, reason, _detail = prof_check.run_check3(
			self.case_dir(a2l_name="OTHER_r2.a2l"))
		self.assertEqual(result, prof_check.RESULT_NG)
		self.assertIsNotNone(reason)


def run_test():
	"""①②のテストと同じ呼び出し方をできるようにした起動口。

	`python テスト\test_xpx_checks.py`で全件実行する。個別実行やパターン指定が
	必要な場合は`python -m unittest test_xpx_checks -v`を使う。
	"""
	loader = unittest.TestLoader()
	suite = loader.loadTestsFromModule(sys.modules[__name__])
	runner = unittest.TextTestRunner(verbosity=1)
	outcome = runner.run(suite)
	if outcome.wasSuccessful():
		print("すべてのテストに成功しました。")
	return outcome.wasSuccessful()


if __name__ == "__main__":
	sys.exit(0 if run_test() else 1)

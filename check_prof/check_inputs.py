"""
XPX INPUT不備チェック ③ のコマンドラインエントリポイント。

使い方:
  python check_inputs.py --hex <HEX> --epd <epd> [--a2l <A2L>]
                         [--model <機種名>] [--expect-boot-ver <16進>]
                         [--ref-hex <基準HEX>] [--extra] [--list-models]

終了コード:
  0 = 全チェックOK(警告のみ含む)
  1 = NG(ソフト/プロファイルの不備を検出)
  2 = 確認不能(入力が読めない、必要な入力が足りない)、または引数不備
"""

import argparse
import sys
from pathlib import Path

from xpx_checks import (
	BOOT_VER_ADDRESSES,
	describe_inputs,
	STATUS_ERROR,
	STATUS_NG,
	STATUS_OK,
	STATUS_SKIP,
	STATUS_WARN,
	run_all,
)

EXIT_OK = 0
EXIT_NG = 1
EXIT_ERROR = 2


def _configure_console():
	"""Windowsコンソール(cp932)で日本語が化けないようUTF-8に切り替える"""
	for stream in (sys.stdout, sys.stderr):
		reconfigure = getattr(stream, "reconfigure", None)
		if reconfigure is not None:
			try:
				reconfigure(encoding="utf-8", errors="replace")
			except (ValueError, OSError):
				pass


def build_parser() -> argparse.ArgumentParser:
	p = argparse.ArgumentParser(
		description="HEX / .epd / .a2l の整合性をチェックする",
		formatter_class=argparse.RawDescriptionHelpFormatter)
	p.add_argument("--hex", dest="hex_path", help="HEXファイル")
	p.add_argument("--epd", dest="epd_path", help="epd(プロファイル)ファイル")
	p.add_argument("--a2l", dest="a2l_path",
	               help="A2Lファイル(省略すると①②③が確認不能になります)")
	p.add_argument("--model", help="boot Ver.照合の機種名(--list-modelsで一覧)")
	p.add_argument("--expect-boot-ver", dest="expected",
	               help="boot Ver.の期待値(16進文字列)")
	p.add_argument("--ref-hex", dest="reference_hex",
	               help="boot Ver.比較の基準HEX(Excel版の2ファイル比較相当)")
	p.add_argument("--list-models", action="store_true",
	               help="boot Ver.アドレス登録済み機種を一覧表示して終了")
	p.add_argument("--extra", action="store_true",
	               help="仕様外の参考チェック(診断CAN ID / デバイス種別)も実施する"
	                    "(判定には影響しません)")
	return p


def print_report(results: list) -> None:
	"""チェック結果を表形式で表示する"""
	mark = {STATUS_OK: "  OK  ", STATUS_NG: "  NG  ", STATUS_WARN: " 警告 ",
	        STATUS_ERROR: "確認不能", STATUS_SKIP: " 対象外 "}
	width = max(len(r.title) for r in results)
	keyw = max(len(r.key) for r in results)
	print("=" * (width + 46))
	for r in results:
		print(f"{r.key.ljust(keyw)} {r.title.ljust(width)} "
		      f"[{mark.get(r.status, r.status)}] {r.message}")
		for line in r.detail:
			print(f"      {line}")
	print("=" * (width + 46))


def main(argv: list) -> int:
	_configure_console()
	parser = build_parser()
	args = parser.parse_args(argv)

	if args.list_models:
		print("boot Ver.アドレス登録済み機種:")
		for name, addr in BOOT_VER_ADDRESSES.items():
			print(f"  {name:<20} 0x{addr:08X}")
		return EXIT_OK

	if not args.hex_path or not args.epd_path:
		print("エラー: --hex と --epd は必須です", file=sys.stderr)
		parser.print_usage(sys.stderr)
		return EXIT_ERROR

	inputs = [("--hex", args.hex_path), ("--epd", args.epd_path)]
	if args.a2l_path:
		inputs.append(("--a2l", args.a2l_path))
	if args.reference_hex:
		inputs.append(("--ref-hex", args.reference_hex))
	for flag, path in inputs:
		if not Path(path).is_file():
			print(f"エラー: {flag} のファイルが見つかりません: {path}", file=sys.stderr)
			return EXIT_ERROR

	try:
		results = run_all(args.hex_path, args.epd_path, args.a2l_path,
		                  args.model, args.expected, args.reference_hex,
		                  extra=args.extra)
	except OSError as e:
		print(f"エラー: 入力ファイルを読めません: {e}", file=sys.stderr)
		return EXIT_ERROR

	print(f"HEX : {args.hex_path}")
	print(f"epd : {args.epd_path}")
	print(f"A2L : {args.a2l_path or '(未指定)'}")
	print("--- 入力ファイルから読み取れる情報(判定対象外) ---")
	for line in describe_inputs(args.hex_path, args.epd_path, args.a2l_path):
		print(f"  {line}")
	print_report(results)

	if any(r.is_ng for r in results):
		print("判定: NG(不備を検出しました)")
		return EXIT_NG
	if any(r.is_error for r in results):
		print("判定: 確認不能(入力が不足しているため判定できません)")
		return EXIT_ERROR
	print("判定: OK")
	return EXIT_OK


if __name__ == "__main__":
	sys.exit(main(sys.argv[1:]))

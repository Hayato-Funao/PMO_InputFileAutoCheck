"""
チェック③(HEX/A2Lチェック)の統合実行アダプタ(2026-09-02新規)。

`unified_main.py`から呼ばれ、②(check_fs_matrix)と同じ`(result, message)`系の
契約で結果を返す。判定ロジック本体は`xpx_checks`、入力ファイルの特定は
`prof_scan`が担当し、本モジュールは両者の接続と結果整形だけを行う。

【`確認不能`の畳み込み】(2026-09-02決定)
`xpx_checks`内部では`STATUS_ERROR = "確認不能"`を維持し、どの項目がなぜ確認できな
かったかを詳細ファイルに残す。一方`run_check3`の戻り値は`"OK"`/`"NG"`の2値に畳む。
PowerAppsが`InputFileCheckResult`列の`"NG"`部分文字列で赤/緑を判定しているため
(`check_mao\\result.py`の`build_result_summary`のコメント参照)、`確認不能`のままでは
緑になってしまう。②が「ファイルを開けない → 理由付きでNG」に統一しているのと
同じfail-closed思想に合わせる。
"""

import os
from datetime import datetime

import prof_scan
import xpx_checks

RESULT_OK = "OK"
RESULT_NG = "NG"

# 詳細ファイルの箇条書き記号(②`unified_main._write_check2_report`と同じ)
BULLET = "・"


def _describe_scan(scan_result):
	"""特定できた入力ファイルを詳細ファイル用の行リストにする"""
	lines = []
	for label, path in (
		(prof_scan.KIND_A2L, scan_result.a2l_path),
		(prof_scan.KIND_EPD, scan_result.epd_path),
		(prof_scan.KIND_HEX, scan_result.hex_path),
	):
		lines.append(f"{label} : {os.path.basename(path) if path else '(特定できず)'}")
	return lines


def _build_detail_message(scan_result, results, input_lines):
	"""詳細ファイルの本文を組み立てる。

	書式は`実行日時:` → `=== OK ===` → `=== NG ===`の順とする（元は②の詳細ファイルに
	合わせた形式。②側は2026-09-03にチェック対象ファイル単位のグループ形式
	（`check_fs_matrix.check2_report`）へ変更したが、③はこの2セクション形式のまま）。
	③固有の情報として入力ファイル一覧と`describe_inputs()`の情報行を前段に足す。
	"""
	ok_lines = []
	ng_lines = []

	# 構造的NG(未提出・複数存在)はチェック実行前に確定しているのでそのまま積む
	for message in scan_result.structural_ng_messages:
		ng_lines.append(message)

	for check in results:
		text = f"[{check.key}] {check.title}: {check.message}"
		if check.status == xpx_checks.STATUS_OK:
			ok_lines.append(text)
		elif check.status == xpx_checks.STATUS_SKIP:
			continue  # 未実施(対象外)は報告しない
		else:
			# NG・確認不能・警告はまとめてNG欄へ出す(確認不能もNG扱いにする決定に合わせる)
			ng_lines.append(f"{text}（{check.status}）")

	timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
	lines = [f"実行日時: {timestamp}", ""]

	lines.append("=== 入力ファイル ===")
	for text in _describe_scan(scan_result):
		lines.append(f"{BULLET}{text}")
	lines.append("")

	if input_lines:
		lines.append("=== 入力から読み取れる情報(判定対象外) ===")
		for text in input_lines:
			lines.append(f"{BULLET}{text}")
		lines.append("")

	lines.append("=== OK ===")
	for text in ok_lines:
		lines.append(f"{BULLET}{text}")
	lines.append("")
	lines.append("=== NG ===")
	for text in ng_lines:
		lines.append(f"{BULLET}{text}")

	return "\n".join(lines) + "\n"


def run_check3(local_input_dir, extra_search_roots=None):
	"""案件フォルダに対してチェック③を実行する。

	引数:
		local_input_dir: ダウンロード済み`01_INPUT`のローカルパス
		extra_search_roots: 追加の探索ルート(①のzip展開先
			`CaseScanResult.work_directory`を渡す)

	戻り値:
		(result, reason, detail_message)
			result: `"OK"`または`"NG"`(2値。確認不能もNGへ畳む)
			reason: NGの主因を1行で表した文字列(OK時はNone)。
				`InputFileCheckResult`列の行に`NG（理由）`として併記する用
			detail_message: 詳細ファイルへ書く本文
	"""
	scan_result = prof_scan.scan_prof_inputs(local_input_dir, extra_search_roots)

	if not scan_result.is_complete:
		# 入力が揃っていないのでチェック本体は実行できない。構造的NGだけを報告する
		detail_message = _build_detail_message(scan_result, [], [])
		return RESULT_NG, scan_result.structural_ng_messages[0], detail_message

	results = xpx_checks.run_all(
		scan_result.hex_path, scan_result.epd_path, scan_result.a2l_path
	)
	try:
		input_lines = xpx_checks.describe_inputs(
			scan_result.hex_path, scan_result.epd_path, scan_result.a2l_path
		)
	except Exception:
		# 情報表示は判定に影響しないため、失敗しても判定を止めない
		input_lines = []

	detail_message = _build_detail_message(scan_result, results, input_lines)

	ng_results = [check for check in results if check.is_ng]
	error_results = [check for check in results if check.is_error]
	if ng_results:
		return RESULT_NG, ng_results[0].message, detail_message
	if error_results:
		# 「確認できなかった」もNGへ畳む。理由は列の行にも残す
		return RESULT_NG, f"{error_results[0].title}を確認できず", detail_message
	return RESULT_OK, None, detail_message

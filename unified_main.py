"""①(check_mao)・②(check_fs_matrix)の判定ロジックを1つのSharePoint連携プロセスでまとめて実行する
統合実行ツールのエントリポイント（2026-09-02新規・方式B）。

【背景】方式A（①②を別プロセス・別タスクスケジューラで独立稼働。共有列は各々ETagマージ＋
①②③順ソート＋412リトライで守る）は実装・動作確認済みだが、①②を同じタイミングで自動実行した際に
共有列`InputFileCheckResult`の更新が競合下で毎回正しく収束するか確信が持てない、という懸念が
②担当者との協議で挙がった。本ツール（方式B）は、①②を1プロセスでまとめて実行し、アイテムごとに
①②行を**1回のMERGEで同時に**書き戻すことで、①↔②間の412（ETag不一致）競合を原理的に発生させない。

【自己完結化】①②の判定ロジック（`case_scan`/`can_info`/`mao_can_id`/`compare`/`result`、
`matrix_check`/`gst_command_check`/`definition_file_check`）は、いずれもSharePoint連携に依存しない
純粋なローカル処理であることを確認済みのため、本フォルダ直下の`check_mao`／`check_fs_matrix`
サブフォルダへ実体コピーして同梱している（2026-09-02改修。本番の`check_mao`／`check_fs_matrix`
フォルダ名・分割単位に合わせたサブフォルダ構成）。これにより、本フォルダ（`ツール本体`）を
どこへコピーしても単独で動作する（①②の元フォルダへの参照は不要）。ただし①②の判定ロジックの
改修時は、元フォルダとこのコピーの両方に反映する必要がある点に注意（重複管理）。

【②の判定条件をマクロ準拠版へ統一】2026-09-02改修。同梱コピーの`check_fs_matrix\\matrix_check.py`
／`gst_command_check.py`は、簡易判定（○×記号リストとの単純な突き合わせ）のままで本フォルダ直下の
マクロ準拠版`matrix_check_ref.py`／`gst_command_check_ref.py`（実マクロ modChkExec.ReadFromFS ／
ReadFromGST をリバースエンジニアリングした版）と判定条件が食い違っていたため、内容を同一の実装へ
差し替えた。これによりunified_main.pyの②チェーンは*_ref.pyと同じ条件でファイルを検査する
（マクロが停止するNGだけでなく、無警告で誤った結果を出すWARNもNG扱い）。*_ref.pyを改修した際は
同梱コピー側にも必ず反映すること。

SharePoint連携コード（`ntlm_proxy.py`/`sp_auth.py`/`sp_integration.py`）と実行用コードは、
元から本フォルダに集約済み。

【②の未完成を考慮したRUN_CHECK2フラグ】②の定義ファイルチェックが参照する2つのローカルパス
（`definition_file_check.SERVER_REFERENCE_FILE_PATH`・本ファイルの`SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`）
は、②担当者の個人環境（RJ067219）のパスのままで本番サーバには存在せず、本番では②の定義ファイル
チェックが必ずNGになる（例外は出ないが、判定はNG固定）。②担当者がこれらを本番用パスへ修正するまで、
`RUN_CHECK2 = False`（既定値）にしておくことで、②の行を列に書き込まず①のみを稼働させる。②担当者の
修正が完了したら`True`に切り替えるだけで②が列へ加わる。

【タスクスケジューラ】方式Bを採用する場合は本スクリプト1本のみを登録する（2時間間隔）。方式Aの
①②個別タスクと同時稼働させると同じ列を二重更新するため、方式A/Bはどちらか一方のみ運用すること。
"""

import os
import re
import shutil
import sys
import tempfile
from datetime import datetime

# ---- ①②判定ロジックの参照先 ----------------------------------------------------------
# 本フォルダ直下に実体コピーした①②の判定ロジックをsys.pathへ追加してimportする（2026-09-02改修。
# 「ツール本体フォルダだけコピーすれば動く」ようにするため、外部の①②フォルダではなく同梱の
# check_mao\/check_fs_matrix\を参照する。環境（ローカル開発／本番）によらずパスが一定になるため、
# デプロイ時の書き換えは不要）。
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
CHECK_MAO_DIR = os.path.join(_THIS_DIR, "check_mao")
CHECK_FS_MATRIX_DIR = os.path.join(_THIS_DIR, "check_fs_matrix")
CHECK_PROF_DIR = os.path.join(_THIS_DIR, "check_prof")
if not os.path.isdir(CHECK_MAO_DIR):
	raise RuntimeError(
		f"①(check_mao)の判定ロジックコピーが見つからない: {CHECK_MAO_DIR}"
		"（ツール本体フォルダ内のcheck_mao\\が欠けている。コピーし直すこと）"
	)
if not os.path.isdir(CHECK_FS_MATRIX_DIR):
	raise RuntimeError(
		f"②(check_fs_matrix)の判定ロジックコピーが見つからない: {CHECK_FS_MATRIX_DIR}"
		"（ツール本体フォルダ内のcheck_fs_matrix\\が欠けている。コピーし直すこと）"
	)
if not os.path.isdir(CHECK_PROF_DIR):
	raise RuntimeError(
		f"③(check_prof)の判定ロジックコピーが見つからない: {CHECK_PROF_DIR}"
		"（ツール本体フォルダ内のcheck_prof\が欠けている。コピーし直すこと）"
	)
sys.path.append(CHECK_MAO_DIR)
sys.path.append(CHECK_FS_MATRIX_DIR)
sys.path.append(CHECK_PROF_DIR)

import result as check1_result  # noqa: E402  ①分。pure（SharePoint連携に依存しない）のため直接import
from case_scan import scan_case_folder  # noqa: E402
from can_info import extract_matrix_can_id_records  # noqa: E402
from compare import compare_can_id_sources  # noqa: E402
from mao_can_id import extract_mao_can_id_records  # noqa: E402
from result import MaskCheckResult  # noqa: E402

import definition_file_check  # noqa: E402  ②分。pure（SharePoint連携に依存しない）のため直接import
import gst_command_check  # noqa: E402
import matrix_check  # noqa: E402

import prof_check  # noqa: E402  ③分。pure（SharePoint連携に依存しない）のため直接import

import sp_auth  # noqa: E402  本フォルダ内のSP連携コード
import sp_integration  # noqa: E402
import unified_result  # noqa: E402

# ---- 設定 ------------------------------------------------------------------------------
# ②の定義ファイルチェックを実行対象に含めるかどうか（2026-09-02改修でTrueへ切り替え）。
# Trueにすると②の行が`InputFileCheckResult`列へ加わり、②詳細ファイル
# （`FSマトリクス_定義ファイルチェック結果_{Title列}.txt`）が案件の`01_INPUT`へアップロードされる。
# 【注意】②の定義ファイルチェックは下記2つのローカルパスを参照する:
#   - `definition_file_check.SERVER_REFERENCE_FILE_PATH`（XPX定義ファイル管理ファイル）
#   - `SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`（下記。SharePointのURL指定時は実行時にダウンロード）
# どちらも実行環境に存在しない場合、②の[元表]／[fcl_list_o_FI]／[fcl_list_x_FI]は
# 「ファイルを開けない」でNG固定になる（例外は出ず、②全体がNGになるだけ）。本番サーバで
# 稼働させる際は、この2つが本番のパスを指しているかを必ず確認すること。
RUN_CHECK2 = True

# ③(check_prof)のHEX/A2Lチェックを実行対象に含めるかどうか（2026-09-02新規）。
# 初回デプロイ時はFalse（既存の①②の動作を一切変えないため）。動作確認後にTrueへ切り替える。
# 【動確時の確認事項】③がNGになった案件の内訳を必ず見ること。③はHEX/.epd/.a2lの
# いずれかが未提出ならNGにする（fail-closed）ため、HEXを伴わない申請（CAN情報のみの
# 変更等）が実在する場合は一律で赤になってしまう。その場合はスコープ判定の追加を検討する。
# 【2026-09-03改修】①②③を合わせた動作確認のためTrueへ切り替え。
RUN_CHECK3 = True

# ②分の第2参照ファイル（SharePoint上にある特定の申請に紐付かない固定ファイル。②の
# `input_check_main.SHAREPOINT_REFERENCE_FILE_PATH`と同じ役割）。
#
# 【2026-09-02改修】ブラウザからコピーしたSharePointのURLを直接指定できるようにした。以前は
# ローカルパスしか受け付けず、URLを入れると`openpyxl`が
# 「does not support .xlsx&action=default&mobileredirect=true file format」で失敗していた
# （openpyxlはローカルファイルしか開けず、URLの末尾クエリを拡張子と誤認するため）。URLの場合は
# 実行時に1回だけRESTでダウンロードし、その一時ファイルを②へ渡す（`_resolve_definition_reference`）。
# 次の3形式に対応する:
#   * Excel Onlineで開くリンク（`_layouts/15/Doc.aspx?sourcedoc={GUID}&file=...`）
#   * 「リンクのコピー」で得られる共有リンク（`.../:x:/r/sites/<site>/...xlsx?web=1`）
#   * 通常のURL（`.../sites/<site>/<ライブラリ>/...xlsx`）／ローカルの絶対パス
# 案件リストのサイト(`sp_auth.SITE_URL`=jphgt105596)とは別サイトのファイルでも取得できる。
#
# 元表ファイルは3ヶ月ごとに更新されるため、このリンクも定期的に更新が必要。
# （現在の値: FI-FailSafeMatrix元表(26.06.30).xlsx / sites/jphgt107305）
SHAREPOINT_REFERENCE_FILE_PATH_OR_URL = (
	r"https://globalhonda.sharepoint.com/:x:/r/sites/jphgt107305/_layouts/15/Doc.aspx?sourcedoc=%7B6FDCABAA-F3F1-4AE6-9F9A-AC20E2B0D1B4%7D&file=FI-FailSafeMatrix%25u5143%25u8868(26.06.30).xlsx&action=default&mobileredirect=true"
)
# ①専用の詳細結果（マスク別・NG明細。①`sp_main.DETAIL_FILE_NAME`と同名）
DETAIL_FILE_NAME_1 = "MAO_CANマトリクス比較結果.txt"

# ②専用の詳細結果（[FS Matrix]／[GSTCommand]／[元表]等の明細。②`input_check_main.DETAIL_FILE_NAME`
# と同名）の基本ファイル名。実際に出力・アップロードするファイル名は、どの申請の結果か一目で
# 分かるようTitle列（案件ID）を末尾に付けた`FSマトリクス_定義ファイルチェック結果_{Title列}.txt`
# （2026-09-02改修。`_check2_detail_file_name`が組み立てる）。中身の書式は②の元実装
# （`input_check_main.write_report`）と同じ「実行日時 → === OK === → === NG ===」のまま。
DETAIL_FILE_NAME_2 = "FSマトリクス_定義ファイルチェック結果.txt"

# ②詳細ファイル名へTitle列の値を埋め込む際に落とす文字。Windowsのファイル名禁止文字
# (\ / : * ? " < > |) に加え、SharePointのアップロードURL（`Files/add(url='...')`のODataリテラル。
# `'`が入るとURLが壊れる）やリンクHTML・URLとして解釈される文字（# % & + { } ~）、制御文字も
# まとめて置換対象にする。
_FILE_NAME_NG_CHARS = re.compile(r"""[\\/:*?"<>|#%&{}~+'\x00-\x1f]""")

# Title列が長い場合にファイル名が肥大しないよう、埋め込む部分の上限文字数。
_FILE_NAME_PART_MAX_LEN = 60

# ③専用の詳細結果（①②③の各判定と入力ファイル情報の明細）。書式は②と同じ
# 「実行日時 → === OK === → === NG ===」。
DETAIL_FILE_NAME_3 = "HEX_A2Lチェック結果.txt"

# ①②③統合の判定サマリ（`InputFileCheckResult`列のマージ後値）を書き込む集約ファイル。①②と同名。
AGGREGATE_FILE_NAME = "インプットファイルチェック結果.txt"

# `InputFileCheckResult`列の書き戻し（ETagによる楽観的並行制御）のリトライ上限。①②を同一プロセスで
# 一括投入するため①↔②間の412は発生しないが、本ツール以外の書き手（③・PMOの手動編集等）に対する
# 保険として残す。
_MERGE_RETRY_LIMIT = 5

# 運用者向けの実行結果レポート（SharePointへはアップロードしない。本ツールと同じフォルダへ出力する）
RUN_REPORT_PATH = os.path.join(_THIS_DIR, "unified_main_実行結果レポート.txt")


def _resolve_local_path(local_input_dir, fallback_dir, access_token, url):
	"""
	②分の対象ファイル(url)のローカルパスを解決する。`sp_integration.download_folder_tree`で
	既に取得済みの`local_input_dir`（01_INPUTツリー全体）配下からurlのbasenameに一致するファイルを
	探し、見つかればそれを使う（②分のファイルも`01_INPUT`配下にある想定のため、通常はここで
	見つかる）。見つからない場合のみ、`fallback_dir`へ直接ダウンロードする
	（②`input_check_main.download_file`と同じ単純ダウンロード）。
	"""
	basename = os.path.basename(url)
	for root, _dirs, files in os.walk(local_input_dir):
		if basename in files:
			return os.path.join(root, basename)

	fallback_path = os.path.join(fallback_dir, basename)
	sp_integration.download_file(access_token, url, fallback_path)
	return fallback_path


def _resolve_definition_reference(access_token, work_directory):
	"""
	②の第2参照ファイル（SharePoint上の元表ファイル）のローカルパスを解決する。

	`SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`がSharePointのURLなら`work_directory`へ
	ダウンロードしてそのパスを返す（`openpyxl`はローカルファイルしか開けないため、URLを
	そのまま渡すことはできない）。ローカルパスならそのまま返す。

	戻り値:
		(ローカルパス, エラーメッセージ) — 取得できた場合は(パス, None)、
		ダウンロードに失敗した場合は(None, 理由)。呼び出し側は理由を[元表]のNGメッセージに使う
		（取得できなくても、XPX定義ファイル側だけで判定できる○×シートのチェックは続行する）。
	"""
	location = SHAREPOINT_REFERENCE_FILE_PATH_OR_URL
	if not sp_integration.is_sharepoint_file_url(location):
		return location, None

	try:
		local_path = sp_integration.download_file_from_url(access_token, location, work_directory)
	except Exception as e:
		return None, (
			f"SharePointから元表ファイルを取得できない: {e}"
			f"（SHAREPOINT_REFERENCE_FILE_PATH_OR_URL={location}）"
		)

	print(f"元表ファイルをSharePointから取得: {os.path.basename(local_path)}")
	return local_path, None


def _sanitize_file_name_part(value):
	"""
	Title列の値を、ファイル名の一部として安全な形へ整える。`_FILE_NAME_NG_CHARS`に該当する文字は
	`_`へ置換し、空白（全角空白・改行を含む）は`_`へ潰し、前後の`.`／`_`／空白を除去して
	`_FILE_NAME_PART_MAX_LEN`文字で切る。

	戻り値:
		整えた文字列。使える文字が1つも残らない場合（Title列が空・記号のみ等）は空文字。
	"""
	text = "" if value is None else str(value)
	text = _FILE_NAME_NG_CHARS.sub("_", text)
	text = re.sub(r"\s+", "_", text)
	return text[:_FILE_NAME_PART_MAX_LEN].strip("._ 　")


def _check2_detail_file_name(case_id):
	"""
	②専用の詳細ファイル名`FSマトリクス_定義ファイルチェック結果_{Title列}.txt`を組み立てる。

	Title列が空、またはファイル名に使える文字が1つも残らない場合は接尾辞を付けず
	`DETAIL_FILE_NAME_2`のままにする（詳細ファイルは申請ごとに別の案件フォルダ配下へ
	アップロードするため、接尾辞が無くても他の申請の結果を上書きすることはない）。
	"""
	suffix = _sanitize_file_name_part(case_id)
	if not suffix:
		return DETAIL_FILE_NAME_2
	stem, extension = os.path.splitext(DETAIL_FILE_NAME_2)
	return f"{stem}_{suffix}{extension}"


def _write_check2_report(ok_lines, ng_lines, output_path):
	"""②専用の詳細ファイルをOK/NG2セクションで書き出す（②`input_check_main.write_report`と
	同じ構成）。SharePointへアップロードしたファイルをブラウザで直接開いた際の文字化けを避けるため、
	①の詳細ファイルと同じ`utf-8-sig`（BOM付きUTF-8）で出力する（②の元実装は`utf-8`のみだったが、
	①`result.write_result_file`の2026-09-01追補・文字化け対策と同じ理由でこちらも合わせた）。"""
	os.makedirs(os.path.dirname(output_path), exist_ok=True)
	timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
	with open(output_path, "w", encoding="utf-8-sig") as f:
		f.write(f"実行日時: {timestamp}\n\n")
		f.write("=== OK ===\n")
		for line in ok_lines:
			f.write(f"・{line}\n")
		f.write("\n=== NG ===\n")
		for line in ng_lines:
			f.write(f"・{line}\n")


def _run_check2(access_token, item, local_input_dir, fallback_dir, definition_parts, definition_result_message):
	"""
	②分（FSマトリクス↔定義ファイルチェック）を1アイテム分実行し、②専用の詳細ファイルを
	`local_input_dir`直下へ書き出す（ファイル名はTitle列付きの`_check2_detail_file_name`。
	申請ごとに別ファイルになるので、複数申請の結果を手元に並べても区別できる）。
	`definition_parts`／`definition_result_message`はアイテム非依存のため`main()`で1回だけ
	計算し、全アイテムで共有する。

	戻り値:
		(check2_line, overall_result, local_detail_path) — `InputFileCheckResult`列へ書き込む
		②分の1行、②全体のOK/NG、書き出した②詳細ファイルのローカルパス（呼び出し側が
		そのままSharePointへアップロードする。ファイル名を2箇所で組み立てて食い違わせないよう、
		パスは呼び出し側で再構築せずこの戻り値を使うこと）
	"""
	ok_lines, ng_lines = [], []
	local_detail_path = os.path.join(
		local_input_dir, _check2_detail_file_name(item.get(sp_integration.TITLE_FIELD_INTERNAL_NAME))
	)

	def record(label, result, message):
		line = f"{label}{message}"
		print(line)
		(ok_lines if result == "OK" else ng_lines).append(line)
		_write_check2_report(ok_lines, ng_lines, local_detail_path)

	for label, part_result, part_message in definition_parts:
		record(f"[{label}]", part_result, part_message)

	fs_matrix_urls = sp_integration.extract_all_file_urls(item.get(sp_integration.FS_MATRIX_FIELD_INTERNAL_NAME))
	if not fs_matrix_urls:
		matrix_result, matrix_message = "NG", f"{sp_integration.FS_MATRIX_FIELD_INTERNAL_NAME}からリンクを取得できない"
	else:
		fs_matrix_local_paths = [
			_resolve_local_path(local_input_dir, fallback_dir, access_token, url) for url in fs_matrix_urls
		]
		matrix_result, matrix_message = matrix_check.check_matrix_sheet_for_all_files(fs_matrix_local_paths)
	record("[FS Matrix]", matrix_result, matrix_message)

	command_support_urls = [
		url
		for url in sp_integration.extract_all_file_urls(item.get(sp_integration.COMMAND_SUPPORT_FIELD_INTERNAL_NAME))
		if gst_command_check.is_command_support_gst_file(os.path.basename(url))
	]
	command_support_local_paths = [
		_resolve_local_path(local_input_dir, fallback_dir, access_token, url) for url in command_support_urls
	]
	gst_result, gst_message = gst_command_check.check_gst_command_files(command_support_local_paths)
	if command_support_urls:
		record("[GSTCommand]", gst_result, gst_message)

	overall_result, overall_message = unified_result.combine_results(
		(matrix_result, matrix_message),
		(gst_result, gst_message),
		definition_result_message,
	)
	return unified_result.build_check2_summary_line(overall_result), overall_result, local_detail_path


def _run_check3(local_input_dir, extra_search_roots, output_path):
	"""③（HEX/A2Lチェック）を実行し、判定サマリ1行・判定結果・詳細ファイルパスを返す。

	引数:
		local_input_dir: ダウンロード済み`01_INPUT`のローカルパス
		extra_search_roots: 追加の探索ルート（①がzipを展開した一時ディレクトリ。
			`CaseScanResult.work_directory`を渡す。zipが無い案件ではNone）
		output_path: 詳細ファイルの出力先

	戻り値:
		(check3_line, check3_overall_result, output_path)

	注意:
		`check3_overall_result`は`"OK"`/`"NG"`の2値のみ（③は`確認不能`をNGへ畳む）。
		①の`対象外`のような「詳細ファイルを出さない」状態は無いため、呼び出し側は
		常に`detail_uploads`へ追加してよい。
	"""
	result, reason, detail_message = prof_check.run_check3(local_input_dir, extra_search_roots)

	# ①②の詳細ファイルと同じくBOM付きUTF-8で書く（BOM無しUTF-8はブラウザで直接開いた際に
	# Shift-JISと誤判定されるため。`check_mao\result.write_result_file`のコメント参照）
	os.makedirs(os.path.dirname(output_path), exist_ok=True)
	with open(output_path, "w", encoding="utf-8-sig") as f:
		f.write(detail_message)

	return unified_result.build_check3_summary_line(result, reason), result, output_path


def _process_one_item(access_token, entity_type, item, definition_parts, definition_result_message):
	"""
	1件のSubmittedアイテムを処理する（①チェーンは必ず実行、②チェーンは`RUN_CHECK2`時のみ）。
	①②の行をまとめて1回のMERGEで`InputFileCheckResult`列へ書き戻す。

	戻り値:
		(item_id, 案件ID, 判定サマリ文字列, 列に書き戻した最終値)
	"""
	item_id = item.get("Id")
	case_id = item.get(sp_integration.TITLE_FIELD_INTERNAL_NAME)

	file_urls = sp_integration.extract_all_file_urls(item.get(sp_integration.FILE_LINK_FIELD_INTERNAL_NAME))
	if not file_urls:
		message = f"{sp_integration.FILE_LINK_FIELD_INTERNAL_NAME}からリンクを取得できないため案件フォルダを特定できない"
		return item_id, case_id, "NG", message

	input_folder_server_relative = sp_integration.derive_input_folder_from_url(file_urls[0])
	if input_folder_server_relative is None:
		message = f"リンク先URLから{sp_integration.INPUT_FOLDER_NAME}フォルダを特定できない: {file_urls[0]}"
		return item_id, case_id, "NG", message

	work_directory = tempfile.mkdtemp(prefix="unified_check_")
	try:
		local_input_dir = os.path.join(work_directory, sp_integration.INPUT_FOLDER_NAME)
		sp_integration.download_folder_tree(access_token, input_folder_server_relative, local_input_dir)

		# ---- ①チェーン（case_scan → can_info/mao_can_id → compare → result） ----
		scan_result = scan_case_folder(local_input_dir)
		try:
			mask_check_results = []
			if scan_result.matrix_path is None:
				for mask_entry in scan_result.mask_entries:
					mask_check_results.append(
						MaskCheckResult(mask_entry.label, skipped_reason="CANマトリクスが特定できないため判定不可")
					)
			else:
				matrix_records = extract_matrix_can_id_records(
					scan_result.matrix_path,
					sheet_name=scan_result.matrix_sheet_name,
					header_row=scan_result.matrix_header_row,
					can_id_column=scan_result.matrix_can_id_column,
				)
				for mask_entry in scan_result.mask_entries:
					mao_records = extract_mao_can_id_records(mask_entry.mao_path)
					comparison_rows = compare_can_id_sources(mao_records, matrix_records)
					check_result = check1_result.build_check_result(comparison_rows)
					mask_check_results.append(MaskCheckResult(mask_entry.label, check_result=check_result))

			case_check_result = check1_result.build_case_result(mask_check_results, scan_result.structural_ng_messages)
			check1_line = check1_result.build_result_summary(case_check_result)

			own_lines = {check1_result.CHECK1_LINE_PREFIX: check1_line}
			judgment_parts = [f"①:{case_check_result.overall_judgment}"]
			detail_uploads = []

			if case_check_result.overall_judgment != "対象外":
				detail_message = check1_result.format_case_result_message(case_check_result)
				local_detail_path_1 = os.path.join(local_input_dir, DETAIL_FILE_NAME_1)
				check1_result.write_result_file(detail_message, local_detail_path_1)
				detail_uploads.append(local_detail_path_1)

			# ---- ②チェーン（matrix_check/gst_command_check/definition_file_check。RUN_CHECK2時のみ） ----
			if RUN_CHECK2:
				fallback_dir = os.path.join(work_directory, "_check2_fallback")
				check2_line, check2_overall_result, local_detail_path_2 = _run_check2(
					access_token, item, local_input_dir, fallback_dir, definition_parts, definition_result_message
				)
				own_lines[unified_result.CHECK2_LINE_PREFIX] = check2_line
				judgment_parts.append(f"②:{check2_overall_result}")
				detail_uploads.append(local_detail_path_2)

			# ---- ③チェーン（prof_scan → xpx_checks。RUN_CHECK3時のみ） ----
			# ①のzip展開先（scan_result.work_directory）を追加の探索ルートとして渡すことで、
			# ③はzip展開を自前で実装せず①の成果を再利用する。この展開先は下の
			# 内側finallyで削除されるため、③は必ずこのtryブロック内で実行する。
			if RUN_CHECK3:
				local_detail_path_3 = os.path.join(local_input_dir, DETAIL_FILE_NAME_3)
				check3_line, check3_overall_result, local_detail_path_3 = _run_check3(
					local_input_dir, scan_result.work_directory, local_detail_path_3
				)
				own_lines[unified_result.CHECK3_LINE_PREFIX] = check3_line
				judgment_parts.append(f"③:{check3_overall_result}")
				detail_uploads.append(local_detail_path_3)

			# 集約ファイル（①②③統合の判定サマリ）のURLは、固定名・overwrite=trueでアップロード
			# するため、アップロード前でも決定的に定まる。列書き戻し（ETagリトライループ）より
			# 先に組み立てておく。
			aggregate_server_relative = sp_integration.result_file_server_relative_url(
				input_folder_server_relative, AGGREGATE_FILE_NAME
			)
			log_html = sp_integration.build_html_link(aggregate_server_relative, AGGREGATE_FILE_NAME)

			# `InputFileCheckResult`列は①②③統合の判定サマリ列。①②の行をまとめて1回のMERGEで
			# 書き戻すため①↔②間の412は発生しないが、③・PMOの手動編集等、本ツール以外の書き手が
			# 割り込んだ場合に備えてETag楽観的並行制御のリトライは残す。
			merged = None
			for _ in range(_MERGE_RETRY_LIMIT):
				existing_value, etag = sp_integration.get_item_result_and_etag(access_token, item_id)
				merged = unified_result.merge_check_lines(existing_value, own_lines)
				try:
					sp_integration.write_result_columns(access_token, entity_type, item_id, merged, log_html, etag=etag)
					break
				except sp_integration.ConcurrentUpdateError:
					continue
			else:
				raise sp_integration.ConcurrentUpdateError(
					f"item {item_id}: {_MERGE_RETRY_LIMIT}回リトライしても列の書き戻しに成功しなかった"
				)

			for local_path in detail_uploads:
				sp_integration.upload_result_file(access_token, input_folder_server_relative, local_path)

			# 列に書き戻した最終的な値（＝真実の源）から集約ファイル本文を再生成し、固定名で
			# アップロードする（overwrite=true）。
			aggregate_content = check1_result.build_aggregate_file_content(merged)
			local_aggregate_path = os.path.join(local_input_dir, AGGREGATE_FILE_NAME)
			check1_result.write_result_file(aggregate_content, local_aggregate_path)
			sp_integration.upload_result_file(access_token, input_folder_server_relative, local_aggregate_path)

			return item_id, case_id, " ".join(judgment_parts), merged
		finally:
			if scan_result.work_directory is not None:
				shutil.rmtree(scan_result.work_directory, ignore_errors=True)
	finally:
		shutil.rmtree(work_directory, ignore_errors=True)


def _write_run_report(lines):
	with open(RUN_REPORT_PATH, "w", encoding="utf-8") as f:
		f.write("\n".join(lines))
		f.write("\n")


def run():
	"""Submitted全件を処理して終了する（単発実行）。"""
	access_token = sp_auth.get_access_token()
	items = sp_integration.get_submitted_items(access_token)

	report_lines = [f"=== ①②統合 SP駆動チェック 実行結果（対象{len(items)}件、RUN_CHECK2={RUN_CHECK2}） ==="]

	if not items:
		print("Submitted状態のアイテムはありませんでした。")
		_write_run_report(report_lines)
		return

	entity_type = sp_integration.get_entity_type_full_name(access_token)

	# ②の定義ファイルチェック（元表比較・○×表チェック）はアイテムに依存しないため、RUN_CHECK2時に
	# 1回だけ実行し、全アイテムで結果を共有する。元表ファイルもここで1回だけ取得する。
	definition_parts = None
	definition_result_message = ("OK", "")
	if RUN_CHECK2:
		reference_work_dir = tempfile.mkdtemp(prefix="unified_ref_")
		try:
			reference_path, reference_error = _resolve_definition_reference(access_token, reference_work_dir)
			# 取得に失敗しても、XPX定義ファイル側だけで判定できる[fcl_list_o_FI]／[fcl_list_x_FI]は
			# 通常どおり評価する。存在しないパスを渡すと[元表]は「ファイルを開けない」という
			# 分かりにくいメッセージでNGになるため、その1件だけ取得失敗の理由へ差し替える。
			definition_parts = definition_file_check.check_definition_file_parts(
				reference_path or os.path.join(reference_work_dir, "_未取得.xlsx")
			)
			if reference_error:
				print(reference_error)
				definition_parts = [
					(label, part_result,
						reference_error if label == definition_file_check.MOTOHYOU_SHEET_NAME else part_message)
					for label, part_result, part_message in definition_parts
				]
			definition_result_message = unified_result.combine_results(
				*[(part_result, part_message) for _, part_result, part_message in definition_parts]
			)
		finally:
			# 元表ファイルは`check_definition_file_parts`が読み終えた時点で不要になる。
			shutil.rmtree(reference_work_dir, ignore_errors=True)

	for item in items:
		try:
			item_id, case_id, judgment, message = _process_one_item(
				access_token, entity_type, item, definition_parts, definition_result_message
			)
			line = f"item {item_id}（{case_id}）: {judgment}\n{message}"
		except Exception as e:
			line = f"item {item.get('Id')}（{item.get(sp_integration.TITLE_FIELD_INTERNAL_NAME)}）: 例外 {e}"

		print(line)
		report_lines.append(line)
		_write_run_report(report_lines)

	print(f"実行結果レポート出力先: {RUN_REPORT_PATH}")


if __name__ == "__main__":
	if len(sys.argv) > 1 and sys.argv[1] == "--list-fields":
		sp_integration.list_fields(sp_auth.get_access_token())
	else:
		run()

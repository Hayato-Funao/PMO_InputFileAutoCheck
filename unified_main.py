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
import shutil
import sys
import tempfile

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

import check2_report  # noqa: E402  ②分。pure（SharePoint連携に依存しない）のため直接import
import definition_file_check  # noqa: E402
import gst_command_check  # noqa: E402
import matrix_check  # noqa: E402

import prof_check  # noqa: E402  ③分。pure（SharePoint連携に依存しない）のため直接import

import sp_auth  # noqa: E402  本フォルダ内のSP連携コード
import sp_integration  # noqa: E402
import unified_result  # noqa: E402

# ---- 設定 ------------------------------------------------------------------------------
# ②の定義ファイルチェックを実行対象に含めるかどうか（2026-09-02改修でTrueへ切り替え）。
# Trueにすると②の行が`InputFileCheckResult`列へ加わり、②詳細ファイル
# （`FSマトリクス_定義ファイルチェック結果.txt`）が案件の`01_INPUT`へアップロードされる。
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
RUN_CHECK3 = False

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

# ②専用の詳細結果（定義ファイル管理Excel／FI FSマトリクスファイル／GSTコマンド装備表の明細。
# ②`input_check_main.DETAIL_FILE_NAME`と同名）。中身の書式は「実行日時 → チェック対象ファイル
# 単位の3グループ（各行はOK/NGを文末で表す）」（2026-09-03改修。`check_fs_matrix.check2_report`が
# 組み立てる。従来の「実行日時 → === OK === → === NG ===」から変更した）。
#
# 【2026-09-03改修】2026-09-02に導入したTitle列（案件ID）付きのファイル名
# （`FSマトリクス_定義ファイルチェック結果_{Title列}.txt`。`_check2_detail_file_name`と
# ファイル名サニタイズ処理）を取消し、①③と同じ固定名へ戻した。詳細ファイルは案件ごとに
# 別の`01_INPUT`配下へアップロードするため、固定名でも他案件の結果を上書きすることはなく、
# ①（MAO_CANマトリクス比較結果.txt）③（HEX_A2Lチェック結果.txt）と名前の付け方が揃う
# （集約ファイルから「詳細は○○.txt」と案内する場合も、案件ごとに名前が変わらない方が扱いやすい）。
DETAIL_FILE_NAME_2 = "FSマトリクス_定義ファイルチェック結果.txt"

# ③専用の詳細結果（①②③の各判定と入力ファイル情報の明細）。書式は
# 「実行日時 → === OK === → === NG ===」（②は2026-09-03にグループ形式へ変更したが、③は
# 従来の2セクション形式のまま。`check_prof.prof_check._build_detail_message`が組み立てる）。
DETAIL_FILE_NAME_3 = "HEX_A2Lチェック結果.txt"

# ①②③統合の判定サマリ（`InputFileCheckResult`列のマージ後値）を書き込む集約ファイル。①②と同名。
#
# 【2026-09-03改修】判定サマリ（①②③のOK/NG）だけでは「何がNGなのか」が分からず、3つの詳細
# ファイルを個別に開く必要があったため、③が元からやっていた`NG（理由）`の形に①②も揃え、
# 各チェックの行末へNGの主因をカッコで併記するようにした
# （`unified_result.build_aggregate_file_content`が組み立てる）。ただし②の
# `定義ファイル管理Excel`グループのNGは主因の候補から除外する（PMO管理の固定ファイルに対する
# チェックで全案件同じ結果になり、案件担当者が対処できる情報ではないため。②の詳細ファイル
# `DETAIL_FILE_NAME_2`には従来どおり出る）。
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


def _write_check2_report(groups, output_path):
	"""②専用の詳細ファイルを、チェック対象ファイル単位の3グループで書き出す
	（書式は`check_fs_matrix.check2_report`側に集約。2026-09-03改修。従来の
	`=== OK ===`／`=== NG ===`2セクション構成から変更した）。SharePointへアップロードした
	ファイルをブラウザで直接開いた際の文字化けを避けるため、①の詳細ファイルと同じ
	`utf-8-sig`（BOM付きUTF-8）で出力する。"""
	check2_report.write_report(groups, output_path)


def _run_check2(
	access_token,
	item,
	local_input_dir,
	fallback_dir,
	definition_parts,
	definition_result_message,
	motohyou_index,
	motohyou_error,
):
	"""
	②分（FSマトリクス↔定義ファイルチェック）を1アイテム分実行し、②専用の詳細ファイルを
	`local_input_dir`直下へ`DETAIL_FILE_NAME_2`（固定名）で書き出す。①③と同じ固定名なので、
	案件ごとの`01_INPUT`配下では常に同じファイル名になる（2026-09-03改修。Title列を付ける
	方式は取消。`DETAIL_FILE_NAME_2`のコメント参照）。
	`definition_parts`／`definition_result_message`／`motohyou_index`／`motohyou_error`は
	アイテム非依存のため`run()`で1回だけ計算し、全アイテムで共有する。

	戻り値:
		(check2_line, overall_result, local_detail_path, ng_detail) — `InputFileCheckResult`列へ
		書き込む②分の1行、②全体のOK/NG、書き出した②詳細ファイルのローカルパス（呼び出し側が
		そのままSharePointへアップロードする。ファイル名を2箇所で組み立てて食い違わせないよう、
		パスは呼び出し側で再構築せずこの戻り値を使うこと）、集約ファイルの②の行末へ併記する
		NG内容（全グループのNGを全件", "で連ねた文字列。`定義ファイル管理Excel`グループは除外。
		NGが無ければ空文字。理由は`unified_result.build_aggregate_file_content`のdocstring参照）
	"""
	local_detail_path = os.path.join(local_input_dir, DETAIL_FILE_NAME_2)

	# グループ見出し → その下に並べる[(result, message), ...]。チェックが1件終わるたびに
	# 詳細ファイルを書き直すので、途中で落ちてもそこまでの結果はファイルに残る。
	groups = [
		(check2_report.GROUP_FS_MATRIX_DEFINITION, []),
		(check2_report.GROUP_FS_MATRIX, []),
		(check2_report.GROUP_GST_COMMAND, []),
	]
	group_lines = dict(groups)

	def record(group_name, result, message):
		print(f"[{group_name}]{message}")
		group_lines[group_name].append((result, message))
		_write_check2_report(groups, local_detail_path)

	fs_matrix_urls = sp_integration.extract_all_file_urls(item.get(sp_integration.FS_MATRIX_FIELD_INTERNAL_NAME))
	fs_matrix_local_paths = []
	if not fs_matrix_urls:
		# 【2026-09-04改修】`commentSystemMatrix`列が空欄（FSマトリクス未添付）の場合はNGにしない。
		# 以前は「リンクを取得できない。」でNGにしていたが、添付が無いこと自体を②のNGとして
		# 扱わない運用に変えた。判定に算入しないOK行として、チェック対象外である旨だけを
		# 詳細ファイルへ残す（②全体のOK/NGには影響せず、集約ファイルにも何も出ない）。
		#
		# 行ごと出さないサイレントスキップ（[GSTコマンド装備表]が添付なしのときの挙動）にはせず、
		# 1行残す方針にしている。FSマトリクスの未添付は稀なので、詳細ファイルに何も出ないと
		# 「チェックが動かなかったのか、添付が無かったのか」が後から判別できないため。
		matrix_parts = [
			(
				"OK",
				f"{sp_integration.FS_MATRIX_FIELD_INTERNAL_NAME}列が空欄のため、"
				f"FSマトリクスのチェックは対象外。",
			)
		]
	else:
		fs_matrix_local_paths = [
			_resolve_local_path(local_input_dir, fallback_dir, access_token, url) for url in fs_matrix_urls
		]
		matrix_parts = matrix_check.check_matrix_sheet_parts_for_all_files(fs_matrix_local_paths)
	for part_result, part_message in matrix_parts:
		record(check2_report.GROUP_FS_MATRIX, part_result, part_message)

	command_support_urls = [
		url
		for url in sp_integration.extract_all_file_urls(item.get(sp_integration.COMMAND_SUPPORT_FIELD_INTERNAL_NAME))
		if gst_command_check.is_command_support_gst_file(os.path.basename(url))
	]
	command_support_local_paths = [
		_resolve_local_path(local_input_dir, fallback_dir, access_token, url) for url in command_support_urls
	]
	# 添付が無ければ空リストが返る（サイレントスキップ。[GSTコマンド装備表]の見出しごと出ない）。
	gst_parts = gst_command_check.check_gst_command_files_parts(command_support_local_paths)
	for part_result, part_message in gst_parts:
		record(check2_report.GROUP_GST_COMMAND, part_result, part_message)

	# ---- [FSマトリクス↔定義ファイル管理Excel]グループ ----------------------------------------
	# 1. 提出FSマトリクス ↔ 定義ファイル管理の`元表`突合（2026-09-04新規）
	# 2. `fcl_list_o_FI`／`fcl_list_x_FI`の〇/×データ不備（`definition_parts`）
	# の2段構成で、1がNGなら2は実行しない（2026-09-04。下の打ち切り参照）。
	#
	# 元表突合は3つに分かれ、②の判定に算入するのは(a)(b)だけ（詳細は`definition_file_check`の
	# 見出しコメント参照）。FSマトリクスが1件も無ければいずれも空リストが返る（サイレントスキップ）。
	#
	#   (a) コード突合    : FSマトリクスのコードが`元表`に在るか → 無ければNG。判定に算入する。
	#   (b) 値の突合(HILS): 両方に在るコードの、HILSコード構成要素の食い違い → NG。判定に算入する。
	#   (c) 値の突合(その他): HILSコード外の列だけの食い違い → 参考情報。判定に算入しない。
	reconcile_code_parts = definition_file_check.check_fs_matrix_codes_in_motohyou(
		motohyou_index, motohyou_error, fs_matrix_local_paths
	)
	reconcile_data_parts, reconcile_data_info_parts = definition_file_check.check_fs_matrix_data_mismatches(
		motohyou_index, motohyou_error, fs_matrix_local_paths
	)
	for part_result, part_message in (
		reconcile_code_parts + reconcile_data_parts + reconcile_data_info_parts
	):
		record(check2_report.GROUP_FS_MATRIX_DEFINITION, part_result, part_message)

	# 【元表突合がNGなら定義ファイル管理Excelのチェックを打ち切る】2026-09-04。
	# 提出FSマトリクスと`元表`が食い違っている状態では、同じ定義ファイルの〇/×シートを
	# 検査しても意味のある結果にならない（まずFSマトリクス側を直す必要がある）ため、
	# `fcl_list_o_FI`／`fcl_list_x_FI`の行は出さず、②の判定にも算入しない。
	# 判定に算入する(a)(b)だけで打ち切りを判断する（(c)の参考情報では打ち切らない）。
	#
	# `definition_parts`自体は`run()`が実行ごとに1回だけ計算済みなので、ここでの打ち切りは
	# 「報告と判定に使わない」という意味（計算をやり直したり省いたりはしない）。
	reconcile_is_ng = any(
		part_result != "OK"
		for part_result, _message in (reconcile_code_parts + reconcile_data_parts)
	)
	if reconcile_is_ng:
		record(
			check2_report.GROUP_FS_MATRIX_DEFINITION,
			"OK",
			f"上記のとおりFSマトリクスと元表シートが一致しないため、"
			f"「{definition_file_check.OK_SHEET_NAME}」／「{definition_file_check.NG_SHEET_NAME}」の"
			f"チェックは実施しない。",
		)
	else:
		for _label, part_result, part_message in definition_parts:
			record(check2_report.GROUP_FS_MATRIX_DEFINITION, part_result, part_message)

	# 詳細ファイルは全チェック分を箇条書きで残すが、`InputFileCheckResult`列は1件しか
	# 持てないため、ここで②全体の1件へまとめる（1つでもNGなら最初のNGのmessage）。
	#
	# 元表突合のうち、②の判定に算入するのは(a)コード突合と(b)HILSコードに影響する値の不一致
	# （2026-09-04。どちらもXPXが引き当てられない状態を確定させるため）。
	# (c)HILSコード外だけの不一致（`reconcile_data_info_parts`）は参考情報なので**渡さない**
	# （DWGの改訂差だけで100件超になるのが常態で、算入すると全案件がNGになる）。
	# `definition_result_message`は打ち切り時に渡さない（上記の打ち切り参照）。
	overall_result, overall_message = unified_result.combine_results(
		*matrix_parts,
		*gst_parts,
		*reconcile_code_parts,
		*reconcile_data_parts,
		*([] if reconcile_is_ng else [definition_result_message]),
	)

	# 集約ファイルの②の行末へ併記するNG内容。①③が主因1件だけなのに対し、②は**全グループの
	# NGを全件**並べる（2026-09-03改修。以前は最初の1件だけを出していたため、
	# [FI FSマトリクスファイル]と[GSTコマンド装備表]の両方がNGの案件で後者が集約ファイルから
	# 消えていた）。前置きの文（「〜にデータ不備がある。」）は`extract_defect_detail`で落とし、
	# 不備箇所（`[FI FSmatrix / Z列] Z11 が ○/× ではありません: '※'`）だけを", "で連ねる。
	#
	# `check2_report.GROUPS_EXCLUDED_FROM_AGGREGATE`のグループ（`定義ファイル管理Excel`と
	# `FSマトリクス↔元表突合`。除外理由は同定数のコメント参照）は除外する。除外の結果NG内容が
	# 0件になることはあり得るが、`InputFileCheckResult`列の②行は上の`overall_result`のまま
	# NGを保つ（除外は集約ファイルの表示だけの話であり、判定は変えない）。`overall_message`を
	# 使わないのは、そちらが除外対象グループの分も含んでしまう上に最初の1件しか持たないため。
	ng_details = [
		check2_report.extract_defect_detail(message)
		for group_name, group_lines in groups
		if group_name not in check2_report.GROUPS_EXCLUDED_FROM_AGGREGATE
		for part_result, message in group_lines
		if part_result != "OK"
	]

	# `FSマトリクス↔元表突合`グループは上の除外対象だが、**判定に算入する2件のNGだけ**は
	# 集約ファイルにも出す（2026-09-04）。欠落コード・不一致の一覧は②の詳細ファイル側にだけ置き、
	# 集約ファイルへは「元表シートにない」「データが一致しない」（または「確認できない」）ことだけを
	# 短文で示す。グループ単位の除外では同グループ内の参考情報行（HILSコード外だけの不一致）まで
	# 出てしまうため、行単位でここで足す。
	for reconcile_summary in (
		definition_file_check.code_check_aggregate_summary(reconcile_code_parts),
		definition_file_check.data_check_aggregate_summary(reconcile_data_parts),
	):
		if reconcile_summary:
			ng_details.append(reconcile_summary)

	return (
		unified_result.build_check2_summary_line(overall_result),
		overall_result,
		local_detail_path,
		", ".join(ng_details),
	)


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

		`check3_line`は元から`NG（理由）`の形でNGの主因を含む（`build_check3_summary_line`）。
		集約ファイルで①②の行末へ主因を併記する処理（`unified_result.ng_reasons`）に③を
		渡さないのはこのため。
	"""
	result, reason, detail_message = prof_check.run_check3(local_input_dir, extra_search_roots)

	# ①②の詳細ファイルと同じくBOM付きUTF-8で書く（BOM無しUTF-8はブラウザで直接開いた際に
	# Shift-JISと誤判定されるため。`check_mao\result.write_result_file`のコメント参照）
	os.makedirs(os.path.dirname(output_path), exist_ok=True)
	with open(output_path, "w", encoding="utf-8-sig") as f:
		f.write(detail_message)

	return unified_result.build_check3_summary_line(result, reason), result, output_path


def _process_one_item(
	access_token,
	entity_type,
	item,
	definition_parts,
	definition_result_message,
	motohyou_index,
	motohyou_error,
):
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

			# 集約ファイル（`AGGREGATE_FILE_NAME`）で各チェックの行末へ併記するNGの主因
			# （2026-09-03新規。1チェック＝1行・主因1件のみ）。③は`check3_line`が元から
			# `NG（理由）`の形で主因を含むため渡さない。
			check1_ng_lines = check1_result.collect_ng_detail_lines(case_check_result)
			ng_reasons = {
				check1_result.CHECK1_LINE_PREFIX: check1_ng_lines[0] if check1_ng_lines else None
			}

			# 集約ファイルで各チェックの行末へ併記する詳細ファイルの場所（2026-09-04新規）。
			# 詳細ファイルを実際に出力したチェックだけを入れる（①が`対象外`の案件や
			# `RUN_CHECK2`／`RUN_CHECK3`がFalseのときは、そのチェックのキーを入れない）。
			# 場所は集約ファイル自身のリンク（`log_html`）と同じサーバー相対URLで表す。
			detail_paths = {}

			def _register_detail_path(line_prefix, detail_file_name):
				"""詳細ファイル名から`01_INPUT`配下の絶対URLを作り、併記対象へ登録する。

				場所の決め方は集約ファイル自身と同じ`result_file_server_relative_url`に揃える
				（固定名・overwrite=trueなのでアップロード前でも決定的に定まる）。そのうえで
				`build_absolute_url`でホストを付けた絶対URLにする（2026-09-04改修。サーバー相対
				URLのままだとコピーしてブラウザへ貼っても開けなかったため）。"""
				detail_paths[line_prefix] = sp_integration.build_absolute_url(
					sp_integration.result_file_server_relative_url(
						input_folder_server_relative, detail_file_name
					)
				)

			if case_check_result.overall_judgment != "対象外":
				detail_message = check1_result.format_case_result_message(case_check_result)
				local_detail_path_1 = os.path.join(local_input_dir, DETAIL_FILE_NAME_1)
				check1_result.write_result_file(detail_message, local_detail_path_1)
				detail_uploads.append(local_detail_path_1)
				_register_detail_path(check1_result.CHECK1_LINE_PREFIX, DETAIL_FILE_NAME_1)

			# ---- ②チェーン（matrix_check/gst_command_check/definition_file_check。RUN_CHECK2時のみ） ----
			if RUN_CHECK2:
				fallback_dir = os.path.join(work_directory, "_check2_fallback")
				check2_line, check2_overall_result, local_detail_path_2, check2_ng_detail = _run_check2(
					access_token,
					item,
					local_input_dir,
					fallback_dir,
					definition_parts,
					definition_result_message,
					motohyou_index,
					motohyou_error,
				)
				own_lines[unified_result.CHECK2_LINE_PREFIX] = check2_line
				judgment_parts.append(f"②:{check2_overall_result}")
				detail_uploads.append(local_detail_path_2)
				ng_reasons[unified_result.CHECK2_LINE_PREFIX] = check2_ng_detail
				_register_detail_path(unified_result.CHECK2_LINE_PREFIX, DETAIL_FILE_NAME_2)

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
				_register_detail_path(unified_result.CHECK3_LINE_PREFIX, DETAIL_FILE_NAME_3)

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

			# `RUN_CHECK2`／`RUN_CHECK3`がFalseでこの実行では走らなかったチェックでも、列に行が
			# 残っている場合（過去の実行で書かれた行が`merge_check_lines`で保持される）は、その
			# 詳細ファイルも過去の実行時に同じ`01_INPUT`へ固定名でアップロードされている。
			# 行があるのに場所だけ空欄になるのを避けるため、ここで補完する（2026-09-04追加。
			# `RUN_CHECK3=False`の運用で③の行だけ詳細の併記が無い状態になっていたため）。
			#
			# ①はフラグ制御が無く常に実行されるので対象外にしている（`対象外`判定の案件は詳細
			# ファイルを出さないため、上の明示登録に任せて補完してはいけない）。
			for line_prefix, detail_file_name in (
				(unified_result.CHECK2_LINE_PREFIX, DETAIL_FILE_NAME_2),
				(unified_result.CHECK3_LINE_PREFIX, DETAIL_FILE_NAME_3),
			):
				if line_prefix in detail_paths:
					continue
				if any(line.startswith(line_prefix) for line in (merged or "").splitlines()):
					_register_detail_path(line_prefix, detail_file_name)

			# 列に書き戻した最終的な値（＝真実の源）から集約ファイル本文を再生成し、固定名で
			# アップロードする（overwrite=true）。各チェックの行末にはNGの主因（2026-09-03改修。
			# ②の`定義ファイル管理Excel`分は`_run_check2`で除外済み）と、詳細ファイルの場所
			# （2026-09-04追加）を併記する。
			# 併記するのは集約ファイルだけで、列（`merged`）の値は従来どおり
			# `チェック①（...）：NG`のまま（PowerAppsの表示を変えないため）。
			aggregate_content = unified_result.build_aggregate_file_content(
				merged, ng_reasons, detail_paths
			)
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
	# 元表突合用のインデックス（2026-09-04新規）。`元表`はアイテム非依存なので、3MB超のxlsmを
	# 案件ごとに開かないよう実行ごとに1回だけ読む。読めなかった場合は理由(`motohyou_error`)を
	# そのまま各アイテムの[FSマトリクス↔元表突合]グループのNGメッセージへ出す。
	motohyou_index = None
	motohyou_error = None
	if RUN_CHECK2:
		# 【2026-09-04 元表比較を無効化】`元表`シートをSP上の元表ファイル
		# （`SHAREPOINT_REFERENCE_FILE_PATH_OR_URL`）の`FI FSmatrix`シートと突き合わせる条件を
		# コメントアウトしたため、その比較専用だったSP側元表ファイルのダウンロードも止めた
		# （読み手のいないファイルを実行ごとにRESTで取得しないため）。
		# `check_definition_file_parts`は[fcl_list_o_FI]／[fcl_list_x_FI]の2件だけを返す。
		# 再開する場合はこのブロックのコメントを外し、下の引数なし呼び出しを消して、
		# `definition_file_check.check_definition_file_parts`側のコメントも戻すこと。
		#
		# reference_work_dir = tempfile.mkdtemp(prefix="unified_ref_")
		# try:
		# 	reference_path, reference_error = _resolve_definition_reference(access_token, reference_work_dir)
		# 	# 取得に失敗しても、XPX定義ファイル側だけで判定できる[fcl_list_o_FI]／[fcl_list_x_FI]は
		# 	# 通常どおり評価する。存在しないパスを渡すと[元表]は「ファイルを開けない」という
		# 	# 分かりにくいメッセージでNGになるため、その1件だけ取得失敗の理由へ差し替える。
		# 	definition_parts = definition_file_check.check_definition_file_parts(
		# 		reference_path or os.path.join(reference_work_dir, "_未取得.xlsx")
		# 	)
		# 	if reference_error:
		# 		print(reference_error)
		# 		# 結果ログの文体を崩さないよう、取得失敗の理由も[元表]行の文へ埋め込む。
		# 		motohyou_unavailable = definition_file_check.motohyou_unavailable_message(reference_error)
		# 		definition_parts = [
		# 			(label, part_result,
		# 				motohyou_unavailable if label == definition_file_check.MOTOHYOU_SHEET_NAME else part_message)
		# 			for label, part_result, part_message in definition_parts
		# 		]
		# 	definition_result_message = unified_result.combine_results(
		# 		*[(part_result, part_message) for _, part_result, part_message in definition_parts]
		# 	)
		# finally:
		# 	# 元表ファイルは`check_definition_file_parts`が読み終えた時点で不要になる。
		# 	shutil.rmtree(reference_work_dir, ignore_errors=True)
		definition_parts = definition_file_check.check_definition_file_parts()
		definition_result_message = unified_result.combine_results(
			*[(part_result, part_message) for _, part_result, part_message in definition_parts]
		)

		# 元表突合用のインデックスも同じくアイテム非依存なので、ここで1回だけ読む。
		# 読めない場合も例外にせず理由を持ち回る（突合だけがNGになり、②の他のチェックは続行する）。
		motohyou_index, motohyou_error = definition_file_check.load_motohyou_index()
		if motohyou_error:
			print(f"[{check2_report.GROUP_FS_MATRIX_DEFINITION}]元表を読めないため突合できない: {motohyou_error}")
		else:
			print(
				f"[{check2_report.GROUP_FS_MATRIX_DEFINITION}]元表を読み込み: "
				f"{motohyou_index['name']}（担当区{definition_file_check.RECONCILE_TARGET_DEPT}の"
				f"突合対象{motohyou_index['row_count']}行）"
			)

	for item in items:
		try:
			item_id, case_id, judgment, message = _process_one_item(
				access_token,
				entity_type,
				item,
				definition_parts,
				definition_result_message,
				motohyou_index,
				motohyou_error,
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

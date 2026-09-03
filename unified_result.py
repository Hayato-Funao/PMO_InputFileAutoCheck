"""統合実行フォルダ専用の小さなヘルパー集（2026-09-02新規・方式B）。

②(check_fs_matrix)の`input_check_main.py`にある②専用ヘルパー（`CHECK2_LINE_PREFIX`／
`CHECK2_LABEL`／`combine_results`／`build_check2_summary_line`）は関数単体はpureだが、
`input_check_main`モジュール全体がmsal・ntlm_proxyをimportするため、本フォルダから
`import input_check_main`はしない（②のSP連携コードを二重に持ち込んでしまう）。代わりに、
これらの数行のロジックだけをここに複製する。

①(check_mao)の`result.CHECK1_LINE_PREFIX`／`sort_key`は`result.py`自体がpure（SharePoint連携に
依存しない）なので、そちらは複製せず`unified_main.py`が`sys.path`経由で直接importして流用する。
`merge_check_lines`はその`sort_key`を実行時にimportして使う（本モジュール単体でも読み込めるよう、
モジュール先頭ではなく呼び出し時にimportする）。

【注意】`merge_check_lines`の`from result import sort_key`は、①(check_mao)フォルダが`sys.path`に
追加済みであることが前提（`unified_main.py`のトップレベルで`sys.path.append(CHECK_MAO_DIR)`済み）。
`unified_main`をimportせずに本モジュールだけを先に使うと`ModuleNotFoundError: No module named 'result'`
になる。
"""

CHECK2_LINE_PREFIX = "チェック②"

# `PMOチェック自動化マニュアル.pptx`8ページの表示イメージに合わせた文言（②`input_check_main.CHECK2_LABEL`
# と同一）。
CHECK2_LABEL = f"{CHECK2_LINE_PREFIX}（FSマトリクス↔定義ファイルチェック）"


def combine_results(*checks):
	"""checksは(result, message)タプルの列。全てOKならOK、1つでもNGなら最初のNGのmessageを返す
	（②`input_check_main.combine_results`と同じロジック）。"""
	for check_result, check_message in checks:
		if check_result != "OK":
			return "NG", check_message
	return "OK", " / ".join(message for _, message in checks if message)


def build_check2_summary_line(result):
	"""SharePoint列`InputFileCheckResult`へ書き戻す②分の判定サマリ1行を組み立てる
	（②`input_check_main.build_check2_summary_line`と同じ）。"""
	return f"{CHECK2_LABEL}：{result}"


CHECK3_LINE_PREFIX = "チェック③"

# 資料スライド5の表題に合わせた文言。スライド9には旧仕様の「ブート領域チェック」と
# 記載されているが、ブート領域チェックは撤回済み(HEXがブート領域のレコードを持つことは
# 全ての正常HEXに当てはまり判別能力が無いため)。スライド9側の修正が別途必要。
CHECK3_LABEL = f"{CHECK3_LINE_PREFIX}（HEX/A2Lチェック）"


def build_check3_summary_line(result, reason=None):
	"""SharePoint列`InputFileCheckResult`へ書き戻す③分の判定サマリ1行を組み立てる。

	`result`は`"OK"`/`"NG"`の2値のみ(③は`確認不能`を`NG`へ畳んでから渡す。
	PowerAppsが列値の`"NG"`部分文字列で赤/緑を判定するため、`確認不能`のままでは
	緑になってしまう)。`reason`を渡すと`NG（理由）`の形で理由を併記する。
	"""
	if reason:
		return f"{CHECK3_LABEL}：{result}（{reason}）"
	return f"{CHECK3_LABEL}：{result}"


def merge_check_lines(existing_value, own_lines):
	"""
	`InputFileCheckResult`列の既存値(`existing_value`)に、`own_lines`（{行プレフィックス: 新しい行}
	の辞書。本ツールは①②を同一プロセスで一括処理するため、通常は
	`{CHECK1_LINE_PREFIX: ..., CHECK2_LINE_PREFIX: ...}`の2件、`RUN_CHECK2=False`時は①分の
	1件のみを渡す）をマージし、①(check_mao)の`result.sort_key`と同じ基準で①②③…の番号順に
	安定ソートして返す。

	既存値の行のうち、`own_lines`のいずれかのプレフィックスで始まる行は新しい行に置き換え、
	それ以外（③以降・PMOによる手動編集等、本ツール以外の書き手による行）はそのまま保持する。
	①②を別プロセスで個別にマージする方式A（①`sp_client.merge_check1_line`／
	②`input_check_main.merge_check2_line`）と異なり、本関数は複数行を一括で差し替えられる
	（アイテムごとに1回のMERGEで①②行を同時更新するため）。

	引数:
		existing_value: 列の現在値（複数行、None/空文字の場合あり）
		own_lines: {行プレフィックス: 新しい1行}の辞書

	戻り値:
		マージ後の全体の値（①②③…の番号順にソートされた改行区切りの複数行文字列）
	"""
	from result import sort_key  # ①(check_mao)のsort_keyをそのまま流用（pure・複製しない）

	existing_lines = (existing_value or "").splitlines()
	prefixes = tuple(own_lines.keys())
	other_lines = [line for line in existing_lines if not line.startswith(prefixes)]
	all_lines = other_lines + list(own_lines.values())
	return "\n".join(sorted(all_lines, key=sort_key))

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


# 集約ファイルで詳細ファイルの場所を併記するときの見出し語（2026-09-04新規）。
AGGREGATE_DETAIL_LABEL = "詳細"


def build_aggregate_file_content(summary_value, ng_reasons=None, detail_paths=None):
	"""
	集約ファイル（`インプットファイルチェック結果.txt`）の本文を組み立てる（2026-09-03新規）。

	従来は①の`result.build_aggregate_file_content`（見出し＋`InputFileCheckResult`列の値を
	そのまま）を使っており、①②③のOK/NGしか分からず「何がNGなのか」は3つの詳細ファイルを
	個別に開かないと読めなかった。本関数は各チェックの行末へNGの主因と詳細ファイルの場所を
	カッコで併記し、集約ファイルだけで「どのチェックがなぜNGで、続きはどこを見るのか」が
	読めるようにする。

	出力例:

		【チェック結果】
		チェック①（MAO↔CANマトリクスチェック）：NG（CAN ID: 112h（10進:274） / MAOラベル: F_FCAN112OK / 欠けている出典: CANマトリクス）（詳細: /sites/jphgt105596/.../01_INPUT/MAO_CANマトリクス比較結果.txt）
		チェック②（FSマトリクス↔定義ファイルチェック）：OK（詳細: /sites/jphgt105596/.../01_INPUT/FSマトリクス_定義ファイルチェック結果.txt）
		チェック③（HEX/A2Lチェック）：NG（A2L↔epd 先頭アドレス一致を確認できず）（詳細: /sites/jphgt105596/.../01_INPUT/HEX_A2Lチェック結果.txt）

	【詳細ファイルの場所を併記する】2026-09-04追加。①②③の詳細ファイルは案件ごとの
	`01_INPUT`直下（集約ファイル自身と同じフォルダ）へアップロードされる。どのチェックの続きが
	どのファイルなのかを集約ファイルだけで辿れるよう、OK/NGに関わらず行末へ併記する
	（NGの主因がある場合は「主因のカッコ」→「詳細のカッコ」の順で2つ並ぶ）。
	詳細ファイルを出さないチェック（①が`対象外`の案件、`RUN_CHECK2`／`RUN_CHECK3`がFalseの
	とき、PMOの手動編集行など）には併記しない。呼び出し側が`detail_paths`へそのチェックの
	キーを渡さなければよい。

	【1チェック＝1行】NGが複数件あっても行は増やさない。行末のカッコに何件入れるかは
	チェックごとに違う（2026-09-03決定）:
		- ①: 主因1件のみ。NGのCAN IDが数十件になる案件があり、全件を併記すると1行が
		  読めない長さになるため（優先順は`result.collect_ng_detail_lines`の並び順）。
		- ②: 全グループのNGを全件（呼び出し側が", "で連結して渡す）。[FI FSマトリクス
		  ファイル]と[GSTコマンド装備表]の両方がNGの案件で片方が消えないようにするため。
		- ③: `build_check3_summary_line`が元から`NG（理由）`の形で主因1件を持つため、
		  ③分は`ng_reasons`へ渡さない（列の値のまま出る）。
	いずれの場合も全件は従来どおり各チェックの詳細ファイル（`MAO_CANマトリクス比較結果.txt`／
	`FSマトリクス_定義ファイルチェック結果.txt`／`HEX_A2Lチェック結果.txt`）で確認できる。

	【②の`定義ファイル管理Excel`を除く運用】②の3グループのうち`定義ファイル管理Excel`
	（元表比較・fcl_list_o_FI／fcl_list_x_FIのデータ不備）は、案件のインプットではなく
	PMO側が管理する固定ファイルに対するチェックであり、全案件で同じ結果になる。案件担当者が
	集約ファイルを見て対処できる情報ではないため、併記の対象から外す（除外は呼び出し側
	（`unified_main._run_check2`）でグループ単位に行う。②の詳細ファイルには従来どおり出る）。
	その結果②のNG内容が0件になることはあり得るが、その場合は列の値のまま`：NG`だけを出す
	（判定は変えない）。

	引数:
		summary_value: `InputFileCheckResult`列のマージ後の値（`merge_check_lines`の戻り値）
		ng_reasons: {行プレフィックス: 併記するNG内容の文字列}。`merge_check_lines`の
			`own_lines`と同じキー（`result.CHECK1_LINE_PREFIX`／`CHECK2_LINE_PREFIX`）で
			渡す。値が空（NG無し・除外で何も残らない）のチェックは何も併記しない。
		detail_paths: {行プレフィックス: 詳細ファイルの場所}。同じキーで渡す。OK/NGに
			関わらず併記する。渡さなかったチェックには何も併記しない。

	戻り値:
		集約ファイルへ書き出す本文（末尾に改行は付けない。`result.write_result_file`が付ける）
	"""
	from result import build_aggregate_file_content as build_summary_section  # ①の見出し付けを流用

	reasons = {prefix: reason for prefix, reason in (ng_reasons or {}).items() if reason}
	paths = {prefix: path for prefix, path in (detail_paths or {}).items() if path}
	if not reasons and not paths:
		return build_summary_section(summary_value)

	lines = []
	for line in (summary_value or "").splitlines():
		# 1行1チェックを崩さないよう、併記する文字列中の改行は空白へ潰す
		# （②`check2_report.build_body`と同じ扱い）。
		for prefix, reason in reasons.items():
			if line.startswith(prefix):
				line = f"{line}（{' '.join(str(reason).splitlines())}）"
				break
		for prefix, path in paths.items():
			if line.startswith(prefix):
				line = f"{line}（{AGGREGATE_DETAIL_LABEL}: {' '.join(str(path).splitlines())}）"
				break
		lines.append(line)

	return build_summary_section("\n".join(lines))


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

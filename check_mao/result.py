"""MAO・CANテーブルの突合結果から、内部仕様書4.5節・5節・6節に定義されたチェック結果
（全体判定OK/NG＋NG明細メッセージ）を組み立てる結果出力モジュール。

判定（ComparisonRow.judgment、OK／NG／－の3値）のうち、**CAN ID単位の明細としてはNGのみ**を
外部出力対象とする。OK・「－」の個々のCAN IDは明細には出さない（内部仕様書4.5節・5節
「データ構造設計」）。ただし**件数（総数・OK・NG・－それぞれの件数）は集計値として出力する**
（2026-08-31追補。動作確認時に、全件内訳を出す開発用ツール`テスト\run_sample.py`と、NG明細のみを
出す本番`main.py`とで表示粒度が違うために結果が異なるように見える、という声を受けて追加した）。

2026-08-31追補で以下を追加した:
	- 案件全体判定を"OK"／"NG"の2値から"OK"／"NG"／"対象外"の3値へ拡張した。照合対象のMAOファイルが
	  1つも見つからない案件（枝番案件・空マスク・zip納品でMAO0件等）を、NGでもOKでもなく「対象外」
	  として扱う（`build_case_result`）。

2026-08-31改修で以下2点を追加した:
	- CANマトリクスをチェック対象から除外したため、NG明細の「欠けている出典」は常に
	  「CANテーブル」のみになった。
	- 複数マスク案件（`case_scan.py`が特定したマスクフォルダ単位）を1つの結果ファイルへ
	  集約する機能（MaskCheckResult／CaseCheckResult／build_case_result等）を追加した。
	  あわせて、同種ファイル（MAOまたはCANテーブル）が同一フォルダに複数存在する等の
	  構造的な問題（`case_scan.py`が検出）も、CAN ID突合とは別に「構造的NG」として
	  結果ファイルに記載する。マスクが1件のみの案件（`01_HEX関連`直下に`.mao`が1個だけ
	  置かれる代表マスク構成）では、マスク見出し（`=== マスク: ... ===`）を出さず、
	  件数・判定をそのまま案件レベルの結果として示す（2026-08-31追補）。

結果ファイル（人間可読テキスト）は、案件フォルダの`01_INPUT`直下へ本ツールが直接出力する
（「3.7」参照）。

【2026-09-01追補】SharePoint連携（SP駆動・単発バッチ化）を導入した。`sp_main.py`が
案件フォルダをSharePointから取得しチェックした場合、この結果ファイルを同じ`01_INPUT`直下へ
アップロードし（`sp_client.upload_result_file`）、そのURLをHTML形式（`<p><a href="...">...</a></p>`）
でSharePointリスト`XPX検証案件リスト2`の列`InputFileCheckResultLog`へ登録する
（`sp_client.build_html_link`／`write_result_columns`）。判定サマリ（OK/NGの色分け表示用の
短い文字列）は`build_result_summary`が組み立て、列`InputFileCheckResult`へ登録する。
PowerApps側はHTMLリンクの表示に合わせてボタン改修が必要（結果未確定時は押せない状態にする。
「PowerApps改修コード」のテキストファイル参照）。

この2列（`InputFileCheckResult`／`InputFileCheckResultLog`）の**実在・内部名（大文字小文字を
含む）は未確認**であり、`sp_client.py`の定数は暫定値である。実運用前に`sp_client.list_fields()`
で確認し、必要なら差し替えること。
"""

import re
from pathlib import Path

from normalize import format_can_id_hex


class NgDetail:
	"""NG判定となった1つのCAN IDについて、結果メッセージに記載する情報をまとめたレコード。"""

	def __init__(self, can_id_decimal, can_id_hex, mao_labels, target_ecus, missing_sources):
		self.can_id_decimal = can_id_decimal
		self.can_id_hex = can_id_hex
		self.mao_labels = mao_labels
		self.target_ecus = target_ecus  # CANテーブル側から判明したECU名（不明な場合あり）
		self.missing_sources = missing_sources  # 欠けている出典（現状は常に["CANテーブル"]）


class CheckResult:
	"""
	1マスク分のチェック結果（外部出力対象）。全体判定・NG明細のほか、CAN ID件数の集計値
	（総数・OK・－。NG件数は`len(ng_details)`で得られるため別途保持しない）を保持する。
	"""

	def __init__(self, overall_judgment, ng_details, total_count=0, ok_count=0, dash_count=0):
		self.overall_judgment = overall_judgment  # "OK" または "NG"
		self.ng_details = ng_details  # NgDetailのリスト（NGが無ければ空リスト）
		self.total_count = total_count  # 突合結果の総CAN ID数
		self.ok_count = ok_count  # 判定"OK"の件数
		self.dash_count = dash_count  # 判定"－"（判定対象外）の件数


class MaskCheckResult:
	"""1マスクの結果（比較できた場合はCheckResult、比較できなかった場合はその理由）を保持する。"""

	def __init__(self, mask_label, check_result=None, skipped_reason=None):
		self.mask_label = mask_label
		self.check_result = check_result
		self.skipped_reason = skipped_reason  # CANテーブルが特定できない等、比較を行えなかった理由


class CaseCheckResult:
	"""案件フォルダ全体（複数マスク＋構造的NG）を集約したチェック結果。"""

	def __init__(self, overall_judgment, mask_results, structural_ng_messages):
		self.overall_judgment = overall_judgment  # "OK"・"NG"・"対象外"の3値（2026-08-31追補で3値化）
		self.mask_results = mask_results  # MaskCheckResultのリスト
		self.structural_ng_messages = structural_ng_messages  # 構造的NGの説明文リスト（無ければ空）


def _build_target_ecu_display(comparison_row):
	"""
	対象ECUの表示用文字列を返す。

	【2026-09-01改修】チェック対象をCANテーブルからCANマトリクスへ変更した。CANマトリクスには
	Transmitter/Receivers相当の列（ECU別Tx/Rxノード列）が存在するが、列位置が機種ごとに異なる
	ため本ツールは読み取らない方針である（`can_info.py`のモジュールdocstring参照）。そのため
	対象ECUは特定できず、常にこの固定文字列を返す（旧CANテーブル時代はTransmitter/Receivers列
	から集めていたが、その実装は使われなくなった）。
	"""
	return "（CANマトリクス側にECU情報の記載なし。列位置が機種依存のため本ツールは読み取らない）"


def build_check_result(comparison_rows):
	"""
	突合結果(ComparisonRowのリスト)に内部仕様書4.5節の3値判定を適用し、
	外部出力対象のチェック結果（全体判定＋NG明細）を組み立てる（1マスク分）。

	引数:
		comparison_rows: compare.compare_can_id_sourcesが返すComparisonRowのリスト

	戻り値:
		CheckResult。NGが1件でもあれば全体判定は"OK"ではなく"NG"、無ければ"OK"
		（内部仕様書4.5節。「－」の行は全体判定に影響しない）。総数・OK件数・－件数の
		集計値も保持する（CAN ID単位の明細としてはNGのみ出力するが、件数は3値とも出力する）。
	"""
	ng_details = []
	ok_count = 0
	dash_count = 0
	for comparison_row in comparison_rows:
		if comparison_row.judgment == "OK":
			ok_count += 1
			continue
		if comparison_row.judgment == "－":
			dash_count += 1
			continue

		# judgment == "NG"。OK・「－」は個々のCAN IDをNG明細には含めない（件数のみ上で集計済み）
		mao_labels = ";".join(record.label_name for record in comparison_row.mao_records)
		ng_details.append(
			NgDetail(
				comparison_row.can_id_decimal,
				format_can_id_hex(comparison_row.can_id_decimal),
				mao_labels,
				_build_target_ecu_display(comparison_row),
				["CANマトリクス"],
			)
		)

	overall_judgment = "NG" if ng_details else "OK"
	return CheckResult(overall_judgment, ng_details, len(comparison_rows), ok_count, dash_count)


def build_case_result(mask_check_results, structural_ng_messages):
	"""
	マスク単位のMaskCheckResultの一覧と構造的NGメッセージから、案件全体のCaseCheckResultを組み立てる。

	`mask_check_results`が1件も無い場合（照合対象のMAOファイルが1つも見つからなかった場合）は
	"対象外"とする（2026-08-31追補）。本番想定データ調査で、1案件に複数の枝番が並存し代表となる
	枝番のみが`01_INPUT`を持つ構成（従属する枝番は`01_HEX関連`・`02_CAN関連`が丸ごと無い）や、
	空のマスクフォルダのみでMAOが1つも無い案件が多数確認された。これらはインプット不備（NG）でも
	照合完了（OK）でもなく、照合すべきインプットがそもそも存在しない「対象外」として扱う。
	この場合、`structural_ng_messages`（例:「01_HEX関連フォルダが見つからない」）は対象外の理由
	として保持するが、全体判定をNGにはしない。

	マスクが1件以上ある場合は従来どおり、いずれかのマスクがNG、または構造的NGが1件以上あれば
	"NG"、それ以外は"OK"とする（比較できなかったマスク（skipped_reasonあり。CANテーブル未特定等）は、
	それ自体を構造的NGとして扱う想定であり、ここでは全体判定に加算しない。呼び出し側で
	skipped_reasonの有無を構造的NGへ含めること）。
	"""
	if not mask_check_results:
		return CaseCheckResult("対象外", mask_check_results, structural_ng_messages)

	any_mask_ng = any(
		result.check_result is not None and result.check_result.overall_judgment == "NG"
		for result in mask_check_results
	)
	overall_judgment = "NG" if any_mask_ng or structural_ng_messages else "OK"
	return CaseCheckResult(overall_judgment, mask_check_results, structural_ng_messages)


CHECK1_LINE_PREFIX = "チェック①"

# `PMOチェック自動化マニュアル.pptx`8ページの表示イメージ（吹き出し）に記載された文言形式に
# 合わせる。①の名称は実装の対象（2026-09-01改修でCANマトリクスへ変更済み）に合わせて
# 「MAO↔CANマトリクスチェック」とする（マニュアル記載の旧「CANテーブル」表記は踏襲しない）。
CHECK1_LABEL = f"{CHECK1_LINE_PREFIX}（MAO↔CANマトリクスチェック）"

# `InputFileCheckResult`列内の「チェック①」〜「チェック⑳」で始まる行を検出するパターン
# （2026-09-02新規、②(check_fs_matrix)とのマージに伴う表示順固定化のため）。
_CHECK_LINE_PATTERN = re.compile(r"^チェック([①-⑳])")


def sort_key(line):
	"""
	`InputFileCheckResult`列の1行（例:"チェック①（...）：OK"）から、①②③…の順序を決める
	ソートキーを返す（2026-09-02新規）。

	`PMOチェック自動化マニュアル.pptx`8ページは①②③を番号順に固定表示する前提だが、
	`sp_client.merge_check1_line`（および②側の同等関数）は「既存値を読んで自分の行だけ
	差し替える」実装のため、素朴に連結すると表示順が実行順（後から書いたチェックの行が
	先頭に来る）に依存してしまう。本関数を`sorted(..., key=sort_key)`に渡すことで、
	実行順に関わらず常に①②③…の番号順になるようにする。

	丸数字①〜⑳はUnicodeで連続したコードポイント（U+2460〜）のため、④⑤…のチェックが
	将来追加されても本関数は無改修で対応できる。プレフィックスが一致しない想定外の行は
	末尾へ回す（Pythonの`sorted`は安定ソートのため、想定外行同士の相対順は入力順のまま
	保持される）。
	"""
	matched = _CHECK_LINE_PATTERN.match(line)
	if matched:
		return (0, ord(matched.group(1)) - ord("①") + 1)
	return (1, 0)


def build_result_summary(case_check_result):
	"""
	SharePoint列`InputFileCheckResult`へ書き戻す①分の判定サマリ1行を組み立てる
	（2026-09-01新規。2026-09-01追補で`PMOチェック自動化マニュアル.pptx`8ページの文言形式へ変更）。

	マニュアルの表示イメージは「チェック①（...）：NG／チェック②（...）：OK／チェック③（...）：NG」
	という①②③統合・複数行の形式だが、本関数が返すのは①分の1行のみである。①②③を1つの列へ
	統合する処理（他の行を保持したまま①行だけを差し替える）は`merge_check1_line`が担う
	（`sp_client.py`。呼び出し側は`sp_main.py`）。PowerApps側は本申請タブでこの文字列に"NG"が
	含まれるかどうかで色分け表示する想定（`PowerApps改修_チェック結果表示ボタン.txt`参照。複数行に
	なっても"NG"の部分一致判定はそのまま機能する）。
	"""
	if case_check_result.overall_judgment == "対象外":
		return f"{CHECK1_LABEL}：対象外"

	if case_check_result.overall_judgment == "OK":
		return f"{CHECK1_LABEL}：OK"

	return f"{CHECK1_LABEL}：NG"


def build_aggregate_file_content(summary_value):
	"""
	`inputFileCheckResult`列（①②③統合の判定サマリ。①分マージ後の値）から、SPへアップロードする
	集約結果ファイル（`sp_main.AGGREGATE_FILE_NAME`＝`インプットファイルチェック結果.txt`）の本文を
	組み立てる（2026-09-01追補8）。

	マニュアル(`PMOチェック自動化マニュアル.pptx`8ページ)の表示イメージに合わせ、先頭に見出しを付す。
	列側の値自体には見出しを付けない（PowerApps側の"NG"部分一致による色分け判定をシンプルに保つため。
	`build_result_summary`docstring参照）。
	"""
	return "【チェック結果】\n" + (summary_value or "")


def collect_ng_detail_lines(case_check_result):
	"""
	案件全体の①結果から、NGに寄与する行だけを重い順に取り出す（2026-09-03新規）。

	`format_case_result_message`（①専用の詳細ファイル本文）は件数・OK情報も含む全量出力だが、
	集約ファイル（`インプットファイルチェック結果.txt`）はチェック①の行末へNGの主因を
	カッコで併記する形にしたいため、本関数はNGに寄与する行のみを次の順で返す:
		- 構造的NG（同種ファイルの重複・ファイル未特定等）
		- 比較を行えなかったマスク（`skipped_reason`あり）
		- NG判定となったCAN IDの明細
	マスクが2件以上ある案件では、どのマスクの行かが分かるようマスクラベルを頭に付ける
	（1件だけの案件では`format_case_result_message`と同様に省く）。

	集約ファイルが使うのは先頭1件（主因）だけだが、どれを主因とするかの優先順は上記の
	並び順で表す（構造的NGがあればそれを最優先で見せる）。呼び出し側
	（`unified_main._process_one_item`）が`[0]`を取る。

	全体判定が"対象外"（照合対象のMAOファイルが1つも無い）の場合は空リストを返す。
	`structural_ng_messages`は「対象外の理由」として保持されているだけでNGではないため
	（`build_case_result`のdocstring参照）、NGの主因としては出さない。

	戻り値:
		NG明細行の文字列リスト（NGが無ければ空リスト）
	"""
	if case_check_result.overall_judgment == "対象外":
		return []

	lines = []
	for message in case_check_result.structural_ng_messages:
		lines.append(f"構造的NG: {message}")

	show_mask_label = len(case_check_result.mask_results) > 1
	for mask_result in case_check_result.mask_results:
		prefix = f"マスク{mask_result.mask_label}: " if show_mask_label else ""
		if mask_result.skipped_reason is not None:
			lines.append(f"{prefix}判定不可（{mask_result.skipped_reason}）")
			continue
		for detail in mask_result.check_result.ng_details:
			missing = "・".join(detail.missing_sources)
			# 対象ECUは常に固定の「記載なし」文（`_build_target_ecu_display`）のため、集約
			# ファイルでは省く（詳細ファイル側には従来どおり出す）。
			lines.append(
				f"{prefix}CAN ID: {detail.can_id_hex}（10進:{detail.can_id_decimal}） "
				f"/ MAOラベル: {detail.mao_labels or '(不明)'} "
				f"/ 欠けている出典: {missing}"
			)
	return lines


def format_result_message(check_result):
	"""
	内部仕様書6節のとおり、CAN ID単位の明細はNG時のみ記載した日本語メッセージを組み立てる
	（1マスク分。SharePoint経由でPowerAppsログに表示する内容に相当）。判定別の件数（総数・
	OK・NG・－）は、OK/NGいずれの場合も表示する（2026-08-31追補。個々のOK・「－」のCAN IDは
	列挙しない）。
	"""
	count_line = (
		f"CAN ID件数: 総数{check_result.total_count}件 / OK {check_result.ok_count}件"
		f" / NG {len(check_result.ng_details)}件 / － {check_result.dash_count}件"
	)

	if check_result.overall_judgment == "OK":
		return "\n".join(["結果: OK（NGのCAN IDはありません）", count_line])

	lines = [
		f"結果: NG（NG件数: {len(check_result.ng_details)}件）",
		count_line,
		"--- NG明細（要問合せ） ---",
	]
	for detail in check_result.ng_details:
		missing = "・".join(detail.missing_sources)
		lines.append(
			f"CAN ID: {detail.can_id_hex}（10進:{detail.can_id_decimal}） "
			f"/ MAOラベル: {detail.mao_labels or '(不明)'} "
			f"/ 対象ECU: {detail.target_ecus} "
			f"/ 欠けている出典: {missing}"
		)
	return "\n".join(lines)


def format_case_result_message(case_check_result):
	"""
	案件全体（複数マスク＋構造的NG）の結果を、マスクごとの見出し付きでまとめた
	日本語メッセージを組み立てる（結果ファイルへそのまま書き出す内容）。

	マスクが2件以上ある場合は、これまでと同様に`=== マスク: {ラベル} ===`見出しを各マスクの前に
	付ける。**マスクが1件だけの場合（`01_HEX関連`直下に`.mao`が1個だけ置かれる代表マスク構成）は
	見出しを付けず、そのマスクの結果（件数・判定）をそのまま案件レベルの結果として示す**
	（2026-08-31追補。単一マスクの案件で「マスク: 01_HEX関連」という技術的な見出しを出しても
	意味が薄いための簡略化）。

	**マスクが1件も無い場合（`overall_judgment=="対象外"`）は、判定OK/NGではなく「照合対象のMAO
	ファイルが見つかりませんでした」という趣旨の1行で示す**（2026-08-31追補。枝番案件・空マスク等、
	照合すべきインプットがそもそも存在しない案件をOK/NGと誤解されないようにするため）。
	"""
	if case_check_result.overall_judgment == "対象外":
		lines = [
			"結果: 対象外（照合対象のMAOファイルが見つかりませんでした）",
			f"（マスク数: 0件、構造的NG: {len(case_check_result.structural_ng_messages)}件）",
		]
	else:
		lines = [
			f"結果: {case_check_result.overall_judgment}",
			f"（マスク数: {len(case_check_result.mask_results)}件、"
			f"構造的NG: {len(case_check_result.structural_ng_messages)}件）",
		]

	if case_check_result.structural_ng_messages:
		lines.append("")
		lines.append("=== 構造的NG（同種ファイルの重複・ファイル未特定等） ===")
		for message in case_check_result.structural_ng_messages:
			lines.append(f"- {message}")

	show_mask_heading = len(case_check_result.mask_results) > 1
	for mask_result in case_check_result.mask_results:
		lines.append("")
		if show_mask_heading:
			lines.append(f"=== マスク: {mask_result.mask_label} ===")
		if mask_result.skipped_reason is not None:
			lines.append(f"結果: 判定不可（{mask_result.skipped_reason}）")
			continue
		lines.append(format_result_message(mask_result.check_result))

	return "\n".join(lines)


def write_result_file(text, output_path):
	"""
	組み立て済みのチェック結果メッセージ(text)を人間可読テキスト(output_path)として保存する。

	このテキストファイルは、案件フォルダの`01_INPUT`直下（`output_path`。「3.7」参照）へ本ツール
	自身が直接出力する成果物である。そのURLはSharePointリスト`XPX検証案件リスト2`の列
	`sp_client.LOG_FIELD_INTERNAL_NAME`（`inputFileCheckResultLog`）へ登録され、判定サマリは列
	`sp_client.RESULT_FIELD_INTERNAL_NAME`（`InputFileCheckResult`）へ登録される
	（`sp_main._process_one_item`が実装済み。2026-09-01）。

	引数:
		text: 出力するメッセージ本文（format_result_messageまたはformat_case_result_messageの戻り値）
		output_path: 人間可読テキストの出力先パス

	戻り値:
		テキスト出力先パス

	【2026-09-01追補・文字化け対策】SharePointへアップロードしたこのファイルをブラウザで直接開くと、
	charset情報を持たないため、日本語ロケール環境ではUTF-8のバイト列がShift-JISとして誤認され
	文字化けする（BOM無しUTF-8の典型的な弱点）。`utf-8-sig`（BOM付きUTF-8）で出力し、ブラウザが
	BOMからUTF-8と正しく判定できるようにする。
	"""
	output_path = Path(output_path)

	with open(output_path, "w", encoding="utf-8-sig") as text_file:
		text_file.write(text)
		text_file.write("\n")

	return output_path

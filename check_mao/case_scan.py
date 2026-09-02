"""案件フォルダ（`01_INPUT`配下）を走査し、マスクごとのMAOファイルとCANマトリクスファイルを
特定するモジュール（2026-08-31新規）。

【2026-09-01改修】チェック対象をCANテーブルからCANマトリクスへ変更したことに伴い、
`02_CAN関連`配下からの対象ファイル特定ロジックを反転した。従来はファイル名に"Matrix"を
含むものを**除外**してCANテーブルを特定していたが、以降はファイル名に"Matrix"を含むものを
**対象として採用**してCANマトリクスを特定する（`_find_matrix_path`。旧`_find_can_table_path`）。

背景・設計方針:
	- 複数マスク案件では`01_HEX関連`配下がマスク別サブフォルダ（`Mask{N}_HEX(...)`。PowerApps
	  ソース`App.pa.yaml`のOnStart・本申請ページ「③選択項目のリンク取得」で確認済み）に分かれ、
	  各フォルダに1つのMAOファイルが置かれる（MAOが置かれないマスクフォルダもある）。フォルダ名の
	  命名規則（呼称・仕向け・TM形式・駆動方式の組み合わせ）はPowerApps側の内部変数（マスク数
	  `intMaskNum`等）に依存し、Python側からは案件フォルダの実際のディレクトリ構造以外に手がかりが
	  無いため、本モジュールは**`01_HEX関連`配下を再帰的に走査し、MAOファイルを直接含むフォルダを
	  1マスクとして扱う**方式を採る（フォルダ名の書式そのものには依存しない）。
	- `02_CAN関連`配下にはCANテーブルとCANマトリクスの両方が格納される。**ファイル名に"Matrix"
	  （大小文字を区別しない）を含むものをCANマトリクスとみなして対象に採用**し、残り（CANテーブル・
	  補助ファイル等）は対象外とする（2026-09-01改修。「モジュールdocstring」冒頭参照）。
	- 1フォルダ内に同種ファイル（MAOまたはCANマトリクス）が複数存在する場合は、どちらを使うべきか
	  自動判別できないため判定NGとする（構造的NG。CAN ID突合とは別に結果ファイルへ記録する）。
	- zip共有時は解凍してからチェックする。案件フォルダ内に`.zip`が見つかった場合、一時フォルダへ
	  展開し、展開後の内容も走査対象に含める（SharePoint上の原本相当は書き換えない。展開先は
	  `tempfile`が作る一時ディレクトリとし、呼び出し側が使用後に削除する）。zip内に更にzipが
	  含まれる場合も再帰的に展開する。
	- 案件ID（SharePointリスト`XPX検証案件リスト2`の`Title`列＝PowerAppsの`strPageId_Daihyou`）を、
	  与えられたパスの`01_INPUT`の親フォルダ名から逆算する（`derive_case_id`、2026-08-31新規）。
	  `01_INPUT`が見つからない場合はフォルダ名から案件ID書式に一致する要素を探すフォールバックを
	  行う（2026-08-31追補。枝番案件対応、後述）。SharePoint連携方針が未整合のため、現状は判別結果を
	  コンソールへ表示する情報提供のみに用い、SharePointへの書き込みには使わない。

2026-08-31追補（本番想定36案件の実データ調査を踏まえた変更）:
	- 1案件に複数の枝番（`_02`以降）が並存し、代表となる枝番のみが`01_INPUT`を持つ構成（従属する
	  枝番は`02_委託見積`・`03_納品物`等はあるが`01_INPUT`・`01_HEX関連`・`02_CAN関連`が丸ごと無い）
	  が実データで多数確認された。これらはインプット不備ではなく「照合対象のMAOがそもそも無い
	  案件」であるため、MAOが1件も見つからない案件（空マスク・zip納品でMAO0件の案件も含む）は
	  `result.py`側で判定OK/NGではなく「対象外」として扱う（本モジュールの走査結果としては、
	  従来通りmask_entries=0件・structural_ng_messagesに理由を記録するのみで、判定区分自体は
	  `result.build_case_result`が決める）。
	- `01_HEX関連`配下がFI/ICM別サブフォルダに分かれ、それぞれの下にさらに`Mask{N}_HEX(...)`が
	  あるケース（202606_005_01・202608_006_01で実データ確認済み）では、マスクのラベルを
	  `01_HEX関連`からの相対パス（`FI/Mask1_HEX(...)`等）にする（`_mask_label`）。これにより、
	  FI/ICM双方に同名のMaskフォルダがあっても、結果ファイルの見出しが重複しない。
"""

import re
import tempfile
import zipfile
from pathlib import Path

from can_info import looks_like_matrix, select_matrix_sheet

HEX_FOLDER_NAME = "01_HEX関連"
CAN_FOLDER_NAME = "02_CAN関連"
INPUT_FOLDER_NAME = "01_INPUT"

MAO_SUFFIX = ".mao"
MATRIX_FILE_SUFFIXES = (".xlsm", ".xlsx", ".xls")
MATRIX_NAME_MARKER = "matrix"  # ファイル名にこの文字列(大小無視)を含むものをCANマトリクスとして採用

# 案件ID（SharePointリスト`XPX検証案件リスト2`の`Title`列＝PowerAppsの`strPageId_Daihyou`）の
# 実データ上のフォーマット。年度月6桁_連番3桁(_枝番2桁は省略可)。実データ155件（2026年度01月〜08月）を
# 確認した結果、英字混在の枝番（資料上の例："002E"）は実在しなかったため数字限定とする
# （2026-08-31調査。将来英字混在が実在した場合は本パターンの見直しが必要）。
CASE_ID_PATTERN = re.compile(r"^20\d{4}_\d{3}(?:_\d{2})?$")


class MaskEntry:
	"""1マスク分の情報（マスクフォルダとそこにあるMAOファイル）を保持するレコード。"""

	def __init__(self, label, mao_path):
		self.label = label  # 結果ファイルに表示するマスクの見出し（フォルダ名相当）
		self.mao_path = mao_path


class CaseScanResult:
	"""案件フォルダの走査結果。マスク一覧・CANマトリクスパス・構造的NGメッセージ・案件IDを保持する。"""

	def __init__(
		self,
		mask_entries,
		matrix_path,
		structural_ng_messages,
		work_directory,
		case_id=None,
		case_id_format_ok=None,
		matrix_sheet_name=None,
		matrix_header_row=None,
		matrix_can_id_column=None,
	):
		self.mask_entries = mask_entries  # MaskEntryのリスト（MAOが1個のフォルダのみ。順序はフォルダ名昇順）
		self.matrix_path = matrix_path  # 特定できた場合のPath、特定できない場合はNone（2026-09-01改修。旧can_table_path）
		self.structural_ng_messages = structural_ng_messages  # 走査時点で判明した構造的NGの説明文リスト
		self.work_directory = work_directory  # zip展開に使った一時ディレクトリ（後始末用。無ければNone）
		# 案件ID（derive_case_idの戻り値。SharePoint未連携のため現状は情報表示のみに使う。「9.」参照）
		self.case_id = case_id
		self.case_id_format_ok = case_id_format_ok  # case_idがCASE_ID_PATTERNに一致するか。case_idがNoneならNone
		# CANマトリクス本体として確定したシート名・ヘッダ行・CAN ID列（2026-09-01改修。matrix_pathが
		# 特定できた際に`select_matrix_sheet`で確定した値。抽出時の再計算・二重読込を避けるため保持する）
		self.matrix_sheet_name = matrix_sheet_name
		self.matrix_header_row = matrix_header_row
		self.matrix_can_id_column = matrix_can_id_column


def _is_lock_file(file_name):
	"""Excelが開いている間に作る一時ロックファイル（"~$"始まり）を除外する。"""
	return file_name.startswith("~$")


def _extract_zips_recursively(root_directory, work_directory, extracted_roots):
	"""
	root_directory配下の.zipを再帰的に見つけて展開し、展開先ディレクトリをextracted_rootsへ追加する。

	zip内に更にzipが含まれる場合も、展開後のディレクトリに対して再度本処理を適用することで対応する。
	"""
	for path in root_directory.rglob("*.zip"):
		if not path.is_file():
			continue
		destination = Path(tempfile.mkdtemp(prefix="unzip_", dir=str(work_directory)))
		try:
			with zipfile.ZipFile(path) as zip_file:
				zip_file.extractall(destination)
		except zipfile.BadZipFile:
			# 壊れたzip・zipを名乗る非zipファイルは展開せず読み飛ばす（構造的NGとしては扱わない。
			# 対象マスク・CANテーブルが見つからなければ別途「見つからない」NGとして表面化する）
			continue
		extracted_roots.append(destination)
		_extract_zips_recursively(destination, work_directory, extracted_roots)


def _find_dirs_by_name(roots, folder_name):
	"""複数の探索起点(roots)配下から、指定フォルダ名に完全一致するディレクトリを再帰的に全て集める。"""
	found = []
	for root in roots:
		if root.name == folder_name and root.is_dir():
			found.append(root)
		for path in root.rglob(folder_name):
			if path.is_dir():
				found.append(path)
	return found


def _mask_label(folder, hex_folder):
	"""
	マスクのラベル（結果ファイルの見出しに使うフォルダ名相当）を組み立てる。

	`01_HEX関連`直下（folder自身がhex_folderの場合＝直置き構成）は`01_HEX関連`のフォルダ名を
	そのまま返す（従来どおり）。それ以外は、マッチした`01_HEX関連`からの相対パスを`/`区切りで
	返す。これにより、FI/ICM配下にそれぞれ`Mask1_HEX(...)`のような同名フォルダがある構成
	（202606_005_01・202608_006_01で実データ確認済み）でも、`FI/Mask1_HEX(...)`・
	`ICM/Mask1_HEX(...)`のように見出しが重複しない（2026-08-31追補）。
	"""
	if folder == hex_folder:
		return hex_folder.name
	return folder.relative_to(hex_folder).as_posix()


def _find_mask_entries(hex_folders):
	"""
	`01_HEX関連`配下（複数見つかった場合は全て）を再帰的に走査し、MAOファイルを直接含む
	フォルダをマスクとして列挙する。1フォルダに複数のMAOがある場合は構造的NGとして記録する。

	戻り値:
		(MaskEntryのリスト（フォルダパス昇順）, 構造的NGメッセージのリスト)
	"""
	mask_entries = []
	structural_ng_messages = []

	# フォルダパス(str) -> (フォルダ, マッチした01_HEX関連) の辞書。ラベル算出に相対パスの
	# 起点となるhex_folderが必要なため、setではなく起点を保持できるdictで管理する。
	candidate_dirs = {}
	for hex_folder in hex_folders:
		candidate_dirs.setdefault(str(hex_folder), (hex_folder, hex_folder))
		for path in hex_folder.rglob("*"):
			if path.is_dir():
				candidate_dirs.setdefault(str(path), (path, hex_folder))

	for _, (folder, owning_hex_folder) in sorted(candidate_dirs.items()):
		mao_files = sorted(
			p for p in folder.iterdir() if p.is_file() and p.suffix.lower() == MAO_SUFFIX
		)
		if len(mao_files) == 0:
			continue
		if len(mao_files) > 1:
			names = "、".join(p.name for p in mao_files)
			structural_ng_messages.append(
				f"マスクフォルダ「{folder}」内に同種ファイル(MAO)が複数存在するため判定不能: {names}"
			)
			continue
		mask_entries.append(MaskEntry(_mask_label(folder, owning_hex_folder), mao_files[0]))

	return mask_entries, structural_ng_messages


def _find_matrix_path(can_folders):
	"""
	`02_CAN関連`配下（複数見つかった場合は全て）から、CANマトリクス候補ファイルを特定する
	（2026-09-01改修。旧`_find_can_table_path`。判定ロジックを反転し、ファイル名に"Matrix"を
	含むものを対象として採用する）。

	ファイル名に"Matrix"(大小無視)を含むものをCANマトリクス候補として採用する。さらに、
	`can_info.looks_like_matrix`でファイルの内容（`CAN ID (HEX)`ヘッダを持つシートの有無）を
	確認し、ファイル名に"Matrix"を含むが実際にはCAN ID (HEX)列を持たないファイルを除外する。
	残った候補がちょうど1個ならそれを採用し、0個または2個以上なら構造的NGとして記録しNoneを返す。

	採用したファイルについては、`can_info.select_matrix_sheet`でCANマトリクス本体のシート・
	ヘッダ行・CAN ID列を確定する。1ブックに複数の`CAN ID (HEX)`ヘッダシートがあり一意に選べない
	場合は、誤ったシートを選ばず判定不可の構造的NGとして記録する。

	戻り値:
		(CANマトリクスのPathまたはNone, シート名またはNone, ヘッダ行番号またはNone,
		 CAN ID列番号またはNone, 構造的NGメッセージのリスト)
	"""
	candidates = []
	for can_folder in can_folders:
		for path in can_folder.rglob("*"):
			if not path.is_file():
				continue
			if _is_lock_file(path.name):
				continue
			if path.suffix.lower() not in MATRIX_FILE_SUFFIXES:
				continue
			if MATRIX_NAME_MARKER not in path.stem.lower():
				continue
			candidates.append(path)

	content_candidates = [path for path in candidates if looks_like_matrix(path)]

	if len(content_candidates) == 0:
		message = (
			f"CANマトリクスが見つからない({CAN_FOLDER_NAME}配下にファイル名'Matrix'を含む"
			"CAN情報ファイルが無い、またはCAN ID (HEX)列を持つシートが無い)"
		)
		return None, None, None, None, [message]
	if len(content_candidates) > 1:
		names = "、".join(str(p) for p in content_candidates)
		return None, None, None, None, [f"CANマトリクスが複数存在するため判定不能: {names}"]

	matrix_path = content_candidates[0]
	selected_sheet = select_matrix_sheet(matrix_path)
	if selected_sheet is None:
		message = f"CANマトリクスのシートを一意に特定できないため判定不能(要手動確認): {matrix_path}"
		return None, None, None, None, [message]

	sheet_name, header_row, can_id_column = selected_sheet
	return matrix_path, sheet_name, header_row, can_id_column, []


def is_valid_case_id_format(text):
	"""案件IDの候補文字列が実データ上のフォーマット（`CASE_ID_PATTERN`）に一致するかを判定する。"""
	if text is None:
		return False
	return CASE_ID_PATTERN.match(text) is not None


def derive_case_id(case_folder_path):
	"""
	案件フォルダのパスから案件ID（SharePointリスト`XPX検証案件リスト2`の`Title`列＝PowerAppsの
	`strPageId_Daihyou`に相当）を逆算する。

	実データ調査（2026-08-31、`SDV XPX検証運営掲示板`配下の2026年度01月〜08月、案件フォルダ155件）で、
	案件フォルダは例外なく`{年度}年度\\{月}月\\{案件ID}\\01_INPUT\\...`という構造をとり、`01_INPUT`の
	親フォルダ名が案件IDであることを確認した。本関数はこの構造にのみ依存し、`01_INPUT`より上の階層数
	（年度・月フォルダの有無等）には依存しない。

	引数として渡されたパスが`01_INPUT`自身であっても、案件ルート（`01_INPUT`を直接含む親）であっても、
	それより深いフォルダ（`01_HEX関連`等）であっても、パスの各構成要素から`01_INPUT`という名前の要素を
	探し、その直前の要素を案件IDの候補として返す（ファイルシステムを実際に読みには行かず、パス文字列
	だけで判定するため、`01_INPUT`が実在するかどうかは問わない）。

	`01_INPUT`が見つからない場合は、フォールバックとしてパスの構成要素を末尾側から見て、案件ID書式
	（`CASE_ID_PATTERN`）に一致する要素を候補として返す（2026-08-31追補）。本番想定データ調査で、
	枝番`_02`以降の案件（1案件に複数の枝番が並存し、代表となる枝番のみが`01_INPUT`を持つ構成。
	`02_委託見積`・`03_納品物`等はあるが`01_INPUT`が無い）が多数存在することを確認しており、これらの
	案件ルートを渡された場合でもフォルダ名から案件IDを判別できるようにするため。

	引数:
		case_folder_path: 案件フォルダ関連のいずれかのパス（Path変換前でも可）

	戻り値:
		案件IDの候補文字列。`01_INPUT`が見つかりその直前の要素があれば、それを最優先で返す。
		`01_INPUT`が先頭要素で直前の要素が無い場合、または`01_INPUT`自体が見つからない場合は、
		末尾側から`CASE_ID_PATTERN`に一致する要素を探し、見つかればそれを返す。いずれの手段でも
		候補が得られない場合はNone。

	注意：
		本関数はSharePointへの書き込みには使用しない（SharePoint連携方針が別担当側で整合してから
		別途実装する。「内部仕様書9. 自動化実装再着手時の参考資産」参照）。現状は判別結果を
		コンソール表示する情報提供のみに用いる。
	"""
	parts = Path(case_folder_path).parts
	for index, part in enumerate(parts):
		if part == INPUT_FOLDER_NAME:
			if index == 0:
				return None
			return parts[index - 1]

	for part in reversed(parts):
		if CASE_ID_PATTERN.match(part):
			return part
	return None


def resolve_input_folder_path(case_folder_path):
	"""
	与えられたパスから`01_INPUT`フォルダそのもののパスを組み立てる（2026-08-31追補4。結果ファイルを
	`01_INPUT`直下へ自動出力するために、`main.py`が使用する）。

	`derive_case_id`と同様、引数として渡されたパスが`01_INPUT`自身であっても、案件ルート
	（`01_INPUT`を直接含む親）であっても、それより深いフォルダ（`01_HEX関連`等）であっても、
	パスの各構成要素から`01_INPUT`という名前の要素を探し、その要素までのパス（`01_INPUT`自身を
	含む）を返す。ファイルシステムを実際に読みには行かず、パス文字列だけで判定するため、
	`01_INPUT`が実在するかどうかは問わない（実在確認は呼び出し側の責務）。

	引数:
		case_folder_path: 案件フォルダ関連のいずれかのパス（Path変換前でも可）

	戻り値:
		`01_INPUT`フォルダのPath。パスの構成要素に`01_INPUT`が見つからない場合はNone
		（`01_INPUT`を持たない従属案件（枝番案件等）の案件ルートを渡した場合に相当。この場合、
		通常は`scan_case_folder`の結果が「対象外」になり、出力自体を行わないため実害は無い想定。
		「内部仕様書」参照）。
	"""
	parts = Path(case_folder_path).parts
	for index, part in enumerate(parts):
		if part == INPUT_FOLDER_NAME:
			return Path(*parts[: index + 1])
	return None


def scan_case_folder(input_root):
	"""
	案件フォルダ（案件ルート、または`01_INPUT`そのもの）を走査し、マスク一覧・CANテーブルパス・
	構造的NGをまとめたCaseScanResultを返す。

	引数:
		input_root: 案件フォルダのパス（`01_INPUT`を含む親でも、`01_INPUT`自身でもよい。
		            `01_HEX関連`・`02_CAN関連`は配下を再帰的に探索するため、両者の相対位置は問わない）

	戻り値:
		CaseScanResult。zipを展開した場合は`work_directory`に一時ディレクトリを保持するので、
		呼び出し側は使用後に`shutil.rmtree(result.work_directory, ignore_errors=True)`で削除すること
		（展開しなかった場合は`work_directory`はNone）。
	"""
	input_root = Path(input_root)

	case_id = derive_case_id(input_root)
	case_id_format_ok = is_valid_case_id_format(case_id) if case_id is not None else None

	work_directory = None
	extracted_roots = []
	if any(input_root.rglob("*.zip")):
		work_directory = Path(tempfile.mkdtemp(prefix="can_mao_check_zip_"))
		_extract_zips_recursively(input_root, work_directory, extracted_roots)

	all_roots = [input_root] + extracted_roots

	hex_folders = _find_dirs_by_name(all_roots, HEX_FOLDER_NAME)
	can_folders = _find_dirs_by_name(all_roots, CAN_FOLDER_NAME)

	structural_ng_messages = []

	if not hex_folders:
		structural_ng_messages.append(f"{HEX_FOLDER_NAME}フォルダが見つからない")
		mask_entries = []
	else:
		mask_entries, mask_ng_messages = _find_mask_entries(hex_folders)
		structural_ng_messages.extend(mask_ng_messages)

	if not can_folders:
		structural_ng_messages.append(f"{CAN_FOLDER_NAME}フォルダが見つからない")
		matrix_path = None
		matrix_sheet_name = None
		matrix_header_row = None
		matrix_can_id_column = None
	else:
		(
			matrix_path,
			matrix_sheet_name,
			matrix_header_row,
			matrix_can_id_column,
			matrix_ng_messages,
		) = _find_matrix_path(can_folders)
		structural_ng_messages.extend(matrix_ng_messages)

	return CaseScanResult(
		mask_entries,
		matrix_path,
		structural_ng_messages,
		work_directory,
		case_id=case_id,
		case_id_format_ok=case_id_format_ok,
		matrix_sheet_name=matrix_sheet_name,
		matrix_header_row=matrix_header_row,
		matrix_can_id_column=matrix_can_id_column,
	)

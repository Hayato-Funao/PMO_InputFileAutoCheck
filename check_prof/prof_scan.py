"""
チェック③(HEX/A2Lチェック)の入力ファイルを案件フォルダから特定する(2026-09-02新規)。

①(check_mao)の`case_scan.py`と同じ作法に倣う:
	- 同種ファイルが複数ある場合はどちらを使うべきか自動判別できないため構造的NGとする
	  (資料スライド2「同じ種類のファイルが複数添付された場合（HEXが2つ、など）、
	  チェックはNGとする」)
	- Excelのロックファイル(`~$`で始まるもの)は除外する
	- 探索ルートは複数取れる(`01_INPUT`本体と、①がzipを展開した一時ディレクトリ)

【zip対応について】
zipの展開は①の`case_scan.scan_case_folder`が既に行っており、展開先は
`CaseScanResult.work_directory`として公開されている。③はそれを追加の探索ルートとして
受け取ることで、zip展開を自前で実装しない。ただし`work_directory`は
`unified_main._process_one_item`の内側`finally`で削除されるため、③はその
tryブロック内で実行する必要がある。
"""

import os

# 種別名(結果メッセージにそのまま出す)
KIND_HEX = "HEX"
KIND_EPD = "ProF(.epd)"
KIND_A2L = "A2L"

# 種別 → 対象拡張子(小文字で比較する)
FILE_KINDS = (
	(KIND_HEX, (".hex",)),
	(KIND_EPD, (".epd",)),
	(KIND_A2L, (".a2l",)),
)

# Excelが作る一時ロックファイルの接頭辞(`case_scan._is_lock_file`と同じ考え方)
LOCK_FILE_PREFIX = "~$"


class ProfScanResult:
	"""案件フォルダの走査結果。③の入力3点と構造的NGメッセージを保持する。"""

	def __init__(self, hex_path, epd_path, a2l_path, structural_ng_messages):
		self.hex_path = hex_path  # 一意に特定できた場合のパス、できない場合はNone
		self.epd_path = epd_path
		self.a2l_path = a2l_path
		self.structural_ng_messages = structural_ng_messages  # 構造的NGの説明文リスト

	@property
	def is_complete(self):
		"""③の判定に必要な3ファイルすべてが一意に特定できたか"""
		return bool(self.hex_path) and bool(self.epd_path) and bool(self.a2l_path)


def _is_lock_file(file_name):
	"""Excelのロックファイル(`~$`始まり)かどうか"""
	return file_name.startswith(LOCK_FILE_PREFIX)


def _normalize_search_roots(local_input_dir, extra_search_roots):
	"""探索ルートを1本のリストにまとめる(Noneや文字列単体も受け付ける)"""
	roots = [local_input_dir]
	if extra_search_roots:
		if isinstance(extra_search_roots, (str, bytes, os.PathLike)):
			roots.append(extra_search_roots)
		else:
			roots.extend(extra_search_roots)
	return [str(root) for root in roots if root]


def _collect_by_suffix(search_roots, suffixes):
	"""探索ルート配下から拡張子が一致するファイルを集める(同一ファイルは重複排除)。

	同じファイルが`01_INPUT`直下とzip展開先の両方にある場合(zipに同じHEXが
	同梱されている等)を「2個」と誤判定しないよう、`(ファイル名小文字, サイズ)`が
	一致するものは同一ファイルのコピーとみなして1個に数える。
	名前もサイズも同じ別物という可能性は実用上無視できる。

	戻り値:
		パスのリスト(重複排除済み。順序は安定させるためソート済み)
	"""
	found = {}
	for root in search_roots:
		if not os.path.isdir(root):
			continue
		for dir_path, _dir_names, file_names in os.walk(root):
			for file_name in file_names:
				if _is_lock_file(file_name):
					continue
				if not file_name.lower().endswith(suffixes):
					continue
				path = os.path.join(dir_path, file_name)
				try:
					size = os.path.getsize(path)
				except OSError:
					# 走査中に消えた等。集計対象から外す(存在しないものは判定できない)
					continue
				found.setdefault((file_name.lower(), size), path)
	return [found[key] for key in sorted(found)]


def scan_prof_inputs(local_input_dir, extra_search_roots=None):
	"""案件フォルダから③の入力3点(HEX / .epd / .a2l)を特定する。

	引数:
		local_input_dir: ダウンロード済み`01_INPUT`のローカルパス
		extra_search_roots: 追加の探索ルート(①のzip展開先など)。文字列単体・
			リスト・Noneのいずれでも可

	戻り値:
		ProfScanResult

	注意:
		未提出・複数存在はいずれも構造的NGとして`structural_ng_messages`へ積む
		(2026-09-02決定。`確認不能`や`対象外`は使わず、すべてNGへ畳む。
		資料スライド1・2がHEX/A2L/ProFを提出必須のINPUTとして挙げているため)。
	"""
	search_roots = _normalize_search_roots(local_input_dir, extra_search_roots)

	selected = {}
	structural_ng_messages = []
	for kind, suffixes in FILE_KINDS:
		paths = _collect_by_suffix(search_roots, suffixes)
		if len(paths) == 0:
			selected[kind] = None
			structural_ng_messages.append(f"{kind}ファイルが提出されていないため判定不能")
			continue
		if len(paths) > 1:
			names = "、".join(os.path.basename(path) for path in paths)
			selected[kind] = None
			structural_ng_messages.append(
				f"同種ファイル({kind})が複数存在するため判定不能: {names}"
			)
			continue
		selected[kind] = paths[0]

	return ProfScanResult(
		selected[KIND_HEX],
		selected[KIND_EPD],
		selected[KIND_A2L],
		structural_ng_messages,
	)

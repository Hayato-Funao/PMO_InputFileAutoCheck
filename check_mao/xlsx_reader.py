"""xlsx/xlsm（OOXML）・xls（旧バイナリ形式）を読み込むための最小限のリーダー。

xlsx/xlsmはzipfile + xml.etree.ElementTreeで内部XML（共有文字列・ワークシートXML）を
直接パースする。数式("f"要素)は評価せず、Excel側が保存時にキャッシュした値("v"要素)のみを
読み取る。xls（OLE2/CFBF形式。3EL・3QJのCANテーブル等で実在を確認）はzipファイルでは
ないため、xlrdで読み込む（xlrdは.xls分岐でのみ遅延importし、xlsx/xlsm専用運用では
未導入でも動作する）。
"""

import re
import zipfile
from xml.etree import ElementTree as ET

# xlsxの名前空間定義
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

CELL_REFERENCE_PATTERN = re.compile(r"([A-Za-z]+)(\d+)")


def column_letters_to_index(letters):
	"""セル参照のアルファベット部分（例:"AB"）を1始まりの列番号に変換する。"""
	index = 0
	for char in letters:
		index = index * 26 + (ord(char.upper()) - ord("A") + 1)
	return index


def parse_cell_reference(reference):
	"""「AB123」のようなセル参照文字列から(行番号, 列番号)を取り出す（いずれも1始まり）。"""
	matched = CELL_REFERENCE_PATTERN.match(reference)
	letters, digits = matched.group(1), matched.group(2)
	return int(digits), column_letters_to_index(letters)


def load_shared_strings(zip_file):
	"""xl/sharedStrings.xmlを読み、インデックス→文字列のリストを返す。共有文字列が無ければ空リスト。"""
	if "xl/sharedStrings.xml" not in zip_file.namelist():
		return []

	root = ET.fromstring(zip_file.read("xl/sharedStrings.xml"))
	strings = []
	for item in root.findall(f"{{{NS_MAIN}}}si"):
		# 単純な文字列(t直下)とリッチテキスト(r要素の連結)の両方に対応する
		direct_text = item.find(f"{{{NS_MAIN}}}t")
		if direct_text is not None:
			strings.append(direct_text.text or "")
			continue
		parts = []
		for run in item.findall(f"{{{NS_MAIN}}}r"):
			run_text = run.find(f"{{{NS_MAIN}}}t")
			if run_text is not None:
				parts.append(run_text.text or "")
		strings.append("".join(parts))
	return strings


def load_sheet_name_map(zip_file):
	"""xl/workbook.xmlとxl/_rels/workbook.xml.relsから、シート名→シートXMLパスの辞書を作る。"""
	workbook_root = ET.fromstring(zip_file.read("xl/workbook.xml"))
	rels_root = ET.fromstring(zip_file.read("xl/_rels/workbook.xml.rels"))

	# 関係ID(r:id) -> ターゲットパス のマップを作る
	relationship_id_to_target = {}
	for relationship in rels_root.findall(f"{{{NS_REL_PKG}}}Relationship"):
		relationship_id_to_target[relationship.get("Id")] = relationship.get("Target")

	sheet_name_to_path = {}
	for sheet in workbook_root.findall(f"{{{NS_MAIN}}}sheets/{{{NS_MAIN}}}sheet"):
		sheet_name = sheet.get("name")
		relationship_id = sheet.get(f"{{{NS_REL_DOC}}}id")
		target = relationship_id_to_target.get(relationship_id, "")
		# Targetは通常"worksheets/sheet1.xml"のような相対パスだが、生成ツールによっては
		# 先頭に"/"が付く絶対パス形式("/xl/worksheets/sheet1.xml")の場合もある。
		# 先頭の"/"を除いてから"xl/"を補うことで、いずれの形式でも二重接頭を防ぐ。
		target = target.lstrip("/")
		if not target.startswith("xl/"):
			target = "xl/" + target
		sheet_name_to_path[sheet_name] = target
	return sheet_name_to_path


def list_sheet_names(workbook_path):
	"""
	xlsx/xlsm/xlsファイルの全シート名を、ワークブック内の宣言順で返す（2026-08-31追補。
	CANテーブルの内容判別で全シートを走査するために使う）。

	引数:
		workbook_path: 読み込むxlsx/xlsm/xlsファイルのパス

	戻り値:
		シート名のリスト（宣言順）
	"""
	if zipfile.is_zipfile(workbook_path):
		with zipfile.ZipFile(workbook_path) as zip_file:
			return list(load_sheet_name_map(zip_file).keys())

	import xlrd

	workbook = xlrd.open_workbook(workbook_path)
	return list(workbook.sheet_names())


def read_sheet(workbook_path, sheet_name=None):
	"""
	xlsx/xlsm/xlsファイルを読み込み、指定シートを行×列の値配列として返す。

	ファイル実体（拡張子ではなく先頭バイト）でxlsx/xlsm（ZIP=OOXML）とxls（OLE2バイナリ）を
	判別し、前者はzipfile+XMLで、後者はxlrdで読み込む。戻り値の形式は両方式で共通。

	引数:
		workbook_path: 読み込むxlsx/xlsm/xlsファイルのパス
		sheet_name: 読み込むシート名。省略時はワークブック内の先頭シートを読む

	戻り値:
		二次元リスト。rows[行番号-1][列番号-1] で値を取得できる（0始まりインデックス）。
		セルが存在しない位置はNoneになる。値は文字列（xlsの数値セルは整数化した文字列。
		小数を含む場合はそのまま文字列化）で返す（数値への変換は呼び出し側の責務とする）。
	"""
	if zipfile.is_zipfile(workbook_path):
		return _read_sheet_ooxml(workbook_path, sheet_name)
	return _read_sheet_legacy_xls(workbook_path, sheet_name)


def _numeric_cell_to_text(value):
	"""xlrdが返すfloatのセル値を文字列化する（整数値は小数点無しにする）。"""
	if value == int(value):
		return str(int(value))
	return str(value)


def _read_sheet_legacy_xls(workbook_path, sheet_name=None):
	"""xls（OLE2バイナリ形式）をxlrdで読み込み、read_sheetと同じ二次元リスト形式で返す。"""
	import xlrd

	workbook = xlrd.open_workbook(workbook_path)
	sheet = workbook.sheet_by_name(sheet_name) if sheet_name is not None else workbook.sheet_by_index(0)

	rows = []
	for row_index in range(sheet.nrows):
		row = []
		for column_index in range(sheet.ncols):
			cell = sheet.cell(row_index, column_index)
			if cell.ctype == xlrd.XL_CELL_EMPTY or cell.ctype == xlrd.XL_CELL_BLANK:
				value = None
			elif cell.ctype == xlrd.XL_CELL_NUMBER:
				value = _numeric_cell_to_text(cell.value)
			else:
				value = str(cell.value)
			row.append(value)
		rows.append(row)
	return rows


def _read_sheet_ooxml(workbook_path, sheet_name=None):
	"""xlsx/xlsm（OOXML）をzipfile+XMLで読み込み、行×列の値配列を返す（read_sheet本体）。"""
	with zipfile.ZipFile(workbook_path) as zip_file:
		shared_strings = load_shared_strings(zip_file)
		sheet_name_to_path = load_sheet_name_map(zip_file)

		if sheet_name is None:
			target_path = next(iter(sheet_name_to_path.values()))
		else:
			target_path = sheet_name_to_path[sheet_name]

		sheet_root = ET.fromstring(zip_file.read(target_path))

		max_row = 0
		max_column = 0
		cell_values = {}  # (行番号, 列番号) -> 値

		for row_element in sheet_root.findall(f"{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
			for cell_element in row_element.findall(f"{{{NS_MAIN}}}c"):
				reference = cell_element.get("r")
				row_index, column_index = parse_cell_reference(reference)
				cell_type = cell_element.get("t")
				value_element = cell_element.find(f"{{{NS_MAIN}}}v")

				if cell_type == "s" and value_element is not None:
					value = shared_strings[int(value_element.text)]
				elif cell_type == "inlineStr":
					inline_string = cell_element.find(f"{{{NS_MAIN}}}is")
					inline_text = inline_string.find(f"{{{NS_MAIN}}}t") if inline_string is not None else None
					value = inline_text.text if inline_text is not None else ""
				elif value_element is not None:
					value = value_element.text
				else:
					value = None

				cell_values[(row_index, column_index)] = value
				max_row = max(max_row, row_index)
				max_column = max(max_column, column_index)

		rows = []
		for row_index in range(1, max_row + 1):
			row = []
			for column_index in range(1, max_column + 1):
				row.append(cell_values.get((row_index, column_index)))
			rows.append(row)

		return rows

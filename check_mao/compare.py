"""MAO・CANマトリクスの2ソースを、正規化CAN IDをキーに突合するモジュール。

2ソースの存在有無はそのまま事実として記録する（category）。加えて、内部仕様書4.5節の
MAOを起点とした非対称判定（OK／NG／－の3値。judgment）を付与する。「対象外CAN ID」の
独立した判別・除外リストは設けず、CAN情報側（CANマトリクス）のみに存在しMAOに無いCAN IDは、
判定対象外として「－」とする（この非対称ルールで代替する）。

【2026-09-01改修】チェック対象をCANテーブルからCANマトリクスへ変更した（船尾さんの指示）。
判定ロジック自体（MAO駆動の非対称3値判定）は変更していない。2026-08-31改修時点では逆に
CANマトリクスを対象外としCANテーブル単独で判定していたが、本改修で再度CANマトリクスを
対象とする形に戻した（CANテーブルは当面未使用。`can_info.py`のモジュールdocstring参照）。
"""

from mao_can_id import build_unique_can_id_map


class ComparisonRow:
	"""1つの正規化CAN IDについて、2ソースの存在有無と付随情報をまとめた行。"""

	def __init__(self, can_id_decimal):
		self.can_id_decimal = can_id_decimal
		self.mao_records = []
		self.matrix_records = []

	@property
	def has_mao(self):
		return len(self.mao_records) > 0

	@property
	def has_matrix(self):
		return len(self.matrix_records) > 0

	@property
	def category(self):
		"""2ソースの存在状況から区分名を決める（対象外判定ではなく事実区分）。"""
		presence = (self.has_mao, self.has_matrix)
		category_by_presence = {
			(True, True): "MAO・マトリクス共通",
			(True, False): "MAO限定",
			(False, True): "マトリクス限定",
		}
		return category_by_presence.get(presence, "該当なし")

	@property
	def judgment(self):
		"""
		内部仕様書4.5節のMAO駆動非対称判定を返す（"OK"／"NG"／"－"の3値）。

		MAOに存在しないCAN IDは判定対象外として"－"とする。MAOに存在するCAN IDは、
		CANマトリクスにも存在すれば"OK"、存在しなければ"NG"（要問合せ）とする。
		"""
		if not self.has_mao:
			return "－"
		if self.has_matrix:
			return "OK"
		return "NG"


def compare_can_id_sources(mao_records, matrix_records):
	"""
	2ソースのレコード一覧を正規化CAN IDで突合し、ComparisonRowのリストを返す（正規化ID昇順）。
	"""
	mao_map = build_unique_can_id_map(mao_records)

	matrix_map = {}
	for record in matrix_records:
		if record.can_id_decimal is None:
			continue
		matrix_map.setdefault(record.can_id_decimal, []).append(record)

	all_can_ids = set(mao_map) | set(matrix_map)

	comparison_rows = []
	for can_id_decimal in sorted(all_can_ids):
		row = ComparisonRow(can_id_decimal)
		row.mao_records = mao_map.get(can_id_decimal, [])
		row.matrix_records = matrix_map.get(can_id_decimal, [])
		comparison_rows.append(row)

	return comparison_rows

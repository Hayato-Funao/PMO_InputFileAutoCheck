"""
②(check_fs_matrix)の結果ログ（詳細ファイル）本文を組み立てるモジュール（2026-09-03新規）。

従来は`=== OK ===`／`=== NG ===`の2セクションに全チェックの行を振り分けて出していたが、
「どのファイルの何を見た結果なのか」がラベル（[元表]／[FS Matrix]等）頼りで読みにくかったため、
チェック対象のファイル単位の3グループへ並べ替える形に変更した。OK/NGは行の振り分けではなく
各行の文末（ある／ない、一致している／一致していない）で表し、NGの行には不備箇所
（セル番地等）を続けて書く。

出力例:

    実行日時: 2026-09-03 15:00:00

    [定義ファイル管理Excel]
    ・「XPX定義ファイル管理_5.22_FI_ELEC」の「元表」シートとSP上のFSマトリクスのシート内容が一致している。
    ・「XPX定義ファイル管理_5.22_FI_ELEC」の「fcl_list_o_FI」にデータ不備がない。
    ・「XPX定義ファイル管理_5.22_FI_ELEC」の「fcl_list_x_FI」にデータ不備がない。

    [FI FSマトリクスファイル]
    ・ツールが想定している「FI FSmatrix」シートがある。
    ・「FI FSmatrix」シートの適用（〇/×）欄にデータ不備がない。

    [GSTコマンド装備表]
    ・ツールが想定している「SID一覧」シートがある。
    ・「SID一覧」シートの適用（〇/×）欄にデータ不備がない。

各行の文言は`definition_file_check`／`matrix_check`／`gst_command_check`が返すmessageそのもの
（本モジュールは並べ替えと箇条書き記号の付与だけを行い、文言は作らない）。行が1つも無い
グループは見出しごと出さない（例: GSTコマンド装備表の添付が無い案件）。
"""

import os
from datetime import datetime

# グループ見出し。②のチェック対象ファイルの呼び名に合わせる。
GROUP_DEFINITION_FILE = "定義ファイル管理Excel"
GROUP_FS_MATRIX = "FI FSマトリクスファイル"
GROUP_GST_COMMAND = "GSTコマンド装備表"

# 提出FSマトリクス↔定義ファイル管理をまとめて扱うグループ（2026-09-04）。
# 中身は次の順に並ぶ:
#   1. `元表`（A〜J列＋Q列）との突合（`definition_file_check.check_fs_matrix_codes_in_motohyou`
#      ／`check_fs_matrix_data_mismatches`）
#   2. `fcl_list_o_FI`／`fcl_list_x_FI`の〇/×データ不備（`check_definition_file_parts`）
# 1がNGの場合は2を実行しない（`unified_main._run_check2`で打ち切る）。
#
# `GROUP_DEFINITION_FILE`（旧・②の1グループ目）は残してある。方式Aの`input_check_main_ref.py`が
# 元表突合を持たず、そちらの見出しを「FSマトリクス↔…」にすると実態と合わないため、
# 統合実行（方式B）の`unified_main`だけがこの新しい見出しを使う。
GROUP_FS_MATRIX_DEFINITION = "FSマトリクス↔定義ファイル管理Excel"

# 集約ファイル（`インプットファイルチェック結果.txt`）の②の行末へ**併記しない**グループ。
# 詳細ファイル（本モジュールが書く`FSマトリクス_定義ファイルチェック結果.txt`）には出るが、
# 集約ファイルには出さない（`unified_main._run_check2`が参照する）。
#
#   * `定義ファイル管理Excel`／`FSマトリクス↔定義ファイル管理Excel`:
#     〇/×データ不備はPMO管理の固定ファイルに対するチェックで全案件同じ結果になり、案件担当者が
#     集約ファイルを見て対処できる情報ではない。元表突合も、機種ごとに新規DTCが増えるのは正常で
#     差異＝不備とは限らず、実データでは1機種あたり百件超の差異が常時出るため、一覧を集約ファイルへ
#     出すと本当のNGが埋もれる。
#     ただし元表突合のうち判定に算入する2件（コード欠落／HILSコードに影響する不一致）だけは、
#     一覧ではなく短文を`unified_main._run_check2`が行単位で足す。
GROUPS_EXCLUDED_FROM_AGGREGATE = (GROUP_DEFINITION_FILE, GROUP_FS_MATRIX_DEFINITION)

# 詳細ファイルの箇条書き記号（③`prof_check.BULLET`と同じ）
BULLET = "・"

# `matrix_check`／`gst_command_check`が○×欄のNGメッセージを組み立てるときの共通の切れ目
# （両モジュールとも`f"「{シート}」シートの{OX_FIELD_LABEL}にデータ不備がある。{problems}"`の形。
# `problems`は`[{where}] {msg}`を" / "で連結した不備箇所の一覧）。
_DEFECT_DETAIL_MARKER = "にデータ不備がある。"


def extract_defect_detail(message):
	"""NGメッセージから不備箇所の部分（`[FI FSmatrix / Z列] Z11 が ○/× ではありません: '※'`）
	だけを取り出す（2026-09-03新規）。

	集約ファイル（`インプットファイルチェック結果.txt`）はチェック②の行末へNG内容を併記するが、
	「「FI FSmatrix」シートの適用（〇/×）欄にデータ不備がある。」という前置きの文はどのNGでも
	同じで、1行に収めるには冗長なため落とす。前置きが無い形のメッセージ
	（「〜シートがない。」「〜を開けない: 」等）はそれ自体が不備の内容なので、そのまま返す。

	引数:
		message: `matrix_check`／`gst_command_check`が返したNGのmessage

	戻り値:
		不備箇所の部分だけの文字列（前置きが無ければmessageそのまま）
	"""
	_prefix, marker, detail = str(message).partition(_DEFECT_DETAIL_MARKER)
	if marker and detail.strip():
		return detail.strip()
	return str(message).strip()


def build_body(groups, timestamp=None):
    """結果ログの本文（文字列）を組み立てる。

    引数:
        groups: [(グループ見出し, [(result, message), ...]), ...]。順序はそのまま出力順に
            なる。行が空のグループは見出しごと省く。resultは現在の書式では使わない
            （OK/NGはmessageの文末で表す）が、将来の書式変更に備えて受け取る。
        timestamp: 見出しに出す実行日時（省略時は現在時刻）。

    戻り値:
        末尾に改行を含む本文文字列。
    """
    stamp = timestamp or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [f"実行日時: {stamp}"]

    for group_name, group_lines in groups:
        if not group_lines:
            continue
        lines.append("")
        lines.append(f"[{group_name}]")
        for _result, message in group_lines:
            # 1行1件を崩さないよう、message中の改行は空白へ潰す。
            lines.append(f"{BULLET}{' '.join(str(message).splitlines())}")

    return "\n".join(lines) + "\n"


def write_report(groups, output_path, encoding="utf-8-sig"):
    """build_bodyの本文を`output_path`へ書き出す。

    既定のencodingは`utf-8-sig`（BOM付きUTF-8）。SharePointへアップロードした結果ログを
    ブラウザで直接開いた際の文字化けを避けるため、①の詳細ファイルと同じ形式に合わせている。
    """
    directory = os.path.dirname(output_path)
    if directory:
        os.makedirs(directory, exist_ok=True)
    with open(output_path, "w", encoding=encoding) as f:
        f.write(build_body(groups))

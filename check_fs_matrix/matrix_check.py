"""
Checks whether an Excel file contains a sheet that is some version of the
"FS Matrix" sheet - the actual sheet name varies across submissions (e.g.
"FI FS Matrix", "FSマトリクス"), so every known variant is kept in one list
here rather than checking a single hardcoded name (this list is broader
than check_fs_ox.py's own DEF_SHEET_CANDIDATES, covering per-domain prefixes
like MOT/BAT/EVCC/ZEV2IN1/AREAFRONT/CORE/ICM it doesn't have).

The range/symbol validation logic below (Strings/Grid/Report,
derive_data_anchor, detect_mark_columns, check_mark_block, etc.) is copied
in directly from check_fs_ox.py's macro-accurate validator
(reverse-engineered from the real VBA macro - modChkExec.ReadFromFS - that
actually consumes these files), rather than imported from that file, so this
module has no dependency on check_fs_ox.py at all. check_fs_ox.py itself is
untouched and still works standalone as its own CLI tool - this is a copy,
not a move.

【メッセージ書式】check_matrix_sheet_parts系が返すmessageは、結果ログ（`check2_report`が
[FI FSマトリクスファイル]見出しの下へ並べる）の箇条書き1行としてそのまま使える文にする。
OK/NGはラベルではなく文末の肯定形／否定形（シートがある↔ない、データ不備がない↔ある）で
表し、NGの場合は続けて不備箇所（セル番地等）を書く。文言を変える場合は
`definition_file_check`／`gst_command_check`側の文と揃えること（2026-09-03改修）。

【2026-09-02改修】unified_main.pyが同梱コピー（本フォルダ）をimportするため、判定条件が
`..\\matrix_check_ref.py`（マクロ準拠版）と食い違わないよう、本ファイルの内容を同ファイルと
同一の実装へ差し替えた。以前の版は「"部番"行の数字始まり列 × "診断識別コード"直下〜列末尾を
○×記号リストと突き合わせるだけ」の簡易判定だったため、マクロが実際に停止・誤読するケース
（部番/仕向け/ソフトの3行の切れ目=FS_TODO_EXEC_EMPTY、診断識別コードのタイトルセル空、
途中空行による静かな打ち切り、キャッシュ値が無い数式、結合セル）を見逃していた。両ファイルは
重複管理のため、改修時は必ず両方へ反映すること。
"""

import csv
import os

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

# Known variants of the FS Matrix sheet name - add more here as new formats
# show up in submitted files.
MATRIX_SHEET_NAME_VARIANTS = [
    "FI FSﾏﾄﾘｸｽ",
    "FI FSﾏﾄﾘｸｽ",
    "FI_FSﾏﾄﾘｸｽ",
    "FI FSmatrix",
    "FI FSマトリクス",
    "FI_FSマトリクス",
    "ＦＩ FSMATRIX",
    "FI Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "MOT FSﾏﾄﾘｸｽ",
    "MOT FSﾏﾄﾘｸｽ",
    "MOT_FSﾏﾄﾘｸｽ",
    "MOT FSmatrix",
    "MOT FSマトリクス",
    "MOT FSMATRIX",
    "MOT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "BAT FSﾏﾄﾘｸｽ",
    "BAT FSﾏﾄﾘｸｽ",
    "BAT_FSﾏﾄﾘｸｽ",
    "BAT FSmatrix",
    "BAT FSマトリクス",
    "BAT FSMATRIX",
    "BAT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "EVCC FSﾏﾄﾘｸｽ",
    "EVCC FSﾏﾄﾘｸｽ",
    "EVCC_FSﾏﾄﾘｸｽ",
    "EVCC FSmatrix",
    "EVCC FSマトリクス",
    "EVCC FSMATRIX",
    "EVCC Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "ZEV2IN1 FSﾏﾄﾘｸｽ",
    "ZEV2IN1 FSﾏﾄﾘｸｽ",
    "ZEV2IN1_FSﾏﾄﾘｸｽ",
    "ZEV2IN1 FSmatrix",
    "ZEV2IN1 FSマトリクス",
    "ZEV2IN1 FSMATRIX",
    "ZEV2IN1 Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "AREAFRONT FSﾏﾄﾘｸｽ",
    "AREAFRONT FSﾏﾄﾘｸｽ",
    "AREAFRONT_FSﾏﾄﾘｸｽ",
    "AREAFRONT FSmatrix",
    "AREAFRONT FSマトリクス",
    "AREAFRONT FSMATRIX",
    "AREAFRONT Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "CORE FSﾏﾄﾘｸｽ",
    "CORE FSﾏﾄﾘｸｽ",
    "CORE_FSﾏﾄﾘｸｽ",
    "CORE FSmatrix",
    "CORE FSマトリクス",
    "CORE FSMATRIX",
    "CORE Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",
    "ICM FSﾏﾄﾘｸｽ",
    "ICM FSﾏﾄﾘｸｽ",
    "ICM_FSﾏﾄﾘｸｽ",
    "ICM FSmatrix",
    "ICM FSマトリクス",
    "ICM_FSマトリクス",
    "ＩＣＭ FSMATRIX",
    "ICM Fail Safe Matrix",
    "FSﾏﾄﾘｸｽ",
    "FSマトリクス",

]

# --- Copied from check_fs_ox.py: macro-side constants -------------------
MAX_CLM = 256                       # NUM_MAX_CLM
LOG_NAME = 'FsMatrix_Info.log'      # STR_SYS_LOG
LABEL_SHIKIBETU = '診断識別コード'
# データ表本体の最小幅。診断識別コード(+0) から SCS(+9) までがデータ列なので、
# ソフト列（○/× 列）はこれより右にしか存在しない。
MIN_DATA_WIDTH = 9                  # NUM_IN_UDS_SCS

# --- MacroStringTable.csv が無い場合の既定値 ---
DEF_OX_O = '○'                      # GBL_OX_O      U+25CB
DEF_OX_X = '×'                      # GBL_OX_X      U+00D7
DEF_O_SUB = '◯●〇０0ＯｏOo'          # GBL_OX_O_SUB
DEF_X_SUB = 'ＸｘXx-ー－‐'            # GBL_OX_X_SUB
DEF_REMOVE_BUBAN = '37820-,28100-,5K800-,5K801-,ZK800-,ZK801-'
DEF_REMOVE_SOFT = ('37805-,37806-,28100-,28101-,28102-,'
                   '5K800-,5K801-,ZK800-,ZK801-')
DEF_SHEET_CANDIDATES = [
    'FI FSﾏﾄﾘｸｽ', 'FI_FSﾏﾄﾘｸｽ', 'FI FSmatrix', 'FI FSマトリクス',
    'FI_FSマトリクス', 'ＦＩ FSMATRIX', 'FI Fail Safe Matrix',
    'FSﾏﾄﾘｸｽ', 'FSマトリクス',
]


class Strings:
    """MacroStringTable.csv（あれば）から定数を読む。"""

    def __init__(self, path=None):
        self.tbl = {}
        if path and os.path.isfile(path):
            with open(path, 'r', encoding='cp932', errors='replace', newline='') as f:
                for row in csv.reader(f):
                    if len(row) >= 2 and row[0]:
                        self.tbl.setdefault(row[0].strip(), row[1])
            self.source = path
        else:
            self.source = None

    def get(self, key, default=''):
        return self.tbl.get(key, default)

    @property
    def ox_o(self):
        return self.get('GBL_OX_O', DEF_OX_O)

    @property
    def ox_x(self):
        return self.get('GBL_OX_X', DEF_OX_X)

    @property
    def o_sub(self):
        return self.get('GBL_OX_O_SUB', DEF_O_SUB)

    @property
    def x_sub(self):
        return self.get('GBL_OX_X_SUB', DEF_X_SUB)

    def remove_buban(self):
        return [s for s in self.get('FS_DWG_REMOVE_BUBAN',
                                    DEF_REMOVE_BUBAN).split(',') if s]

    def remove_soft(self):
        return [s for s in self.get('FS_DWG_REMOVE_SOFT',
                                    DEF_REMOVE_SOFT).split(',') if s]

    def sheet_candidates(self):
        if not self.tbl:
            return list(DEF_SHEET_CANDIDATES)
        out = []
        for n in [''] + [str(i) for i in range(2, 11)]:
            v = self.get('SHEET_FSMATRIX' + n)
            if v and v not in out:
                out.append(v)
        return out or list(DEF_SHEET_CANDIDATES)


class Grid:
    """VBA の .Value / IsEmpty / MergeArea を再現するシートラッパ。"""

    def __init__(self, ws_val, ws_raw):
        self._v = ws_val    # data_only=True  : キャッシュ値
        self._r = ws_raw    # data_only=False : 数式
        self.title = ws_val.title
        self.max_row = max(ws_val.max_row or 1, ws_raw.max_row or 1)
        self.max_col = max(ws_val.max_column or 1, ws_raw.max_column or 1)
        self.merged_followers = set()
        self.merge_of = {}          # (row, col) -> (anchor_row, anchor_col, max_row)
        for rng in ws_val.merged_cells.ranges:
            for row in range(rng.min_row, rng.max_row + 1):
                for col in range(rng.min_col, rng.max_col + 1):
                    self.merge_of[(row, col)] = (rng.min_row, rng.min_col,
                                                 rng.max_row)
                    if (row, col) != (rng.min_row, rng.min_col):
                        self.merged_followers.add((row, col))

    def formula(self, row, col):
        if row < 1 or col < 1:
            return None
        f = self._r.cell(row=row, column=col).value
        return f if isinstance(f, str) and f.startswith('=') else None

    def _raw(self, row, col):
        if row < 1 or col < 1:
            return None
        return self._v.cell(row=row, column=col).value

    def text(self, row, col):
        """VBA の .Value を文字列化（未入力は ''）。トリムしない。"""
        if row < 1 or col < 1 or (row, col) in self.merged_followers:
            return ''
        return self._to_text(self._raw(row, col))

    def merge_text(self, row, col):
        """.MergeArea.Cells(1).Value 相当（結合の左上を読む）。"""
        anchor = self.merge_of.get((row, col))
        if anchor:
            return self._to_text(self._raw(anchor[0], anchor[1]))
        return self.text(row, col)

    @staticmethod
    def _to_text(v):
        if v is None:
            return ''
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        return str(v)

    def has_content(self, row, col):
        """IsEmpty が False になるか（数式は結果が "" でも空でない）。"""
        if row < 1 or col < 1 or (row, col) in self.merged_followers:
            return False
        if self.formula(row, col) is not None:
            return True
        return self._raw(row, col) is not None

    def is_error(self, row, col):
        v = self._raw(row, col)
        return isinstance(v, str) and v.startswith('#') and v.endswith('!')

    def unevaluated(self, row, col):
        """数式なのにキャッシュ値が無い＝このツールでは値を判定できない。

        Excel はファイルを開くと再計算するのでマクロ側は正しい値を読むが、
        openpyxl はキャッシュしか見られないため判定を保留する必要がある。
        """
        return (self.formula(row, col) is not None
                and self._raw(row, col) is None)

    def merge_bottom(self, row, col):
        anchor = self.merge_of.get((row, col))
        return anchor[2] if anchor else row

    def a1(self, row, col):
        return '%s%d' % (get_column_letter(col), row)


class Report:
    def __init__(self):
        self.ng = []
        self.warn = []
        self.info = []

    def add_ng(self, where, msg):
        self.ng.append((where, msg))

    def add_warn(self, where, msg):
        self.warn.append((where, msg))

    def add_info(self, msg):
        self.info.append(msg)


def describe(value):
    """同形異字・空白混入を目視で判別できる形にする。"""
    if value == '':
        return '(空欄)'
    return '%r' % (value,)


def classify(value, st):
    """マクロの判定を再現。

    ChkArrayDat で ○/× なら即OK。違えば InStr で
    GBL_OX_O_SUB / GBL_OX_X_SUB の「部分文字列」として探す。
    空欄は明示的に除外（マクロ側も同じガード）。
    戻り値: 'canonical' / 'sub' / 'ng'
    """
    if value == '':
        return 'ng'
    if value == st.ox_o or value == st.ox_x:
        return 'canonical'
    # InStr(1, strSubOs, strYryrMark, vbBinaryCompare) 相当（部分文字列判定）
    if value in st.o_sub or value in st.x_sub:
        return 'sub'
    return 'ng'


# ----------------------------------------------------------------------
# アンカー
# ----------------------------------------------------------------------
def read_log(path, rep):
    """FsMatrix_Info.log を読む。戻り値 dict（読めなければ None）。"""
    if not os.path.isfile(path):
        return None
    out = {}
    with open(path, 'r', encoding='cp932', errors='replace', newline='') as f:
        for row in csv.reader(f):
            if len(row) >= 2 and row[0]:
                out[row[0].strip()] = row[1].strip()
    need = ('FS_CLMS', 'FS_BUBAN_ROW', 'FS_SIMUKE_ROW', 'FS_DWG_ROW',
            'FS_DATA_ROW', 'FS_DATA_CLM')
    if not all(k in out for k in need):
        rep.add_warn(LOG_NAME, '%s の項目が不足しているため無視します。' % path)
        return None
    return out


def resolve_sheet(wb_names, st, override):
    if override:
        return override if override in wb_names else None
    for name in st.sheet_candidates():
        if name in wb_names:
            return name
    return None


def derive_data_anchor(g):
    """診断識別コード ラベルから データ開始行/列 を導出。"""
    for row in range(1, min(g.max_row, 60) + 1):
        for col in range(1, min(g.max_col, 12) + 1):
            if g.text(row, col).replace('\n', '').strip() == LABEL_SHIKIBETU:
                return g.merge_bottom(row, col) + 1, col, row
    return None, None, None


def find_end_row(g, data_row, data_col):
    """診断識別コードが空になる直前の行（マクロの Exit Do 位置 - 1）。"""
    row = data_row
    while row <= g.max_row and g.has_content(row, data_col):
        row += 1
    return row - 1


def detect_mark_columns(g, data_row, end_row, data_col, st):
    """マーク密度から ○/× 列ブロックを推定（アンカー行が不明なとき）。

    密度は「正規の ○/×」だけで数える。補助リストには '-' が含まれるため、
    Readiness Group や RGID のような '-' 埋めの列まで拾ってしまうのを避ける。
    ソフト列は正規の ○/× がほぼ100%を占めるので、これで明確に分離できる。
    """
    n_rows = end_row - data_row + 1
    if n_rows < 5:
        return []
    hits = []
    for col in range(data_col + 1, min(g.max_col, MAX_CLM) + 1):
        n_ok = 0
        for row in range(data_row, end_row + 1):
            if classify(g.text(row, col), st) == 'canonical':
                n_ok += 1
        if n_ok >= n_rows * 0.5:
            hits.append(col)
    # 連続ブロックのうち最長のものを採用（同長なら右側＝ソフト列側を優先）
    best, cur = [], []
    for col in hits:
        if cur and col == cur[-1] + 1:
            cur.append(col)
        else:
            cur = [col]
        if len(cur) >= len(best):
            best = list(cur)
    return best


def mark_columns_from_headers(g, rep, clms, brow, srow, drow):
    """マクロと同じ規則: 3行が埋まっている連続列を clms から右へ。"""
    cols = []
    for col in range(clms, min(g.max_col, MAX_CLM) + 1):
        missing = [lbl for lbl, r in (('部番', brow), ('仕向け', srow), ('ソフト', drow))
                   if not g.has_content(r, col)]
        if missing:
            break
        cols.append(col)
        for lbl, r in (('部番', brow), ('仕向け', srow), ('ソフト', drow)):
            if g.is_error(r, col):
                rep.add_warn('%s / %s列' % (g.title, get_column_letter(col)),
                             '%s (%s行) がエラー値です。この列は3点一致の対象外に'
                             'なります。' % (g.a1(r, col), lbl))
    return cols


# ----------------------------------------------------------------------
# 検査
# ----------------------------------------------------------------------
def header_rows(brow, srow, drow):
    return (('部番', brow), ('仕向け', srow), ('ソフト', drow))


def find_gap(g, rep, cols, brow, srow, drow):
    """3行の最初の切れ目と、その右に残るソフト列を返す。

    マクロは clms から右へ走査し、3行のどれかが空のセルに当たった時点で
    FS_TODO_EXEC_EMPTY で停止する（それより前で一致すれば Exit For で抜ける）。
    したがって切れ目は「その右にソフト情報が残っている場合だけ」問題になる。
    切れ目の右にある列は選択しても到達できない。
    戻り値: (gap_col, missing_labels, [到達不能な列]) / 切れ目なしなら None
    """
    if not cols:
        return None
    limit = min(g.max_col, MAX_CLM)
    rows = header_rows(brow, srow, drow)
    gap_col = None
    for col in range(cols[0], limit + 1):
        missing = [lbl for lbl, r in rows if not g.has_content(r, col)]
        if missing:
            gap_col, missing_labels = col, missing
            break
    if gap_col is None:
        return None
    beyond = [c for c in range(gap_col + 1, limit + 1)
              if any(g.has_content(r, c) for _, r in rows)]
    return gap_col, missing_labels, beyond


def check_header_formulas(g, rep, cols, brow, srow, drow):
    """部番/仕向け/ソフト行が「値の無い数式」なら判定を保留する。

    これらは他シート参照の数式であることが多い（例: =ｼｽﾃﾑﾏﾄﾘｸｽ!G3）。
    Excel は開いた時に再計算するのでマクロは正しく読むが、
    キャッシュが空のファイルではこのツールが3点一致を判定できない。
    """
    bad = [g.a1(r, c) for _, r in header_rows(brow, srow, drow)
           for c in cols if g.unevaluated(r, c)]
    if bad:
        rep.add_warn('%s / ヘッダ3行' % g.title,
                     '%s などが数式でキャッシュ値が無いため、3点一致と切れ目の判定が'
                     'できません（%d セル）。Excel で開いて保存し直してから'
                     '再実行してください。'
                     % (', '.join(bad[:5]), len(bad)))
        return False
    return True


def check_header_rows(g, rep, cols, brow, srow, drow, matched_col=None):
    """3行の切れ目を検査する。"""
    gap = find_gap(g, rep, cols, brow, srow, drow)
    if gap is None:
        return
    gap_col, missing, beyond = gap
    if not beyond:
        return              # 表の右端。マクロは一致済みで抜けているので問題なし
    where = '%s / %s列' % (g.title, get_column_letter(gap_col))
    cells = ', '.join('%s(%s行)' % (g.a1(r, gap_col), lbl)
                      for lbl, r in header_rows(brow, srow, drow)
                      if lbl in missing)
    msg = ('%s が空欄です。マクロはこの列で FS_TODO_EXEC_EMPTY'
           '（選択部番・仕向け・ソフトウェア情報が見つかりません）として停止するため、'
           '右側の %s 列のソフトは選択できません。'
           % (cells, ', '.join(get_column_letter(c) for c in beyond)))
    if matched_col is not None and matched_col < gap_col:
        rep.add_info('%s: %s（今回の対象列 %s は切れ目より左なので今回は影響なし）'
                     % (where, msg, get_column_letter(matched_col)))
    else:
        rep.add_ng(where, msg)


def check_shikibetu_column(g, rep, data_row, data_col, end_row):
    """診断識別コード列の空欄（静かな打ち切り）とタイトルセルを検査。"""
    where = '%s / 診断識別コード列(%s)' % (g.title, get_column_letter(data_col))

    title = g.merge_text(data_row - 1, data_col)
    if title == '':
        rep.add_ng(where,
                   '%s のタイトルセルが空です（結合範囲を辿っても空）。'
                   'FS_TODO_EXEC_SHIKIBETU_TITLE_ERR になります。'
                   % g.a1(data_row - 1, data_col))

    # end_row より下に、離れてデータが残っていないか（＝途中に空行がある）
    row = end_row + 2
    limit = min(g.max_row, end_row + 5000)
    while row <= limit:
        if g.has_content(row, data_col):
            rep.add_warn(where,
                         '%s 以降にも診断識別コードがありますが、%s が空欄のため'
                         'マクロは %d行で読み込みを打ち切り、それ以降を無警告で'
                         '無視します。'
                         % (g.a1(row, data_col), g.a1(end_row + 1, data_col),
                            end_row))
            break
        row += 1


def check_mark_block(g, rep, data_row, end_row, cols, st, strict, verbose):
    """○/× 範囲の本体チェック。"""
    for col in cols:
        where = '%s / %s列' % (g.title, get_column_letter(col))
        if verbose:
            rep.add_info('%s: 検証範囲 %s:%s'
                         % (where, g.a1(data_row, col), g.a1(end_row, col)))
        n_sub = 0
        for row in range(data_row, end_row + 1):
            if g.unevaluated(row, col):
                rep.add_warn(where,
                             '%s は数式でキャッシュ値が無いため判定できません。'
                             'Excel で開いて保存し直してから再実行してください。'
                             % g.a1(row, col))
                continue
            value = g.text(row, col)
            kind = classify(value, st)

            if kind == 'ng':
                rep.add_ng(where,
                           '%s が ○/× ではありません: %s'
                           % (g.a1(row, col), describe(value)))
            elif kind == 'sub':
                n_sub += 1
                target = st.ox_o if value in st.o_sub else st.ox_x
                msg = ('%s は %s に自動正規化されます: %s'
                       % (g.a1(row, col), target, describe(value)))
                if strict:
                    rep.add_ng(where, msg + ' （--strict）')
                else:
                    rep.add_info('%s: %s' % (where, msg))
                if len(value) > 1:
                    rep.add_warn(where,
                                 '%s は複数文字ですが補助リストの部分文字列として'
                                 '通ってしまいます: %s'
                                 % (g.a1(row, col), describe(value)))

            if g.formula(row, col) is not None:
                rep.add_warn(where,
                             '%s は数式です（適用欄は直値が前提）。キャッシュ値が'
                             '古いとマクロは意図しない値を読みます。' % g.a1(row, col))
            if (row, col) in g.merged_followers:
                rep.add_warn(where,
                             '%s は結合セルの2番目以降のため Empty として読まれ、'
                             'エラーになります。' % g.a1(row, col))
        if n_sub and not strict:
            rep.add_info('%s: 正規化される値 %d 件（表記統一を推奨。'
                         'CommandSupportGST 側のマクロはこの正規化を行いません）'
                         % (where, n_sub))


def check_workbook(xlsx_path, sheet=None, log_path=None, strings_path=None, strict=False):
    """Programmatic entry point - runs the anchor-derivation and 適用
    （○/×）欄 validation above, minus the interactive --buban/--simuke/--soft
    single-column targeting (that narrows to one column by hand; this
    checks every ○/× column found instead). Uses FsMatrix_Info.log /
    MacroStringTable.csv if found next to the file, otherwise derives
    everything from the sheet itself. Returns the resulting Report: rep.ng
    entries mean the macro would error out and stop; rep.warn entries mean
    it would run but silently produce a wrong result - callers should
    usually treat both as a real problem."""
    rep = Report()

    if strings_path is None:
        # 本ファイルは check_fs_matrix\ 配下の同梱コピーなので、MacroStringTable.csv は
        # 自フォルダだけでなく1つ上（ツール本体フォルダ = ..\matrix_check_ref.py と同じ位置）も
        # 探す。これが無いと ..\ に置かれたテーブルを ref 版だけが読み、判定条件がずれる。
        here = os.path.dirname(os.path.abspath(__file__))
        parent = os.path.dirname(here)
        for cand in (os.path.join(here, 'systemfiles', 'MacroStringTable.csv'),
                     os.path.join(here, 'MacroStringTable.csv'),
                     os.path.join(parent, 'systemfiles', 'MacroStringTable.csv'),
                     os.path.join(parent, 'MacroStringTable.csv')):
            if os.path.isfile(cand):
                strings_path = cand
                break
    st = Strings(strings_path)

    wb_val = load_workbook(xlsx_path, data_only=True)
    wb_raw = load_workbook(xlsx_path, data_only=False)

    resolved_sheet = resolve_sheet(wb_val.sheetnames, st, sheet)
    if resolved_sheet is None:
        rep.add_ng('(sheet)',
                    'FSマトリクスのシートが見つかりません（候補: %s / 実在: %s）。'
                    % (', '.join(st.sheet_candidates()), ', '.join(wb_val.sheetnames)))
        return rep
    g = Grid(wb_val[resolved_sheet], wb_raw[resolved_sheet])

    if log_path is None:
        d = os.path.dirname(os.path.abspath(xlsx_path))
        for cand in (os.path.join(d, LOG_NAME), os.path.join(d, 'systemfiles', LOG_NAME)):
            if os.path.isfile(cand):
                log_path = cand
                break
    log = read_log(log_path, rep) if log_path else None

    data_row = data_col = None
    brow = srow = drow = clms = None
    if log:
        try:
            clms = int(log['FS_CLMS'])
            brow = int(log['FS_BUBAN_ROW'])
            srow = int(log['FS_SIMUKE_ROW'])
            drow = int(log['FS_DWG_ROW'])
            data_row = int(log['FS_DATA_ROW'])
            data_col = int(log['FS_DATA_CLM'])
            name = log.get('FS_FILE_NAME', '').strip('"')
            if name and name != os.path.basename(xlsx_path):
                rep.add_warn(LOG_NAME,
                             'ログの対象ファイルは %r です。今回の %r とは別のため、'
                             'マクロはログを破棄してセル選択を再度求めます。'
                             % (name, os.path.basename(xlsx_path)))
        except (ValueError, KeyError):
            rep.add_warn(LOG_NAME, 'ログの数値が不正なため導出に切り替えます。')
            log = None

    if data_row is None:
        data_row, data_col, _label_row = derive_data_anchor(g)
        if data_row is None:
            rep.add_ng(g.title,
                        '「%s」ラベルが見つからず、データ開始セルを特定できません。' % LABEL_SHIKIBETU)
            return rep

    end_row = find_end_row(g, data_row, data_col)
    if end_row < data_row:
        rep.add_ng(g.title, '%s が空欄のためデータ行がありません。' % g.a1(data_row, data_col))
        return rep

    dcols = detect_mark_columns(g, data_row, end_row, data_col + MIN_DATA_WIDTH, st)

    if brow and srow and drow:
        if clms is None:
            clms = dcols[0] if dcols else None
        if clms is None:
            for c in range(data_col + MIN_DATA_WIDTH + 1, min(g.max_col, MAX_CLM) + 1):
                if all(g.has_content(r, c) for r in (brow, srow, drow)):
                    clms = c
                    break
        if clms is None:
            rep.add_ng(g.title, '部番/仕向け/ソフト の3行が揃う列が見つかりません。')
            return rep
        cols = mark_columns_from_headers(g, rep, clms, brow, srow, drow)
    else:
        cols = dcols

    if not cols:
        rep.add_ng(g.title, '○/× 列を特定できませんでした。')
        return rep

    if brow and srow and drow:
        usable = check_header_formulas(g, rep, cols, brow, srow, drow)
        if usable:
            check_header_rows(g, rep, cols, brow, srow, drow)

    check_shikibetu_column(g, rep, data_row, data_col, end_row)
    check_mark_block(g, rep, data_row, end_row, cols, st, strict, verbose=False)

    return rep


# 結果ログで「ツールが想定しているシート」を指す代表名。実在シート名は提出物ごとに
# MATRIX_SHEET_NAME_VARIANTSのいずれかで揺れるため、見つかった場合はその実名を、
# 1つも見つからなかった場合はこの代表名を文中に出す。
CANONICAL_SHEET_NAME = "FI FSmatrix"

# 適用欄の呼び方（②の結果ログ共通。`gst_command_check`側と揃えること）。
OX_FIELD_LABEL = "適用（〇/×）欄"


def check_matrix_sheet_parts(local_path):
    """Returns [(result, message)] - the FS Matrix file's report bullets, in
    the order they appear under the [FI FSマトリクスファイル] heading:

      1. シートの有無（ツールが想定している「FI FSmatrix」シートがある／ない）
      2. その適用（〇/×）欄のデータ不備の有無（NGなら不備箇所を続けて書く）

    Finds whichever MATRIX_SHEET_NAME_VARIANTS name exists as a sheet in
    local_path, reports that as found, then delegates the range/symbol
    validation to check_workbook (passing the already-matched variant as its
    `sheet` argument, so it doesn't need to re-resolve the sheet name
    against its own smaller candidate list). A WARN there counts as NG,
    since it means the macro would silently produce a wrong result rather
    than stopping outright.

    シートが無い／ファイルを開けない場合は適用欄を検査できないので、2行目は返さない
    （返る行数は1〜2行。呼び出し側はリストを走査して扱うこと）。"""
    try:
        wb = load_workbook(local_path, data_only=True)
    except Exception as e:
        return [("NG", f"「{os.path.basename(local_path)}」を開けない: {e}")]

    variant = next((name for name in MATRIX_SHEET_NAME_VARIANTS if name in wb.sheetnames), None)
    if variant is None:
        return [
            (
                "NG",
                f"ツールが想定している「{CANONICAL_SHEET_NAME}」シートがない。"
                f"（このファイルにあるシート: {', '.join(wb.sheetnames)}）",
            )
        ]

    parts = [("OK", f"ツールが想定している「{variant}」シートがある。")]

    try:
        rep = check_workbook(local_path, sheet=variant)
    except Exception as e:
        parts.append(("NG", f"「{variant}」シートの{OX_FIELD_LABEL}にデータ不備がないか確認できない。検証中にエラー: {e}"))
        return parts

    if rep.ng or rep.warn:
        problems = " / ".join(f"[{where}] {msg}" for where, msg in (rep.ng + rep.warn))
        parts.append(("NG", f"「{variant}」シートの{OX_FIELD_LABEL}にデータ不備がある。{problems}"))
    else:
        parts.append(("OK", f"「{variant}」シートの{OX_FIELD_LABEL}にデータ不備がない。"))

    return parts


def check_matrix_sheet_parts_for_all_files(local_paths):
    """Runs check_matrix_sheet_parts on every given local file path and
    returns all of their bullets as one flat [(result, message)] list - for a
    column that may have more than one file attached, every one of them needs
    a matching sheet.

    ファイルが2つ以上ある場合だけ、どのファイルの行なのか分かるよう各行の先頭へ
    ファイル名を付ける（1ファイルなら提出物は1つに決まるので付けない）。"""
    parts = []
    for local_path in local_paths:
        for result, message in check_matrix_sheet_parts(local_path):
            if len(local_paths) > 1:
                message = f"「{os.path.basename(local_path)}」: {message}"
            parts.append((result, message))
    return parts


def _combine(parts):
    """箇条書きの行リストを、SharePoint列へ書き戻す1件の(result, message)へまとめる
    （②`unified_result.combine_results`と同じ「1つでもNGなら最初のNGを返す」規則）。"""
    for result, message in parts:
        if result != "OK":
            return "NG", message
    return "OK", " / ".join(message for _, message in parts if message)


def check_matrix_sheet(local_path):
    """Returns one combined (result, message) for local_path - the same
    checks as check_matrix_sheet_parts, merged into a single result for
    callers that can only hold one (e.g. the SharePoint write-back)."""
    return _combine(check_matrix_sheet_parts(local_path))


def check_matrix_sheet_for_all_files(local_paths):
    """Runs the FS Matrix checks on every given local file path and returns
    one combined (result, message): OK only if every file passed; otherwise
    NG with whichever bullet failed first."""
    return _combine(check_matrix_sheet_parts_for_all_files(local_paths))

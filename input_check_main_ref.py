"""
Sample script - a focused demo of the core idea, not the full periodic
service: go to SharePoint, find items whose status column is "submitted",
download each one's file, and check only whether it contains a sheet
named "元表" (the reference "base table" sheet used in the real FI FS
Matrix check). Returns OK if found, NG with a message if not.

basic idea for talking to SharePoint: MSAL device-code sign-in with a locally
cached token, then plain SharePoint REST API calls with that token as a
Bearer header.

Network calls go through ntlm_proxy.py, which builds an NTLM-authenticated
tunnel through this network's corporate proxy - a plain unauthenticated
HTTPS_PROXY setting gets a 407 Proxy Authentication Required error here.
That module is its own code (adapted from sp_helper.py's proven approach,
not an import of that file), used both for msal's http_client and for every
SharePoint REST call below instead of the plain requests library.

Setup
-----
1. Fill in SITE_URL / LIST_GUID / SUBMIT_STATUS_FIELD_INTERNAL_NAME below
   with your real SharePoint site, list, and status column (currently
   dummy placeholders - see the TODO comments).

2. Run:
     input_check_main.py
   The first run prints a device-code sign-in URL and code - open it in a
   browser and sign in with an account that has access to the site. After
   that, the token is cached locally (token_cache.json) and reused
   silently on later runs.
"""

import os
import re
import sys
import tempfile
from urllib.parse import urlparse

import msal

import check2_report
import definition_file_check
import gst_command_check
import matrix_check
import ntlm_proxy

# ---- CONFIG ----------------------------------------------------------
# TODO: dummy value - replace with your org's real registered app's client ID
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
AUTHORITY = "https://login.microsoftonline.com/organizations"

# TODO: dummy link - replace with your real SharePoint site
SITE_URL = "https://globalhonda.sharepoint.com/sites/jphgt105596" #XPX
# TODO: dummy value - replace with the real list GUID backing your FI FS Matrix submissions library
LIST_GUID = "021F8116-DDD4-4E22-B982-4B738D2BD0CD"

# TODO: dummy placeholder - replace with the real internal name of your submission-status column
SUBMIT_STATUS_FIELD_INTERNAL_NAME = "SubmitStatus"
SUBMIT_STATUS_VALUE = "Submitted"

# TODO: using the display name "ステータス" directly as the internal name for
# now - replace with the real internal name later if it turns out to differ
# (same SubmitStatus/FilePath mismatch risk as elsewhere in this file).
STATUS_FIELD_INTERNAL_NAME = "OData__x30b9__x30c6__x30fc__x30bf__x30"#SPの　ステータス　列
STATUS_VALUE = "3.本申請確認"

# TODO: dummy placeholder - replace with the real internal name of the column
# storing the file link, e.g. <p><a href=".../Practice.xlsx">Practice.xlsx</a></p>
FILE_PATH_FIELD_INTERNAL_NAME = "commentSystemMatrix"

# Built-in per-item column every SharePoint list has - its internal name is
# reliably "Title" (not a mangled display-name encoding, unlike most fields
# above), used to name each item's own report file below.
TITLE_FIELD_INTERNAL_NAME = "Title"

# TODO: dummy placeholders - confirm these match the real internal names of
# these two columns (display name and internal name can differ - see the
# SubmitStatus/FilePath mismatches hit earlier)
RESULT_FIELD_INTERNAL_NAME = "InputFileCheckResult"#InputFileチェック結果
#MESSAGE_FIELD_INTERNAL_NAME = "InputFile_x306e__x30c1__x30a7__x"#InputFileのチェックの詳細メッセージ

# Confirmed via the column's Field= URL parameter (lowercase "i", unlike
# the display name) - the uploaded report file's link gets written here as
# an HTML href, same format as FILE_PATH_FIELD_INTERNAL_NAME etc.
RESULT_LOG_FIELD_INTERNAL_NAME = "inputFileCheckResultLog"

# ①(check_mao)とのマージに伴い追加（2026-09-02新規）。`InputFileCheckResult`列は②専用ではなく
# ①②③統合の判定サマリ列のため、②はこの行だけを担当する（`PMOチェック自動化マニュアル.pptx`
# 8ページの表示イメージに合わせた文言）。
CHECK2_LINE_PREFIX = "チェック②"
CHECK2_LABEL = f"{CHECK2_LINE_PREFIX}（FSマトリクス↔定義ファイルチェック）"

# `InputFileCheckResult`列内の「チェック①」〜「チェック⑳」で始まる行を検出するパターン
# （2026-09-02新規）。マニュアル8ページは①②③を番号順に固定表示する前提のため、マージ後の
# 行順を本パターンに基づく`sort_key`で①②③…の番号順に安定ソートする。①(check_mao)の
# `result.sort_key`と同じ実装だが、①②はリポジトリを分けているため、この小さなロジックのみ
# 複製している（③以降が増えても丸数字がUnicodeで連続コードポイントのため無改修で対応できる）。
_CHECK_LINE_PATTERN = re.compile(r"^チェック([①-⑳])")


def sort_key(line):
    """`InputFileCheckResult`列の1行から、①②③…の順序を決めるソートキーを返す。
    想定外の行（プレフィックス不一致）は末尾へ回す（`sorted`は安定ソートのため、
    想定外行同士の相対順は入力順のまま保持される）。"""
    matched = _CHECK_LINE_PATTERN.match(line)
    if matched:
        return (0, ord(matched.group(1)) - ord("①") + 1)
    return (1, 0)


# ②の統合結果は①(check_mao)と同じETag楽観的並行制御で書き戻すため、412 Precondition Failed時の
# 読み直し・再マージ・再書き込みのリトライ上限（2026-09-02新規。①の`sp_main._MERGE_RETRY_LIMIT`
# と同じ値）。
_MERGE_RETRY_LIMIT = 5

# Three more link columns, same HTML-link format as FILE_PATH_FIELD_INTERNAL_NAME,
# each holding a different related file - only kept if its file name matches
# the expected pattern for that column (see the is_*_file() functions below).
HEX_FI_FIELD_INTERNAL_NAME = "commentHexFi"
COMMAND_SUPPORT_FIELD_INTERNAL_NAME = "commentCommandSupport"
CAN_FIELD_INTERNAL_NAME = "commentCan"

DOWNLOAD_DIR = "./テキスト"
CACHE_FILE = "./token_cache.json"

# TODO: dummy link - replace with the real server-relative path of the
# second fixed reference file ("sharepoint/excel.xlsm") - this also lives on
# SharePoint (not tied to any particular submission) and its 元表 sheet gets
# compared against SERVER_REFERENCE_FILE_PATH's 元表 in definition_file_check.py
#
# 【2026-09-02改修で導入したSharePointダウンロード方式（resolve_sharepoint_reference_file）は
# 同日中に取消】①(check_mao)とのマージ作業で一度実装したが、実際のSharePoint上のサーバー相対
# パスが不明なままだったため、ユーザー判断により取消し、元の個人ローカルコピー直接参照へ戻した。
# 本番用ファイルへの変更は②担当者に委ねる。
SHAREPOINT_REFERENCE_FILE_PATH = r"C:\Users\RJ067219\OneDrive - Honda\デスクトップ\work\2026_タスク\08_タスク\Check_PythonCode\FI-FailSafeMatrix元表_SP.xlsx" #SPからの元表ファイル

# ①(check_mao)とのマージに伴い、②の出力ファイルを①と同じ「集約ファイル（固定名・リンクあり・
# ①②③統合サマリ）＋詳細ファイル（固定名・リンクなし・②専用の明細）」の2階層構成に変更した
# （2026-09-02改修）。旧実装は詳細レポートに`_{item Title}`が付く可変ファイル名だったが、
# 案件フォルダ（`01_INPUT`）ごとにアップロードするため固定名で問題ない（案件をまたいで
# ファイル名が競合することはない）。

# ②専用の詳細結果（[FS Matrix]／[GSTCommand]／[元表]等の明細）。案件の`01_INPUT`へアップロード
# するが、`inputFileCheckResultLog`からはリンクしない（①の`sp_main.DETAIL_FILE_NAME`と同じ契約）。
DETAIL_FILE_NAME = "FSマトリクス_定義ファイルチェック結果.txt"
DETAIL_LOCAL_PATH = os.path.join(DOWNLOAD_DIR, DETAIL_FILE_NAME)

# ①②③統合の判定サマリ（`InputFileCheckResult`列のマージ後値）を書き込む集約ファイル。
# ①(check_mao)の`sp_main.AGGREGATE_FILE_NAME`と同名・同じ内容形式のため、①②のどちらが後に
# 走っても同じファイル名を指し、`inputFileCheckResultLog`のリンク先が競合しない（2026-09-02新規）。
AGGREGATE_FILE_NAME = "インプットファイルチェック結果.txt"
AGGREGATE_LOCAL_PATH = os.path.join(DOWNLOAD_DIR, AGGREGATE_FILE_NAME)
# ------------------------------------------------------------------------


def get_access_token():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())

    app = msal.PublicClientApplication(
        CLIENT_ID, authority=AUTHORITY, token_cache=cache, http_client=ntlm_proxy.NtlmSession()
    )

    site_host = SITE_URL.split("://", 1)[1].split("/", 1)[0]
    scopes = [f"https://{site_host}/.default"]

    result = None
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            raise RuntimeError(f"device flow failed to start: {flow.get('error_description')}")
        print(flow["message"])  # sign-in URL and code
        result = app.acquire_token_by_device_flow(flow)

    if cache.has_state_changed:
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            f.write(cache.serialize())

    if "access_token" not in result:
        raise RuntimeError(f"sign-in failed: {result.get('error_description')}")
    return result["access_token"]


def get_submitted_items(access_token):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
    }
    url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items"
    select_fields = ",".join([
        "Id",
        TITLE_FIELD_INTERNAL_NAME,
        FILE_PATH_FIELD_INTERNAL_NAME,
        SUBMIT_STATUS_FIELD_INTERNAL_NAME,
        STATUS_FIELD_INTERNAL_NAME,
        HEX_FI_FIELD_INTERNAL_NAME,
        COMMAND_SUPPORT_FIELD_INTERNAL_NAME,
        CAN_FIELD_INTERNAL_NAME,
    ])
    filter_query = (
        f"{SUBMIT_STATUS_FIELD_INTERNAL_NAME} eq '{SUBMIT_STATUS_VALUE}'"
        f" and {STATUS_FIELD_INTERNAL_NAME} eq '{STATUS_VALUE}'"
    )
    params = {
        "$select": select_fields,
        "$filter": filter_query,
        "$top": "5000",
    }
    resp = ntlm_proxy.http_request("GET", url, headers=headers, params=params)
    resp.raise_for_status()
    return resp.json()["d"]["results"]


def extract_all_file_urls(file_path_html):
    """FilePath-style columns are stored as HTML links, e.g.
    <p><a href=".../Practice.xlsx">Practice.xlsx</a></p> - possibly more
    than one per column. This pulls every href out (not just the first),
    reducing absolute URLs to the server-relative path
    GetFileByServerRelativeUrl expects."""
    hrefs = re.findall(r'href=["\']([^"\']+)["\']', file_path_html or "")

    urls = []
    for href in hrefs:
        if href.startswith("http"):
            href = urlparse(href).path
        href = re.sub(r"/{2,}", "/", href)
        urls.append(href)
    return urls


def is_mao_file(name):
    return name.lower().endswith(".mao")


def is_can_file(name):
    return name.lower().endswith(".xlsx") and "Matrix" not in name


def save_paths_to_text_file(paths, output_path):
    """Writes each given file path on its own line to a plain text file -
    nothing but the paths themselves. Entries that are None (a column whose
    file didn't match its expected pattern) are skipped."""
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    with open(output_path, "w", encoding="utf-8") as f:
        for path in paths:
            if path:
                f.write(path + "\n")


def get_case_input_folder(server_relative_url):
    """Given a server-relative path like '/sites/.../2026年度/08月/
    202608_9999/01_INPUT/03_FSMatrix/file.xlsx' (from commentSystemMatrix),
    returns the path truncated at "01_INPUT" (inclusive) - the shared case
    folder that 03_FSMatrix/02_CAN関連/etc. all live under. None if
    "01_INPUT" isn't one of the path's segments."""
    parts = server_relative_url.split("/")
    if "01_INPUT" not in parts:
        return None
    return "/".join(parts[: parts.index("01_INPUT") + 1])


def upload_file(access_token, folder_server_relative_url, file_name, file_bytes):
    """Uploads file_bytes as file_name into the given SharePoint folder,
    overwriting any existing file of that name."""
    headers = {"Authorization": f"Bearer {access_token}"}
    url = (
        f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{folder_server_relative_url}')"
        f"/Files/add(url='{file_name}',overwrite=true)"
    )
    resp = ntlm_proxy.http_request("POST", url, headers=headers, data=file_bytes)
    resp.raise_for_status()


def download_file(access_token, server_relative_url, local_path):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"{SITE_URL}/_api/web/GetFileByServerRelativeUrl('{server_relative_url}')/$value"
    resp = ntlm_proxy.http_request("GET", url, headers=headers)
    resp.raise_for_status()
    os.makedirs(os.path.dirname(local_path), exist_ok=True)
    with open(local_path, "wb") as f:
        f.write(resp.content)


def get_entity_type_full_name(access_token):
    """SharePoint REST needs this type name to accept a write - one call,
    reused for every item's update."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
    }
    url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')"
    resp = ntlm_proxy.http_request("GET", url, headers=headers, params={"$select": "ListItemEntityTypeFullName"})
    resp.raise_for_status()
    return resp.json()["d"]["ListItemEntityTypeFullName"]


def write_check_result(access_token, entity_type, item_id, result, message, result_log_html=None, etag="*"):
    """
    引数:
        etag: 楽観的並行制御用（2026-09-02新規。①(check_mao)の`sp_client.write_result_columns`と
            同じ設計）。`get_item_result_and_etag`で取得したETagを渡すと、他チェック（①③）が
            その後に列を更新していた場合はSharePointが412 Precondition Failedを返し、本関数は
            `ConcurrentUpdateError`を送出する（呼び出し側は最新値を読み直してリトライする想定）。
            省略時は`"*"`（無条件上書き）で従来どおり動作する。
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "IF-MATCH": etag,
        "X-HTTP-Method": "MERGE",
    }
    url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items({item_id})"
    body = {
        "__metadata": {"type": entity_type},
        RESULT_FIELD_INTERNAL_NAME: result,
        #MESSAGE_FIELD_INTERNAL_NAME: message,
    }
    if result_log_html is not None:
        body[RESULT_LOG_FIELD_INTERNAL_NAME] = result_log_html
    resp = ntlm_proxy.http_request("POST", url, headers=headers, json=body)
    if resp.status_code == 412:
        raise ConcurrentUpdateError(
            f"item {item_id}: ETagが一致しないため書き戻しを中止した（他チェックが先に更新した）"
        )
    resp.raise_for_status()


class ConcurrentUpdateError(Exception):
    """
    列の書き戻し時にETagが一致しなかった（他チェックが割り込んで更新した）ことを示す
    （2026-09-02新規。①(check_mao)の`sp_client.ConcurrentUpdateError`と同じ設計。ロストアップデート
    対策の楽観的並行制御）。呼び出し側は`get_item_result_and_etag`で最新値を読み直してマージ・
    リトライすること（`main()`のリトライループ参照）。
    """


def get_item_result_and_etag(access_token, item_id):
    """
    指定アイテムの現在の`RESULT_FIELD_INTERNAL_NAME`列値とETagを取得する（2026-09-02新規。
    ①(check_mao)の`sp_client.get_item_result_and_etag`と同じ設計）。

    戻り値:
        (現在の列値(str/None), ETag文字列)
    """
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
    }
    url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items({item_id})"
    resp = ntlm_proxy.http_request(
        "GET", url, headers=headers, params={"$select": RESULT_FIELD_INTERNAL_NAME}
    )
    resp.raise_for_status()
    data = resp.json()["d"]
    return data.get(RESULT_FIELD_INTERNAL_NAME), data["__metadata"]["etag"]


def build_check2_summary_line(result):
    """SharePoint列`InputFileCheckResult`へ書き戻す②分の判定サマリ1行を組み立てる
    （2026-09-02新規。`PMOチェック自動化マニュアル.pptx`8ページの文言形式に合わせた）。"""
    return f"{CHECK2_LABEL}：{result}"


def merge_check2_line(existing_value, new_line):
    """
    `InputFileCheckResult`列の既存値(existing_value。①③が書き込んだ行を含みうる。無ければ
    None／空文字）に、②分の新しい行(new_line)をマージする（2026-09-02新規。①(check_mao)の
    `sp_client.merge_check1_line`と同じ設計）。既存値の行のうち`CHECK2_LINE_PREFIX`
    （"チェック②"）で始まる行があれば置き換え、無ければ追加する。①③の行（他のプレフィックスで
    始まる行）はそのまま保持し、結果は`sort_key`で①②③…の番号順に安定ソートする。

    引数:
        existing_value: 列の現在値（複数行、None/空文字の場合あり）
        new_line: ②分の新しい1行

    戻り値:
        マージ後の全体の値（①②③…の番号順にソートされた改行区切りの複数行文字列）
    """
    existing_lines = (existing_value or "").splitlines()
    other_lines = [line for line in existing_lines if not line.startswith(CHECK2_LINE_PREFIX)]
    all_lines = other_lines + [new_line]
    return "\n".join(sorted(all_lines, key=sort_key))


def build_aggregate_file_content(summary_value):
    """
    `InputFileCheckResult`列（①②③統合の判定サマリ。②分マージ後の値）から、SPへアップロードする
    集約結果ファイル（`AGGREGATE_FILE_NAME`）の本文を組み立てる（2026-09-02新規。①(check_mao)の
    `result.build_aggregate_file_content`と同じ形式で①と集約ファイルの内容形式を揃える）。
    """
    return "【チェック結果】\n" + (summary_value or "")


def list_fields(access_token):
    """Diagnostic helper - prints every column's display Title alongside its
    real InternalName, so you can find what RESULT_FIELD_INTERNAL_NAME /
    MESSAGE_FIELD_INTERNAL_NAME (and any other *_FIELD_INTERNAL_NAME above)
    should actually be, instead of guessing from the display name."""
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Accept": "application/json;odata=verbose",
    }
    url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/fields"
    params = {
        "$select": "Title,InternalName,TypeAsString",
        "$filter": "Hidden eq false",
    }
    resp = ntlm_proxy.http_request("GET", url, headers=headers, params=params)
    resp.raise_for_status()

    for f in resp.json()["d"]["results"]:
        print(f"{f['Title']}\t{f['InternalName']}\t{f['TypeAsString']}")


def combine_results(*checks):
    """checks is a sequence of (result, message) tuples. Returns the overall
    (result, message): OK only if every check passed; otherwise NG with the
    first failing check's message."""
    for check_result, check_message in checks:
        if check_result != "OK":
            return "NG", check_message
    return "OK", " / ".join(message for _, message in checks if message)


def new_report_groups():
    """DETAIL_LOCAL_PATHへ書き出す結果ログのグループ構造（[(見出し, 行リスト), ...]）を
    作る。見出しの順序がそのまま出力順になる（2026-09-03改修。従来の`=== OK ===`／
    `=== NG ===`2セクション構成から、チェック対象ファイル単位の3グループへ変更した）。

    戻り値:
        (groups, group_lines) — groupsは`check2_report.build_body`へそのまま渡す構造、
        group_linesは見出し名で行リストを引ける同じリストの辞書ビュー
        （`record_check_result`が追記に使う）
    """
    groups = [
        (check2_report.GROUP_DEFINITION_FILE, []),
        (check2_report.GROUP_FS_MATRIX, []),
        (check2_report.GROUP_GST_COMMAND, []),
    ]
    return groups, dict(groups)


def record_check_result(groups, group_lines, group_name, result, message):
    """Records one check's result as its own bullet under `group_name`
    (e.g. check2_report.GROUP_FS_MATRIX) - so different checks that concern
    different files never get merged into one combined line. Always uses the
    check's own message: it is already written as a full sentence whose
    ending states the verdict (ある／ない、一致している／一致していない), with
    the offending sheet/cell named after it on NG - see
    check_matrix_sheet_parts / check_sid_list_sheet_parts /
    check_definition_file_parts."""
    print(f"[{group_name}]{message}")
    group_lines[group_name].append((result, message))
    write_report(groups)


def write_report(groups):
    """Writes DETAIL_LOCAL_PATH from the current state of `groups`, with a
    timestamp of when the report was (last) written at the top. Since this
    is called after every single result, that timestamp reflects the most
    recent update, not just when the run started, and the report reflects
    everything checked so far even if the run stops partway through."""
    check2_report.write_report(groups, DETAIL_LOCAL_PATH, encoding="utf-8")


def main():
    access_token = get_access_token()
    entity_type = get_entity_type_full_name(access_token)
    items = get_submitted_items(access_token)

    if not items:
        os.makedirs(DOWNLOAD_DIR, exist_ok=True)
        no_items_path = os.path.join(DOWNLOAD_DIR, "no_submitted_items.txt")
        with open(no_items_path, "w", encoding="utf-8") as f:
            f.write("no list with Submitted status is found\n")
        print(f"submitted状態のファイルはありませんでした -> {no_items_path}")
        return

    # The definition-file check (元表 comparison between the two fixed
    # reference files, plus the 〇/× sheet checks) doesn't depend on which
    # item is being processed, so it only needs to run once per script run,
    # not once per item. SHAREPOINT_REFERENCE_FILE_PATH is now a local file -
    # read directly, no download_file() call needed for this one (still used
    # below for the actual SharePoint-submitted files). Each of the three
    # topics (元表, fcl_list_o_FI, fcl_list_x_FI) is recorded into every
    # item's own report below, alongside that item's own checks.
    definition_parts = definition_file_check.check_definition_file_parts(SHAREPOINT_REFERENCE_FILE_PATH)

    # The SharePoint write-back further down needs one combined OK/NG for
    # this whole definition-file check.
    definition_result, definition_message = combine_results(
        *[(part_result, part_message) for _, part_result, part_message in definition_parts]
    )

    for item in items:
        item_id = item.get("Id")

        groups, group_lines = new_report_groups()
        for _label, part_result, part_message in definition_parts:
            record_check_result(groups, group_lines, check2_report.GROUP_DEFINITION_FILE, part_result, part_message)

        file_urls = extract_all_file_urls(item.get(FILE_PATH_FIELD_INTERNAL_NAME))
        if not file_urls:
            result, message = "NG", f"{FILE_PATH_FIELD_INTERNAL_NAME}からリンクを取得できない。"
            write_check_result(access_token, entity_type, item_id, result, message)
            # 提出物そのものが取れないので、FSマトリクスファイルの行としてその理由を残す。
            record_check_result(groups, group_lines, check2_report.GROUP_FS_MATRIX, result, message)
            continue

        names = [os.path.basename(file_url) for file_url in file_urls]

        # All three of these may have more than one file attached, and none
        # of them are an error if nothing matches at all - every matching
        # file gets found/logged, none required.
        hex_fi_paths = [
            url for url in extract_all_file_urls(item.get(HEX_FI_FIELD_INTERNAL_NAME)) if is_mao_file(os.path.basename(url))
        ]
        command_support_urls = [
            url
            for url in extract_all_file_urls(item.get(COMMAND_SUPPORT_FIELD_INTERNAL_NAME))
            if gst_command_check.is_command_support_gst_file(os.path.basename(url))
        ]
        can_paths = [
            url for url in extract_all_file_urls(item.get(CAN_FIELD_INTERNAL_NAME)) if is_can_file(os.path.basename(url))
        ]

        print(f"  commentHexFi (.mao): {hex_fi_paths or '該当ファイルなし'}")
        print(f"  commentCommandSupport (CommandSupportGST*.xlsx): {command_support_urls or '該当ファイルなし'}")
        print(f"  commentCan (.xlsx, Matrixを含まない): {can_paths or '該当ファイルなし'}")

        paths_output_path = os.path.join(DOWNLOAD_DIR, f"file_paths_item_{item_id}.txt")
        save_paths_to_text_file([*file_urls, *hex_fi_paths, *command_support_urls, *can_paths], paths_output_path)

        # "copy and paste temporarily locally": every file attached to
        # commentSystemMatrix is downloaded into the same temp directory,
        # which is automatically deleted once all checks are done.
        with tempfile.TemporaryDirectory() as tmp_dir:
            local_paths = []
            for file_url, name in zip(file_urls, names):
                local_path = os.path.join(tmp_dir, name)
                download_file(access_token, file_url, local_path)
                local_paths.append(local_path)

            matrix_parts = matrix_check.check_matrix_sheet_parts_for_all_files(local_paths)

            # check_gst_command_files_parts() itself returns [] when
            # command_support_urls is empty - no separate skip needed here.
            command_support_local_paths = []
            for url in command_support_urls:
                local_path = os.path.join(tmp_dir, os.path.basename(url))
                download_file(access_token, url, local_path)
                command_support_local_paths.append(local_path)

            gst_parts = gst_command_check.check_gst_command_files_parts(command_support_local_paths)

        # Each check concerns a different file, so each gets its own bullet
        # under its own group heading instead of being merged into one
        # combined line. GST contributes no bullets at all when there was no
        # matching file to check (same "skip silently" rule as before) - its
        # heading is then left out of the report entirely.
        for part_result, part_message in matrix_parts:
            record_check_result(groups, group_lines, check2_report.GROUP_FS_MATRIX, part_result, part_message)
        for part_result, part_message in gst_parts:
            record_check_result(groups, group_lines, check2_report.GROUP_GST_COMMAND, part_result, part_message)

        # ②自身の統合OK/NGは従来どおりcombine_resultsで求めるが、①(check_mao)とのマージに伴い、
        # `InputFileCheckResult`列へ書き込む値は生の"OK"/"NG"ではなく②分の1行
        # （`チェック②（...）：OK/NG`）に変換する（2026-09-02改修）。
        result, message = combine_results(
            *matrix_parts,
            *gst_parts,
            (definition_result, definition_message),
        )
        check2_line = build_check2_summary_line(result)

        # ②専用の詳細（[FS Matrix]／[GSTCommand]／[元表]等の明細、DETAIL_LOCAL_PATH）と、
        # ①②③統合の判定サマリ（AGGREGATE_FILE_NAME）を、この案件のcommentSystemMatrix由来の
        # 01_INPUTフォルダへそれぞれアップロードする（2026-09-02改修。①の2階層構成に合わせた。
        # 詳細ファイルはリンクせず、集約ファイルのみinputFileCheckResultLogからリンクする）。
        input_folder = get_case_input_folder(file_urls[0])
        if input_folder is None:
            print(f"  01_INPUTフォルダが見つからないためレポートをアップロードできない: {file_urls[0]}")
            continue

        with open(DETAIL_LOCAL_PATH, "rb") as f:
            detail_bytes = f.read()
        upload_file(access_token, input_folder, DETAIL_FILE_NAME, detail_bytes)

        # `InputFileCheckResult`列は②専用列ではなく①②③統合の判定サマリ列であり、①③が別
        # プロセスとして独立に同じ列を更新しうる。バッチ開始時に読んだ値のまま書き戻すと
        # ロストアップデートが起きるため、書き戻し直前に最新値・ETagを読み直し、他チェックが
        # 割り込んでいた場合（412 Precondition Failed）は数回まで読み直し・再マージ・再書き込みを
        # リトライする（2026-09-02新規。①(check_mao)の`sp_main._process_one_item`と同じ設計）。
        aggregate_server_relative = f"{input_folder}/{AGGREGATE_FILE_NAME}"
        site_origin = f"{urlparse(SITE_URL).scheme}://{urlparse(SITE_URL).netloc}"
        result_log_html = f'<p><a href="{site_origin}{aggregate_server_relative}">{AGGREGATE_FILE_NAME}</a></p>'

        merged = None
        for _ in range(_MERGE_RETRY_LIMIT):
            existing_value, etag = get_item_result_and_etag(access_token, item_id)
            merged = merge_check2_line(existing_value, check2_line)
            try:
                # SharePoint's "Single line of text" columns cap at 255 characters -
                # writing anything longer raises SPException ("テキストの値が正しく
                # ありません"). Truncate only what goes to SharePoint; the detail
                # report file above already has each check's full, untruncated message.
                write_check_result(
                    access_token, entity_type, item_id, merged, message[:255], result_log_html, etag=etag
                )
                break
            except ConcurrentUpdateError:
                continue
        else:
            raise ConcurrentUpdateError(
                f"item {item_id}: {_MERGE_RETRY_LIMIT}回リトライしても列の書き戻しに成功しなかった"
            )

        # 列に書き戻した最終的な値（＝真実の源）から集約ファイル本文を再生成し、固定名で
        # アップロードする（overwrite=true）。①②③のうち最後に走ったチェックが最も完全な内容を
        # 残す。列側はETagで保護済みのため、この再生成タイミングのわずかな後追いは実害が無い
        # （2026-09-02新規。①(check_mao)の`sp_main._process_one_item`と同じ設計）。
        aggregate_content = build_aggregate_file_content(merged)
        with open(AGGREGATE_LOCAL_PATH, "w", encoding="utf-8-sig") as f:
            f.write(aggregate_content)
            f.write("\n")
        with open(AGGREGATE_LOCAL_PATH, "rb") as f:
            aggregate_bytes = f.read()
        upload_file(access_token, input_folder, AGGREGATE_FILE_NAME, aggregate_bytes)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--list-fields":
        list_fields(get_access_token())
    else:
        main()

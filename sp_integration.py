"""SharePoint REST APIとのやり取り（一覧取得・案件フォルダのダウンロード・結果アップロード・
列の書き戻し）を、①②統合実行用に1本化したモジュール（2026-09-02新規・方式B）。

①(check_mao)の`sp_client.py`をベースに、②(check_fs_matrix)の`input_check_main.py`が個別に
実装していたSP連携関数（`get_case_input_folder`・`upload_file`・`download_file`等）を統合した。
通信は全て`ntlm_proxy.http_request`経由で行う。

①②の判定ロジック本体（`sys.path`経由で既存フォルダから流用するモジュール群）はSharePoint連携に
一切依存しないため、本モジュールが「唯一のSP連携窓口」になる（`unified_main.py`参照）。名前を
`sp_client`から変更したのは、①フォルダを`sys.path`へ追加した際に①の`sp_client.py`と名前が
衝突しないようにするため。

案件リスト`XPX検証案件リスト2`（`sp_auth.LIST_GUID`）から本申請確認済み（`ステータス`列＝
`"3.本申請確認"`、かつ`SubmitStatus`列＝`"Submitted"`のAND条件。①②と同一条件）のアイテムを
取得し、アイテムに添付されたリンク列（`.mao`ファイルへのリンク等）のURLから案件フォルダ
（`01_INPUT`）のサーバー相対パスを逆算する。案件フォルダ配下は`01_HEX関連`・`02_CAN関連`・
`03_FSMatrix`・`04_GSTコマンド装備表`等を含むツリー構造のため、`download_folder_tree`で再帰的に
取得しローカルの一時ディレクトリへ再現する（①の`case_scan.scan_case_folder`が期待するのと
同じフォルダ構造）。②分の対象ファイル（FSマトリクス・GSTコマンド装備表）がこのツリーに含まれない
場合に備え、単体ファイルを直接取得する`download_file`も用意する。

`InputFileCheckResult`列は①②③統合の判定サマリ列であり、本ツールは①②の行を**1回のMERGEで
同時に**書き戻す（`unified_result.merge_check_lines`が担う）。①②を同一プロセスで一括投入するため
①↔②間の412（ETag不一致）は原理的に発生しないが、将来の③やPMOによる手動編集など「本ツール以外の
書き手」に対する保険として、①②が単独稼働していた際と同じETag楽観的並行制御（`get_item_result_and_etag`
／`write_result_columns(..., etag=...)`）は残す。
"""

import os
import re
from urllib.parse import parse_qs, unquote, urlparse

import ntlm_proxy
from sp_auth import LIST_GUID, SITE_URL

# 「ステータス」列（Choice型）。REST `$select`/`$filter`で使う名前はInternalNameではなく
# EntityPropertyNameのため、`OData_`接頭辞を付けている（①`sp_client.py`の実機確認済み値と同一）。
STATUS_FIELD_INTERNAL_NAME = "OData__x30b9__x30c6__x30fc__x30bf__x30"
STATUS_VALUE = "3.本申請確認"

# 「SubmitStatus」列（Text型）。①②と絞り込み条件を統一するため、「ステータス」列に加えて
# この列もAND条件で使う。
SUBMIT_STATUS_FIELD_INTERNAL_NAME = "SubmitStatus"
SUBMIT_STATUS_VALUE = "Submitted"

# 案件フォルダ配下のファイルへのリンク列（.maoファイル添付列）。この列のURLから`01_INPUT`祖先
# パスを逆算する（`derive_input_folder_from_url`）。①`sp_client.FILE_LINK_FIELD_INTERNAL_NAME`
# と同一列。
FILE_LINK_FIELD_INTERNAL_NAME = "commentHexFi"

# ②分のチェック対象ファイルへのリンク列（②`input_check_main.py`と同一列名）。
FS_MATRIX_FIELD_INTERNAL_NAME = "commentSystemMatrix"
COMMAND_SUPPORT_FIELD_INTERNAL_NAME = "commentCommandSupport"

# 案件ID（SharePointリストの`Title`列）
TITLE_FIELD_INTERNAL_NAME = "Title"

# ①②③統合の結果反映列（①`sp_client.py`の実機確認済み値と同一）。
RESULT_FIELD_INTERNAL_NAME = "InputFileCheckResult"
LOG_FIELD_INTERNAL_NAME = "inputFileCheckResultLog"

# 結果ファイルのアップロード先（案件フォルダ直下。実運用前に書込権限をdry-runで確認すること）
RESULT_UPLOAD_SUBFOLDER = None  # Noneなら01_INPUT直下へアップロードする

INPUT_FOLDER_NAME = "01_INPUT"


def _auth_headers(access_token, extra=None):
	headers = {
		"Authorization": f"Bearer {access_token}",
		"Accept": "application/json;odata=verbose",
	}
	if extra:
		headers.update(extra)
	return headers


def get_submitted_items(access_token):
	"""
	本申請確認済み（`ステータス`列＝`"3.本申請確認"`、かつ`SubmitStatus`列＝`"Submitted"`）の
	アイテムを取得する。①②分に必要な列（案件フォルダ特定用のリンク列、②分のFSマトリクス・
	GSTコマンド装備表リンク列）をすべて`$select`する。

	`RESULT_FIELD_INTERNAL_NAME`（判定サマリ列）はここでは取得しない。書き戻し直前に
	`get_item_result_and_etag`で読み直してから`unified_result.merge_check_lines`する
	（ロストアップデート対策。楽観的並行制御。`unified_main.py`参照）。

	戻り値:
		SharePointアイテム(dict)のリスト。
	"""
	headers = _auth_headers(access_token)
	url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items"
	select_fields = ",".join([
		"Id",
		TITLE_FIELD_INTERNAL_NAME,
		FILE_LINK_FIELD_INTERNAL_NAME,
		FS_MATRIX_FIELD_INTERNAL_NAME,
		COMMAND_SUPPORT_FIELD_INTERNAL_NAME,
		STATUS_FIELD_INTERNAL_NAME,
		SUBMIT_STATUS_FIELD_INTERNAL_NAME,
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
	"""
	FilePath系の列がHTMLリンク形式（例:`<p><a href="...">名前</a></p>`）で保持している
	href（1個以上）を全て取り出す。絶対URLはサーバー相対パスへ還元する。
	"""
	hrefs = re.findall(r'href=["\']([^"\']+)["\']', file_path_html or "")

	urls = []
	for href in hrefs:
		if href.startswith("http"):
			href = urlparse(href).path
		href = re.sub(r"/{2,}", "/", href)
		urls.append(href)
	return urls


def derive_input_folder_from_url(server_relative_url):
	"""
	案件フォルダ配下ファイルのサーバー相対URLから、`01_INPUT`フォルダ自身のサーバー相対パスを
	逆算する（①`sp_client.derive_input_folder_from_url`・②`input_check_main.get_case_input_folder`
	と同じロジック。①②で別々に実装されていたものを統合実行用に1つへまとめた）。

	戻り値:
		`01_INPUT`までのサーバー相対パス。`01_INPUT`が見つからない場合はNone。
	"""
	parts = [p for p in server_relative_url.split("/") if p]
	for index, part in enumerate(parts):
		if part == INPUT_FOLDER_NAME:
			return "/" + "/".join(parts[: index + 1])
	return None


def _list_folder_children(access_token, server_relative_folder):
	"""指定フォルダ直下のサブフォルダ一覧・ファイル一覧を取得する。戻り値: (フォルダ名リスト, ファイル情報リスト)。"""
	headers = _auth_headers(access_token)

	folders_url = f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{server_relative_folder}')/Folders"
	folders_resp = ntlm_proxy.http_request(
		"GET", folders_url, headers=headers, params={"$select": "Name,ServerRelativeUrl"}
	)
	folders_resp.raise_for_status()
	folder_names = [
		f["Name"] for f in folders_resp.json()["d"]["results"] if f["Name"] != "Forms"
	]

	files_url = f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{server_relative_folder}')/Files"
	files_resp = ntlm_proxy.http_request(
		"GET", files_url, headers=headers, params={"$select": "Name,ServerRelativeUrl"}
	)
	files_resp.raise_for_status()
	files = files_resp.json()["d"]["results"]

	return folder_names, files


def download_folder_tree(access_token, server_relative_folder, local_dest_dir):
	"""
	`server_relative_folder`配下（サブフォルダを含む）を再帰的に取得し、`local_dest_dir`へ
	同じフォルダ構造で保存する（①の`case_scan.scan_case_folder`がそのまま走査できるように
	するため。②分のFSマトリクス・GSTコマンド装備表ファイルも`03_FSMatrix`等の配下にあれば
	ここで一括取得される）。

	引数:
		access_token: アクセストークン
		server_relative_folder: 取得元フォルダのサーバー相対パス（`01_INPUT`自身を渡す）
		local_dest_dir: 保存先のローカルディレクトリ（`01_INPUT`という名前で作成される）
	"""
	os.makedirs(local_dest_dir, exist_ok=True)
	headers = {"Authorization": f"Bearer {access_token}"}

	folder_names, files = _list_folder_children(access_token, server_relative_folder)

	for file_info in files:
		file_url = f"{SITE_URL}/_api/web/GetFileByServerRelativeUrl('{file_info['ServerRelativeUrl']}')/$value"
		resp = ntlm_proxy.http_request("GET", file_url, headers=headers)
		resp.raise_for_status()
		local_path = os.path.join(local_dest_dir, file_info["Name"])
		with open(local_path, "wb") as f:
			f.write(resp.content)

	for folder_name in folder_names:
		child_server_relative = f"{server_relative_folder}/{folder_name}"
		child_local_dir = os.path.join(local_dest_dir, folder_name)
		download_folder_tree(access_token, child_server_relative, child_local_dir)


def download_file(access_token, server_relative_url, local_path):
	"""
	単一ファイルをサーバー相対URLからローカルパスへ取得する（②`input_check_main.download_file`と
	同じロジック）。`download_folder_tree`で取得済みの`01_INPUT`ツリー内に②分の対象ファイルが
	見つからない場合のフォールバックとして使う（`unified_main.py`参照）。
	"""
	headers = {"Authorization": f"Bearer {access_token}"}
	url = f"{SITE_URL}/_api/web/GetFileByServerRelativeUrl('{server_relative_url}')/$value"
	resp = ntlm_proxy.http_request("GET", url, headers=headers)
	resp.raise_for_status()
	os.makedirs(os.path.dirname(local_path), exist_ok=True)
	with open(local_path, "wb") as f:
		f.write(resp.content)


# ---- ブラウザからコピーしたURLでファイルを取得する（2026-09-02新規） --------------------
# SharePointの共有リンクに付く `/:x:/r/` 形式の接頭辞（`:x:`=Excel・`:w:`=Word等、`/r/`=redirect）。
# サーバー相対パスとしては存在しない飾りなので、パスから取り除く必要がある。
_SHARING_LINK_PREFIX_RE = re.compile(r"^/:[a-zA-Z]:/[a-zA-Z]+/")

# `%u5143`（=元）のようなレガシー（IE系）エスケープ。Doc.aspxリンクの`file=`パラメータで使われる
# ことがあり、`unquote`では戻らないため個別にデコードする。
_LEGACY_PERCENT_U_RE = re.compile(r"%u([0-9a-fA-F]{4})")

_GUID_RE = re.compile(r"^[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$")


def is_sharepoint_file_url(location):
	"""`location`がSharePointのURL（http/https）か。Falseならローカルパスとして扱ってよい。"""
	return isinstance(location, str) and location.strip().lower().startswith(("http://", "https://"))


def _decode_legacy_percent_u(text):
	return _LEGACY_PERCENT_U_RE.sub(lambda m: chr(int(m.group(1), 16)), text)


def parse_file_url(absolute_url):
	"""
	ブラウザのアドレスバー／「リンクのコピー」から得たSharePointのファイルURLを、REST APIで
	取得できる形へ分解する。

	対応する形式:
	  * `.../:x:/r/sites/<site>/_layouts/15/Doc.aspx?sourcedoc=%7BGUID%7D&file=xxx.xlsx&...`
	    （Excel Onlineで開くリンク。パスにファイルの場所が入っていないため、`sourcedoc`の
	    GUID（ファイルのUniqueId）を使って`GetFileById`で取得する）
	  * `.../:x:/r/sites/<site>/<ライブラリ>/.../xxx.xlsx?web=1`（共有リンク）
	  * `.../sites/<site>/<ライブラリ>/.../xxx.xlsx`（通常のURL）

	サイトURLはURL自身から求めるため、`SITE_URL`（案件リストのサイト）とは別サイトのファイルでも
	取得できる。アクセストークンはテナント（ホスト）単位なのでそのまま使える。

	戻り値:
		(サイトURL, ファイル取得用のRESTセレクタ, URLから読み取れたファイル名（無ければ空文字）)
	"""
	parsed = urlparse(absolute_url)
	origin = f"{parsed.scheme}://{parsed.netloc}"

	path = _SHARING_LINK_PREFIX_RE.sub("/", unquote(parsed.path))
	segments = [s for s in path.split("/") if s]

	site_url = origin
	for index, segment in enumerate(segments):
		if segment.lower() in ("sites", "teams") and index + 1 < len(segments):
			site_url = f"{origin}/{segment}/{segments[index + 1]}"
			break

	query = parse_qs(parsed.query)
	source_doc = (query.get("sourcedoc") or [""])[0].strip().strip("{}")
	file_name = os.path.basename(_decode_legacy_percent_u(unquote((query.get("file") or [""])[0])).strip())

	if source_doc:
		if not _GUID_RE.match(source_doc):
			raise ValueError(f"sourcedocをGUIDとして解釈できない: {source_doc!r}")
		return site_url, f"GetFileById(guid'{source_doc}')", file_name

	if any(s.lower() == "_layouts" for s in segments):
		raise ValueError(
			"_layouts配下のリンクだが`sourcedoc`が無いためファイルを特定できない。"
			f"ブラウザで対象ファイルを開き直してURLをコピーし直すこと: {absolute_url}"
		)

	if not segments:
		raise ValueError(f"URLからファイルのパスを読み取れない: {absolute_url}")

	server_relative = "/" + "/".join(segments)
	return site_url, f"GetFileByServerRelativeUrl('{server_relative}')", file_name or segments[-1]


def download_file_from_url(access_token, absolute_url, local_dir, file_name=None):
	"""
	`absolute_url`（ブラウザからコピーしたSharePointのURL）が指すファイルを`local_dir`へ
	ダウンロードし、そのローカルパスを返す。`parse_file_url`が解釈できる全形式に対応する。

	保存名はURLから読み取ったファイル名（`file_name`で明示指定も可）。拡張子が取れない場合は
	`.xlsx`を付ける（拡張子が無いとopenpyxlが「サポート外の形式」で落ちるため）。

	戻り値:
		ダウンロードしたファイルのローカルパス
	"""
	site_url, selector, parsed_name = parse_file_url(absolute_url)

	name = os.path.basename(file_name or parsed_name or "sharepoint_reference")
	if not os.path.splitext(name)[1]:
		name += ".xlsx"

	headers = {"Authorization": f"Bearer {access_token}"}
	url = f"{site_url}/_api/web/{selector}/$value"
	resp = ntlm_proxy.http_request("GET", url, headers=headers)
	resp.raise_for_status()

	os.makedirs(local_dir, exist_ok=True)
	local_path = os.path.join(local_dir, name)
	with open(local_path, "wb") as f:
		f.write(resp.content)
	return local_path


def get_entity_type_full_name(access_token):
	"""SharePoint RESTが書き込みを受け付けるために必要な型名を取得する（1回取得し、全アイテムの更新で再利用する）。"""
	headers = _auth_headers(access_token)
	url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')"
	resp = ntlm_proxy.http_request("GET", url, headers=headers, params={"$select": "ListItemEntityTypeFullName"})
	resp.raise_for_status()
	return resp.json()["d"]["ListItemEntityTypeFullName"]


def _resolve_upload_folder(server_relative_folder):
	"""結果ファイルのアップロード先フォルダのサーバー相対パスを求める（`RESULT_UPLOAD_SUBFOLDER`が
	指定されていればその配下、Noneなら`server_relative_folder`直下）。"""
	if RESULT_UPLOAD_SUBFOLDER is not None:
		return f"{server_relative_folder}/{RESULT_UPLOAD_SUBFOLDER}"
	return server_relative_folder


def result_file_server_relative_url(server_relative_folder, file_name):
	"""
	結果ファイル（固定名・`overwrite=true`でアップロードする前提）のサーバー相対URLを、実際に
	アップロードする前に決定的に求める。集約ファイルのリンクHTMLを、列書き戻し（ETagリトライ
	ループ）より前に組み立てるために使う（`unified_main.py`参照）。
	"""
	return f"{_resolve_upload_folder(server_relative_folder)}/{file_name}"


def upload_result_file(access_token, server_relative_folder, local_file_path):
	"""
	結果ファイル(local_file_path)を`server_relative_folder`（`RESULT_UPLOAD_SUBFOLDER`が指定されて
	いればその配下、Noneなら`server_relative_folder`直下）へアップロードする。

	戻り値:
		アップロードしたファイルのサーバー相対パス
	"""
	upload_folder = _resolve_upload_folder(server_relative_folder)
	file_name = os.path.basename(local_file_path)
	with open(local_file_path, "rb") as f:
		content = f.read()

	# `_auth_headers`経由でAcceptヘッダーを付与する（①`sp_client.upload_result_file`と同じ対策。
	# 独自にヘッダーを組み立ててAcceptを付けないと、SharePointがJSONでなくXML/Atomで応答し
	# resp.json()が失敗する）。
	headers = _auth_headers(access_token, {"Content-Length": str(len(content))})
	url = (
		f"{SITE_URL}/_api/web/GetFolderByServerRelativeUrl('{upload_folder}')"
		f"/Files/add(url='{file_name}',overwrite=true)"
	)
	resp = ntlm_proxy.http_request("POST", url, headers=headers, data=content)
	resp.raise_for_status()
	return resp.json()["d"]["ServerRelativeUrl"]


def build_html_link(server_relative_url, link_text):
	"""結果ファイルのサーバー相対URLから、SharePointのリンク列と同じHTML形式（`<p><a href="...">...</a></p>`）の文字列を組み立てる。"""
	site_host = SITE_URL.split("://", 1)[1].split("/", 1)[0]
	absolute_url = f"https://{site_host}{server_relative_url}"
	return f'<p><a href="{absolute_url}">{link_text}</a></p>'


class ConcurrentUpdateError(Exception):
	"""
	列の書き戻し時にETagが一致しなかった（本ツール以外の書き手が割り込んで更新した）ことを示す
	（①`sp_client.ConcurrentUpdateError`と同じ設計。ロストアップデート対策の楽観的並行制御）。
	呼び出し側は`get_item_result_and_etag`で最新値を読み直してマージ・リトライすること
	（`unified_main.py`参照）。
	"""


def get_item_result_and_etag(access_token, item_id):
	"""
	指定アイテムの現在の`RESULT_FIELD_INTERNAL_NAME`列値とETagを取得する。

	戻り値:
		(現在の列値(str/None), ETag文字列)
	"""
	headers = _auth_headers(access_token)
	url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items({item_id})"
	resp = ntlm_proxy.http_request(
		"GET", url, headers=headers, params={"$select": RESULT_FIELD_INTERNAL_NAME}
	)
	resp.raise_for_status()
	data = resp.json()["d"]
	return data.get(RESULT_FIELD_INTERNAL_NAME), data["__metadata"]["etag"]


def write_result_columns(access_token, entity_type, item_id, summary, log_html, etag="*"):
	"""
	案件アイテム(item_id)へ判定サマリ(`RESULT_FIELD_INTERNAL_NAME`)・結果リンク
	(`LOG_FIELD_INTERNAL_NAME`)を書き戻す（MERGE更新）。

	引数:
		etag: 楽観的並行制御用。`get_item_result_and_etag`で取得したETagを渡すと、本ツール以外の
			書き手がその後に列を更新していた場合はSharePointが412 Precondition Failedを返し、
			本関数は`ConcurrentUpdateError`を送出する（呼び出し側は最新値を読み直してリトライ
			する想定）。省略時は`"*"`（無条件上書き）で動作する。
	"""
	headers = _auth_headers(
		access_token,
		{
			"Content-Type": "application/json;odata=verbose",
			"IF-MATCH": etag,
			"X-HTTP-Method": "MERGE",
		},
	)
	url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/items({item_id})"
	body = {
		"__metadata": {"type": entity_type},
		RESULT_FIELD_INTERNAL_NAME: summary,
		LOG_FIELD_INTERNAL_NAME: log_html,
	}
	resp = ntlm_proxy.http_request("POST", url, headers=headers, json=body)
	if resp.status_code == 412:
		raise ConcurrentUpdateError(
			f"item {item_id}: ETagが一致しないため書き戻しを中止した（他の書き手が先に更新した）"
		)
	resp.raise_for_status()


def list_fields(access_token):
	"""
	診断用: `XPX検証案件リスト2`の全列の表示名(Title)・内部名(InternalName)・
	REST `$select`/`$filter`で実際に使う名前(EntityPropertyName)・型(TypeAsString)を出力する
	（①`sp_client.list_fields`と同じ）。
	"""
	headers = _auth_headers(access_token)
	url = f"{SITE_URL}/_api/web/lists(guid'{LIST_GUID}')/fields"
	params = {
		"$select": "Title,InternalName,EntityPropertyName,TypeAsString",
		"$filter": "Hidden eq false",
	}
	resp = ntlm_proxy.http_request("GET", url, headers=headers, params=params)
	resp.raise_for_status()

	for f in resp.json()["d"]["results"]:
		print(f"{f['Title']}\t{f['InternalName']}\t{f['EntityPropertyName']}\t{f['TypeAsString']}")

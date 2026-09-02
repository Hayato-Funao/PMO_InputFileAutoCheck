"""SharePoint REST APIへアクセスするためのアクセストークン取得モジュール。

①(check_mao)の`sp_auth.py`と同じ実装（`ntlm_proxy.NtlmSession`をMSALの`http_client`に渡し、
トークン取得もNTLM認証プロキシ経由にする。②(check_fs_matrix)の`input_check_main.get_access_token`
も同じ考え方）。

【2026-09-02新規・方式B(統合実行)】①②を1プロセスでまとめて実行する`unified_main.py`用に、
SharePoint連携コードを本フォルダへ集約する一環でこのファイルを持ち込んだ。CLIENT_ID・SITE_URL・
LIST_GUIDは①②と同一の値を暫定採用している。

【大前提】対象SharePoint（`XPX検証案件リスト2`）にアクセスできる全員がユーザーである。本ツールは
提出者ごとの認証を行わず、**運用者ID1つ**で本申請確認済みの全案件を処理する。初回実行時のみ、運用者が
コンソールに表示されるdevice-codeサインインURL・コードでブラウザから対話サインインし、以降は
`CACHE_FILE`（ローカルのトークンキャッシュ）を使って無人で再利用する（`acquire_token_silent`）。

`CACHE_FILE`は①②のトークンキャッシュとは別ファイル（本フォルダ内）であり、初回は本フォルダ単独で
device-codeサインインが必要（①②のキャッシュを流用することはできない）。無人・バッチログオン環境で
対話サインインができない場合は、Windows資格情報マネージャーへ`ntlm_proxy._PROXY_CRED_TARGET`
（`HILS_PROXY_AUTH`）を登録する運用に合わせ、SharePoint側の署名も同様に無人化できる代替手段を
別途検討すること（本モジュールでは未対応）。
"""

import os

import msal

import ntlm_proxy

# TODO: ①②と同一値。実運用前にlist_fields()で疎通確認すること。
CLIENT_ID = "d3590ed6-52b3-4102-aeff-aad2292ab01c"
AUTHORITY = "https://login.microsoftonline.com/organizations"

# TODO: ①②と同一値。実運用前にlist_fields()で疎通確認すること。
SITE_URL = "https://globalhonda.sharepoint.com/sites/jphgt105596"

# 案件リスト`XPX検証案件リスト2`のGUID（①②と同一値。既知の値と一致確認済み）
LIST_GUID = "021F8116-DDD4-4E22-B982-4B738D2BD0CD"

# トークンキャッシュの保存先（本ツールと同じフォルダ内。運用者の初回サインイン後、以降無人で再利用する）
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token_cache.json")


def get_access_token():
	"""
	SharePoint REST APIへアクセスするためのアクセストークンを取得する。

	戻り値:
		アクセストークン文字列

	初回はdevice-codeサインインURL・コードをコンソールへ表示し、対話サインインを求める
	（運用者が1度だけ実施する）。以降は`CACHE_FILE`のキャッシュから無人で取得する。
	"""
	cache = msal.SerializableTokenCache()
	if os.path.exists(CACHE_FILE):
		with open(CACHE_FILE, "r", encoding="utf-8") as f:
			cache.deserialize(f.read())

	# http_clientにNtlmSessionを渡すことで、トークン取得（MSALがAuthorityへ行うHTTP通信）も
	# NTLM認証プロキシ経由にする（①`sp_auth.py`・②`input_check_main.py`と同じ構成）
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
		print(flow["message"])  # サインインURL・コード
		result = app.acquire_token_by_device_flow(flow)

	if cache.has_state_changed:
		with open(CACHE_FILE, "w", encoding="utf-8") as f:
			f.write(cache.serialize())

	if "access_token" not in result:
		raise RuntimeError(f"sign-in failed: {result.get('error_description')}")
	return result["access_token"]

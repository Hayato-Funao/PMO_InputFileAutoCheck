# 出典: ①(check_mao)の`ntlm_proxy.py`（200 Connectedフォールバック済み版）をコピーして流用している。
# 元々は\\RC25020358\c\server\ModelInputCheck\check_fs_matrix\ntlm_proxy.py（②担当作成）が
# 出典であり、①がそれをコピーしてNTLM Type2チャレンジ不在時のフォールバックを追加したもの
# （本番サーバ(RC25020358)実機確認で判明。送信元IPが認証除外設定等の理由でプロキシが認証を
# 要求せず最初のCONNECTで200 Connectedを返す環境に対応）。
#
# 【2026-09-02新規・方式B(統合実行)】①②を1プロセスでまとめて実行する`unified_main.py`用に、
# SharePoint連携コードを本フォルダへ集約する一環でこのファイルを持ち込んだ。①②の既存フォルダは
# 改変しない方針のため、以後のntlm_proxy.py修正は①②とも別々に反映されている可能性があり、
# 本フォルダの版が最新であるとは限らない。将来プロキシ対応の修正が必要になった場合は、
# ①(check_mao)・②(check_fs_matrix)・本フォルダの3箇所を同期すること。
"""
NTLM-authenticated corporate proxy support, adapted from sp_helper.py's
already-proven approach on this network - copied here as this script's own
code rather than importing sp_helper.py as a dependency.

Why this exists: a plain, unauthenticated HTTPS_PROXY setting gets a
"407 Proxy Authentication Required" error on this network - the corporate
proxy requires NTLM authentication for any outbound HTTPS connection. This
module builds an NTLM-authenticated CONNECT tunnel using Windows' own SSPI
(secur32.dll), which authenticates using the current Windows login's
credentials automatically (no password needed), then wraps that tunnel in
TLS and reuses it for one HTTP/HTTPS request.

Provides:
  - http_request(method, url, ...): a requests-like function other scripts
    can use in place of requests.get/requests.post for calls that must go
    through this proxy.
  - NtlmSession: a requests.Session-like object with get()/post() methods,
    usable as msal's `http_client=` parameter so token acquisition also
    goes through this same NTLM tunnel.
"""

import base64
import ctypes
import json as _json
import os
import socket
import ssl
import time
import http.client
from urllib.parse import urlparse, urlencode, quote

PROXY_URL = os.environ.get("HTTPS_PROXY", "http://o365sfprx01.hm.jp.honda.com:8080")

# The proxy on this network has been observed to unreliably drop CONNECT
# tunnels - retry establishing the tunnel a few times before giving up.
_PROXY_CONNECT_RETRY = int(os.environ.get("PROXY_CONNECT_RETRY", "8"))
_PROXY_RETRY_WAIT = float(os.environ.get("PROXY_RETRY_WAIT", "0.4"))

# ── Windows SSPI NTLM proxy authentication ──────────────────────────────

_sec = ctypes.WinDLL("secur32.dll")
_advapi32 = ctypes.WinDLL("advapi32.dll")


class _SecHandle(ctypes.Structure):
    _fields_ = [("dwLower", ctypes.c_uint64), ("dwUpper", ctypes.c_uint64)]


class _TimeStamp(ctypes.Structure):
    _fields_ = [("LowPart", ctypes.c_uint32), ("HighPart", ctypes.c_int32)]


class _SecBuffer(ctypes.Structure):
    _fields_ = [("cbBuffer", ctypes.c_ulong), ("BufferType", ctypes.c_ulong), ("pvBuffer", ctypes.c_void_p)]


class _SecBufferDesc(ctypes.Structure):
    _fields_ = [("ulVersion", ctypes.c_ulong), ("cBuffers", ctypes.c_ulong), ("pBuffers", ctypes.POINTER(_SecBuffer))]


_SECPKG_CRED_OUTBOUND = 2
_ISC_FLAGS = 0x4 | 0x8 | 0x10 | 0x800  # REPLAY_DETECT|SEQUENCE_DETECT|CONFIDENTIALITY|CONNECTION
_SECURITY_NATIVE_DREP = 0x10
_SEC_I_CONTINUE_NEEDED = 0x00090312
_SECBUFFER_TOKEN = 2

# If a credential is explicitly registered in Windows Credential Manager
# under this target name (e.g. via cmdkey, for unattended/batch-logon
# sessions), it's used instead of the current user's implicit credentials.
_PROXY_CRED_TARGET = "HILS_PROXY_AUTH"
_SEC_WINNT_AUTH_IDENTITY_UNICODE = 0x2
_CRED_TYPE_GENERIC = 1


class _CredFiletime(ctypes.Structure):
    _fields_ = [("dwLowDateTime", ctypes.c_uint32), ("dwHighDateTime", ctypes.c_uint32)]


class _CREDENTIAL(ctypes.Structure):
    pass


_CREDENTIAL._fields_ = [
    ("Flags", ctypes.c_uint32),
    ("Type", ctypes.c_uint32),
    ("TargetName", ctypes.c_wchar_p),
    ("Comment", ctypes.c_wchar_p),
    ("LastWritten", _CredFiletime),
    ("CredentialBlobSize", ctypes.c_uint32),
    ("CredentialBlob", ctypes.POINTER(ctypes.c_byte)),
    ("Persist", ctypes.c_uint32),
    ("AttributeCount", ctypes.c_uint32),
    ("Attributes", ctypes.c_void_p),
    ("TargetAlias", ctypes.c_wchar_p),
    ("UserName", ctypes.c_wchar_p),
]


class _SEC_WINNT_AUTH_IDENTITY_W(ctypes.Structure):
    _fields_ = [
        ("User", ctypes.c_wchar_p),
        ("UserLength", ctypes.c_uint32),
        ("Domain", ctypes.c_wchar_p),
        ("DomainLength", ctypes.c_uint32),
        ("Password", ctypes.c_wchar_p),
        ("PasswordLength", ctypes.c_uint32),
        ("Flags", ctypes.c_uint32),
    ]


def _read_stored_credential(target_name):
    """Reads a Generic credential from Windows Credential Manager. Returns
    None if none is registered under target_name."""
    p_cred = ctypes.POINTER(_CREDENTIAL)()
    ok = _advapi32.CredReadW(ctypes.c_wchar_p(target_name), _CRED_TYPE_GENERIC, 0, ctypes.byref(p_cred))
    if not ok:
        return None
    try:
        cred = p_cred.contents
        username = cred.UserName or ""
        blob_size = cred.CredentialBlobSize
        if blob_size and cred.CredentialBlob:
            raw = bytes(bytearray(cred.CredentialBlob[i] & 0xFF for i in range(blob_size)))
            password = raw.decode("utf-16-le", errors="ignore")
        else:
            password = ""
        return username, password
    finally:
        _advapi32.CredFree(p_cred)


def _build_auth_identity():
    """Builds a SEC_WINNT_AUTH_IDENTITY_W from a Credential-Manager-registered
    credential, if one is present under _PROXY_CRED_TARGET; otherwise None,
    which falls back to the current Windows login's implicit credentials."""
    stored = _read_stored_credential(_PROXY_CRED_TARGET)
    if not stored:
        return None
    username, password = stored
    if "\\" in username:
        domain, user = username.split("\\", 1)
    else:
        domain, user = "", username
    identity = _SEC_WINNT_AUTH_IDENTITY_W(
        User=user, UserLength=len(user),
        Domain=domain, DomainLength=len(domain),
        Password=password, PasswordLength=len(password),
        Flags=_SEC_WINNT_AUTH_IDENTITY_UNICODE,
    )
    return identity


def _ntlm_connect(proxy_host, proxy_port, target_host, target_port):
    """Uses Windows SSPI to establish an NTLM-authenticated CONNECT tunnel
    through the proxy, and returns the resulting raw socket."""
    identity = _build_auth_identity()
    p_auth_data = ctypes.byref(identity) if identity is not None else None

    h_cred = _SecHandle()
    ts = _TimeStamp()
    if _sec.AcquireCredentialsHandleW(
            None, "NTLM", _SECPKG_CRED_OUTBOUND, None, p_auth_data, None, None,
            ctypes.byref(h_cred), ctypes.byref(ts)) != 0:
        raise RuntimeError("AcquireCredentialsHandle failed")

    # NTLM Type 1 (Negotiate)
    buf1 = ctypes.create_string_buffer(4096)
    sec1 = _SecBuffer(cbBuffer=4096, BufferType=_SECBUFFER_TOKEN, pvBuffer=ctypes.cast(buf1, ctypes.c_void_p))
    desc1 = _SecBufferDesc(ulVersion=0, cBuffers=1, pBuffers=ctypes.pointer(sec1))
    h_ctxt = _SecHandle()
    attrs = ctypes.c_ulong(0)
    ts2 = _TimeStamp()
    _sec.InitializeSecurityContextW(
        ctypes.byref(h_cred), None, f"HTTP/{target_host}",
        _ISC_FLAGS, 0, _SECURITY_NATIVE_DREP, None, 0,
        ctypes.byref(h_ctxt), ctypes.byref(desc1), ctypes.byref(attrs), ctypes.byref(ts2))
    type1_b64 = base64.b64encode(bytes(buf1[:sec1.cbBuffer])).decode()

    sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
    sock.settimeout(30)
    sock.connect((proxy_host, proxy_port))

    req1 = (f"CONNECT {target_host}:{target_port} HTTP/1.1\r\n"
            f"Host: {target_host}:{target_port}\r\n"
            f"Proxy-Authorization: NTLM {type1_b64}\r\n"
            "Proxy-Connection: Keep-Alive\r\n\r\n")
    sock.sendall(req1.encode())

    # Type 2 (Challenge) response
    raw = b""
    while b"\r\n\r\n" not in raw:
        chunk = sock.recv(4096)
        if not chunk:
            break
        raw += chunk

    hdr_end = raw.index(b"\r\n\r\n")
    hdr_str = raw[:hdr_end].decode("latin-1")
    already_body = raw[hdr_end + 4:]
    content_length = 0
    for line in hdr_str.split("\r\n"):
        if line.lower().startswith("content-length:"):
            content_length = int(line.split(":", 1)[1].strip())
    remaining = max(0, content_length - len(already_body))
    if remaining > 0:
        body_read = b""
        while len(body_read) < remaining:
            chunk = sock.recv(remaining - len(body_read))
            if not chunk:
                break
            body_read += chunk

    type2_b64 = None
    for line in hdr_str.split("\r\n"):
        if line.lower().startswith("proxy-authenticate: ntlm "):
            type2_b64 = line.split(" ", 2)[2].strip()
            break
    if not type2_b64:
        # 【2026-09-01追補4】本番サーバ(RC25020358)で実機確認：送信元IPが認証除外設定等の理由で
        # プロキシがNTLM認証を要求せず、Type1トークン付きCONNECTへ最初から200 Connectedを返す
        # 環境がある。この場合トンネル自体は既に確立済みのため、Type2チャレンジが無いことを
        # エラーにせず、ステータス行が200であればこのソケットをそのまま使う。
        status_line = hdr_str.split("\r\n", 1)[0].strip()
        if status_line.startswith("HTTP/1.1 200") or status_line.startswith("HTTP/1.0 200"):
            return sock
        sock.close()
        raise RuntimeError(f"no NTLM Type 2 challenge - proxy response:\n{hdr_str[:300]}")

    # NTLM Type 3 (Authenticate)
    type2 = base64.b64decode(type2_b64)
    in_mem = ctypes.create_string_buffer(type2)
    in_sec = _SecBuffer(cbBuffer=len(type2), BufferType=_SECBUFFER_TOKEN, pvBuffer=ctypes.cast(in_mem, ctypes.c_void_p))
    in_desc = _SecBufferDesc(ulVersion=0, cBuffers=1, pBuffers=ctypes.pointer(in_sec))
    buf3 = ctypes.create_string_buffer(4096)
    sec3 = _SecBuffer(cbBuffer=4096, BufferType=_SECBUFFER_TOKEN, pvBuffer=ctypes.cast(buf3, ctypes.c_void_p))
    desc3 = _SecBufferDesc(ulVersion=0, cBuffers=1, pBuffers=ctypes.pointer(sec3))
    h_ctxt2 = _SecHandle()
    st3 = _sec.InitializeSecurityContextW(
        ctypes.byref(h_cred), ctypes.byref(h_ctxt), f"HTTP/{target_host}",
        _ISC_FLAGS, 0, _SECURITY_NATIVE_DREP, ctypes.byref(in_desc), 0,
        ctypes.byref(h_ctxt2), ctypes.byref(desc3), ctypes.byref(attrs), ctypes.byref(ts2))
    if st3 not in (0, _SEC_I_CONTINUE_NEEDED):
        sock.close()
        raise RuntimeError(f"InitializeSecurityContext (Type 3) failed: {st3:#010x}")
    type3_b64 = base64.b64encode(bytes(buf3[:sec3.cbBuffer])).decode()

    req2 = (f"CONNECT {target_host}:{target_port} HTTP/1.1\r\n"
            f"Host: {target_host}:{target_port}\r\n"
            f"Proxy-Authorization: NTLM {type3_b64}\r\n"
            "Proxy-Connection: Keep-Alive\r\n\r\n")
    sock.sendall(req2.encode())

    raw2 = b""
    while b"\r\n\r\n" not in raw2:
        chunk = sock.recv(4096)
        if not chunk:
            break
        raw2 += chunk

    if b"200" not in raw2[:30]:
        sock.close()
        raise RuntimeError(f"CONNECT failed: {raw2[:200].decode('latin-1')}")

    return sock


def _parsed_proxy():
    """Returns (host, port) from PROXY_URL, or None if PROXY_URL is empty."""
    if not PROXY_URL:
        return None
    p = urlparse(PROXY_URL)
    return (p.hostname, p.port or 8080)


# ── NTLM-aware HTTP client ────────────────────────────────────────────────

class Response:
    """requests.Response-like minimal implementation - status_code, headers,
    text/json()/content, raise_for_status()."""

    def __init__(self, status_code, headers, body_bytes):
        self.status_code = status_code
        self.headers = {k.lower(): v for k, v in headers}
        self.content = body_bytes

    @property
    def text(self):
        return self.content.decode("utf-8", errors="replace")

    def json(self):
        return _json.loads(self.content)

    def raise_for_status(self):
        if self.status_code >= 400:
            raise Exception(f"HTTP {self.status_code}: {self.text[:200]}")


class _NtlmHTTPSConnection(http.client.HTTPSConnection):
    """HTTPSConnection that reuses an already-established NTLM tunnel socket
    instead of opening a new one."""

    def __init__(self, host, port, established_sock, timeout=60):
        super().__init__(host, port, timeout=timeout)
        self._established_sock = established_sock

    def connect(self):
        self.sock = self._established_sock


def http_request(method, url, headers=None, data=None, json=None, params=None, timeout=60):
    """Performs one HTTP/HTTPS request through the NTLM-authenticated
    corporate proxy (PROXY_URL). Drop-in-ish replacement for
    requests.get/requests.post for calls that must go through that proxy."""
    parsed = urlparse(url)
    host = parsed.hostname
    port = parsed.port or (443 if parsed.scheme == "https" else 80)
    # requests/urllib3 percent-encode the path automatically; http.client does
    # not, and rejects a raw space or non-ASCII character (e.g. Japanese
    # folder names) in the request line with "InvalidURL: ... control
    # characters". The OData literal syntax GetFileByServerRelativeUrl('...')
    # relies on characters like '()$, staying literal, so only encode what's
    # actually unsafe rather than the whole path.
    path = quote(parsed.path or "/", safe="/:()$,'=!*;@&+?#[]%")
    if params:
        path += "?" + urlencode(params)
    elif parsed.query:
        path += "?" + parsed.query

    proxy = _parsed_proxy()
    req_headers = dict(headers or {})

    if json is not None:
        data = _json.dumps(json).encode()
        req_headers.setdefault("Content-Type", "application/json")

    if proxy and parsed.scheme == "https":
        # The proxy has been observed to unreliably drop CONNECT tunnels -
        # retry establishing the tunnel (not the request itself, so no risk
        # of a double-submitted POST) a few times before giving up.
        ssl_sock = None
        last_err = None
        for _attempt in range(_PROXY_CONNECT_RETRY):
            try:
                raw_sock = _ntlm_connect(proxy[0], proxy[1], host, port)
                ctx = ssl.create_default_context()
                ssl_sock = ctx.wrap_socket(raw_sock, server_hostname=host)
                break
            except (ConnectionResetError, ConnectionAbortedError, socket.timeout, ssl.SSLError, OSError) as e:
                last_err = e
                if _attempt < _PROXY_CONNECT_RETRY - 1:
                    time.sleep(_PROXY_RETRY_WAIT)
        if ssl_sock is None:
            raise RuntimeError(f"proxy CONNECT tunnel failed after {_PROXY_CONNECT_RETRY} attempts: {last_err!r}")
        conn = _NtlmHTTPSConnection(host, port, ssl_sock, timeout=timeout)
    elif proxy and parsed.scheme == "http":
        conn = http.client.HTTPConnection(proxy[0], proxy[1], timeout=timeout)
        path = url  # absolute-form request line, required for plain-HTTP proxying
    else:
        if parsed.scheme == "https":
            conn = http.client.HTTPSConnection(host, port, timeout=timeout)
        else:
            conn = http.client.HTTPConnection(host, port, timeout=timeout)

    req_headers.setdefault("Connection", "close")
    req_headers.setdefault("Host", host)

    body = None
    if data is not None:
        if isinstance(data, dict):
            filtered = {k: v for k, v in data.items() if v is not None}
            body = urlencode(filtered).encode()
            req_headers.setdefault("Content-Type", "application/x-www-form-urlencoded")
        elif isinstance(data, str):
            body = data.encode("utf-8")
        else:
            body = data
        req_headers["Content-Length"] = str(len(body))

    conn.request(method, path, body=body, headers=req_headers)
    resp = conn.getresponse()
    body_bytes = resp.read()
    result = Response(resp.status, resp.getheaders(), body_bytes)
    conn.close()
    return result


class NtlmSession:
    """requests.Session-like wrapper (get()/post()), usable as msal's
    `http_client=` parameter so token acquisition also goes through the
    same NTLM-authenticated proxy tunnel as every other request here."""

    proxies = {}
    verify = True
    auth = None

    def get(self, url, **kw):
        return http_request("GET", url, **kw)

    def post(self, url, **kw):
        return http_request("POST", url, **kw)

    def close(self):
        pass

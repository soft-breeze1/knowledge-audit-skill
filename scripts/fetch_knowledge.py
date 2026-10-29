#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_knowledge.py — 从在线知识库抓取已转写好的 Markdown（审查流程的输入缓存）。

输入知识库列表页地址 + 访问令牌，接口契约见下方常量区（``--api-base`` 可整体指定接口前缀）：
    列表  GET <api-base>/knowledge/assets?limit=&offset=&workspaceId=
    版本  GET <api-base>/knowledge/assets/<assetId>/versions
    正文  GET <api-base>/knowledge/assets/<assetId>/versions/<assetVersionId>/content

默认只抓有效文档（assetType=DOCUMENT / lifecycleStatus=ACTIVE / qualityStatus=PASS），且同一文档只抓第一份；
正文按 displayPath 保留目录结构写成 .md。默认并发、断点续跑、失败自动补抓、直连/系统代理自适应、
拒收接口报错体——装好即用，无需改本机设置。本步骤不生成 HTML/XLSX/JSON，结果可直接交给审查脚本。
退出码：0=完成（个别失败也算完成）；1=网络或输入/输出错误；2=参数错误。
"""
from __future__ import annotations

import argparse
import concurrent.futures
import hashlib
import json
import os
import re
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import (  # noqa: E402
    Progress,
    make_timestamped_dir,
    print_error,
    safe_print,
    write_text_file,
)

TOOL_NAME = "fetch_knowledge"

# --------------------------------------------------------------------------- 接口契约
API_PREFIX_DEFAULT = "/api/v1"  # 接口前缀（默认取「站点根 + 本前缀」）
ASSETS_PATH = "/knowledge/assets"  # 资产列表；版本与正文挂在它下面
WORKSPACES_PATH = "/knowledge/workspaces"
TOKEN_HEADER_DEFAULT = "x-access-token"
TOKEN_ENV = "KB_ACCESS_TOKEN"
DEFAULT_LIMIT = 200
PAGE_SIZE = 200  # 单页条数
DEFAULT_MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB，与 audit_markdown.DEFAULT_MAX_FILE_SIZE 对齐
DEFAULT_TIMEOUT = 60
DEFAULT_WORKERS = 8  # 并发线程数（1 = 单线程）
MAX_WORKERS = 32  # 并发线程数上限（避免把知识库打满）
DEFAULT_RETRIES = 3  # 单个请求失败后的重试次数（指数退避：1、2、4 秒）
MAX_REFETCH_ROUNDS = 3  # 抓取结束后对失败文档的补抓轮数上限
RETRY_BACKOFF_SECONDS = 1.0  # 重试退避基数（秒）
CONNECT_PROBE_SECONDS = 8  # 连接方式探测的单次请求超时（秒）
CONTENT_ACCEPT = "application/json, text/markdown, text/plain, */*"
# 接口错误体的可疑顶层键：JSON 对象若只由这些键组成，判定为报错体（正常 Markdown 不会长这样）
_ERROR_BODY_KEYS = frozenset(
    {"success", "error", "errors", "code", "message", "msg", "timestamp", "status", "path", "detail", "trace", "data"}
)
_HTML_ERROR_PREFIXES = ("<!doctype html", "<html", "<head", "<body", "<title")
OUTPUT_SUBDIR = Path("output") / "_kb_source"
# displayPath 常见的原始文档扩展名：抓取结果统一为 .md，避免出现 xxx.xlsx.md 这类双扩展名
_DOC_SOURCE_EXTENSIONS = frozenset(
    {".docx", ".doc", ".xlsx", ".xls", ".pptx", ".ppt", ".pdf", ".txt", ".html", ".htm", ".csv", ".rtf"}
)
MAX_PATH_SEGMENT = 120  # 单层目录/文件名长度上限（Windows 路径长度限制下留出余量）
_ASSET_TYPE_FIELD, _ASSET_TYPE_VALUE = "assetType", "DOCUMENT"
_LIFECYCLE_FIELD, _LIFECYCLE_VALUE = "lifecycleStatus", "ACTIVE"
_QUALITY_FIELD, _QUALITY_VALUE = "qualityStatus", "PASS"
# 响应字段别名：按顺序取第一个非空值，方便对接命名不同的后端
ITEM_KEYS = ("items", "list", "records", "rows")
TOTAL_KEYS = ("totalCount", "total")
ASSET_ID_KEYS = ("assetId", "id")
DISPLAY_PATH_KEYS = ("displayPath", "path", "fileName", "name")
CONTENT_PATH_KEYS = ("filePath", "contentPath")
VERSION_ID_KEYS = ("assetVersionId", "versionId")
WORKSPACE_ID_KEYS = ("workspaceId", "id")
WORKSPACE_NAME_KEYS = ("workspaceName", "name")
_ILLEGAL_PATH_CHARS = re.compile(r'[<>:"|?*\x00-\x1f]')
_RE_PATH_SEPARATORS = re.compile(r"[\\/]+")


def pick(record: Dict[str, Any], keys: Tuple[str, ...], default: Any = None) -> Any:
    """按别名列表取第一个非空字段值（后端命名不同也能对上）。"""
    for key in keys:
        value = record.get(key)
        if value not in (None, ""):
            return value
    return default


class KnowledgeApiError(RuntimeError):
    """知识库接口调用失败（网络异常、HTTP 错误码、响应结构非预期）。

    ``retryable`` 标记该错误是否可能为瞬时问题（超时 / 连接重置 / 429 / 5xx），供请求级重试
    与抓取结束后的补抓轮次判断；令牌无效（401/403）、404 等确定性错误不重试。
    """

    def __init__(self, message: str, retryable: bool = False):
        super().__init__(message)
        self.retryable = retryable


class ContentRejectedError(KnowledgeApiError):
    """正文被判定为接口错误体（JSON / HTML 报错页）：属“漏取风险”，不写盘、也不当作文档。

    单独成类是为了让统计口径分得清“接口报错体被拒收”与“网络/写入失败”，便于事后核对。
    """

    def __init__(self, message: str):
        super().__init__(message, retryable=False)


def resolve_origin(url: str) -> str:
    """从知识库页面地址解析站点根：``https://kb.example.com/knowledge/asset`` → ``https://kb.example.com``。"""
    parts = urllib.parse.urlsplit(url if "://" in url else f"http://{url}")
    if not parts.netloc:
        raise ValueError(f"无法从地址解析出主机：{url}")
    return f"{parts.scheme or 'http'}://{parts.netloc}"


def safe_segment(segment: str) -> str:
    """知识库的目录/文件名段 → 本地可用文件名段（非法字符替换、超长截断并追加哈希避免重名）。"""
    cleaned = _ILLEGAL_PATH_CHARS.sub("_", str(segment)).strip().rstrip(". ")
    if not cleaned:
        cleaned = "_"
    if len(cleaned) > MAX_PATH_SEGMENT:
        digest = hashlib.md5(cleaned.encode("utf-8")).hexdigest()[:8]
        cleaned = f"{cleaned[:MAX_PATH_SEGMENT - 9]}_{digest}"
    return cleaned


def build_relative_path(display_path: str) -> Path:
    """知识库 ``displayPath`` → 本地相对路径（保留目录结构，文件名追加 ``.md``）。"""
    segments = [seg for seg in _RE_PATH_SEPARATORS.split(str(display_path or "")) if seg and seg not in (".", "..")]
    segments = [safe_segment(seg) for seg in (segments or ["未命名文档"])]
    last = segments[-1]
    lower = last.lower()
    if lower.endswith((".md", ".markdown")):
        pass  # 已是 Markdown，保持原样
    elif Path(last).suffix.lower() in _DOC_SOURCE_EXTENSIONS:
        last = last[: -len(Path(last).suffix)] + ".md"  # 源扩展名替换为 .md，避免 xxx.xlsx.md
    else:
        last += ".md"
    segments[-1] = last
    return Path(*segments)


def normalize_doc_key(display_path: str) -> str:
    """同一文档判定键：路径归一（斜杠统一、去扩展名、去空白、小写），用于“只抓第一份”。"""
    text = _RE_PATH_SEPARATORS.sub("/", str(display_path or "")).strip().lower()
    stem, dot, ext = text.rpartition(".")
    if dot and stem and "/" not in ext:
        text = stem
    return text


def normalize_proxy(value: Optional[str]) -> str:
    """``--proxy`` 取值归一：``auto``/空 → 自适应探测；``direct`` → 强制直连；``system`` → 跟随系统代理；
    其它（``host:port`` 或 URL）按显式代理地址处理。"""
    text = str(value or "").strip()
    lowered = text.lower()
    if not text or lowered in ("auto", "detect"):
        return ""
    if lowered in ("direct", "none", "no", "off", "false"):
        return "direct"
    if lowered in ("system", "default", "on", "true"):
        return "system"
    return text


def looks_like_error_body(text: str) -> str:
    """判断响应体是否为接口错误体：是则返回原因（非空），正常 Markdown 返回空串。

    知识库接口出错时可能返回 HTTP 200 + JSON / HTML 报错页；若原样写盘，会把“接口错误”当成正文
    混进审查语料、污染质量报告，因此在写入前做一次廉价校验（``--keep-raw-content`` 可关闭该校验）。
    判定保持保守：只拦截整篇就是报错 JSON（顶层键落在可疑键集合内、且**没有**成功标记）或以
    ``<html`` 等开头的页面；带 ``data`` 载荷或 ``code=0/status=200/success=true`` 的正常响应不会被误判。
    """
    stripped = text.lstrip()
    if not stripped:
        return ""
    lowered = stripped[:120].lower()
    if lowered.startswith("<") and lowered.startswith(_HTML_ERROR_PREFIXES):
        return "HTML 页面（疑似错误页）"
    if stripped[0] in "{[":
        try:
            payload = json.loads(stripped)
        except (ValueError, TypeError):
            return ""
        if isinstance(payload, dict) and payload:
            keys = {str(key).lower() for key in payload}
            failed = payload.get("success") is False or bool(payload.get("error")) or bool(payload.get("errors"))
            # 正常载荷常以 data 承载正文并带 code=0 / status=200 / success=true；这类不应误判为报错体。
            # 兜底判据「顶层键全落在可疑键集合内」只在明确不是成功响应时才成立，并排除 data 键。
            ok_values = (
                payload.get("success") is True
                or payload.get("code") in (0, "0", 200, "200")
                or payload.get("status") in (0, "0", 200, "200", "ok", "success")
            )
            if failed or (not ok_values and keys <= (_ERROR_BODY_KEYS - {"data"})):
                return "JSON 报错体（" + "、".join(sorted(keys)[:4]) + "）"
    return ""


def scan_existing_documents(output_dir: Path) -> Set[str]:
    """扫描目标目录，返回“已存在且非空”的**本地相对路径键**（断点续跑的本地进度记录）。

    键 = 本地相对路径（``/`` 分隔、小写），与 ``build_relative_path(displayPath)`` 的产物逐字对应，
    因此不会出现“带扩展名的 displayPath（如 ``视频.mp4``）与本地文件 ``视频.mp4.md`` 对不上”的错位；
    空文件视为未抓取（上次可能中途中断），会被重新抓取。全程不写任何状态文件。
    """
    keys: Set[str] = set()
    try:
        if not output_dir.is_dir():
            return keys
    except OSError:
        return keys
    for root, _dirnames, filenames in os.walk(output_dir):
        for name in filenames:
            if Path(name).suffix.lower() not in (".md", ".markdown"):
                continue
            path = Path(root) / name
            try:
                if path.stat().st_size <= 0:
                    continue
                rel = path.relative_to(output_dir).as_posix()
            except (OSError, ValueError):
                continue
            keys.add(rel.lower())
    return keys


class KnowledgeClient:
    """知识库接口客户端（仅需访问令牌，标准库实现）。"""

    def __init__(
        self,
        api_base: str,
        token: str,
        timeout: int = DEFAULT_TIMEOUT,
        token_header: str = TOKEN_HEADER_DEFAULT,
        retries: int = DEFAULT_RETRIES,
        proxy: Optional[str] = None,
        validate_content: bool = True,
    ):
        self.api_base = api_base.rstrip("/")
        self.token = token
        self.timeout = timeout
        self.token_header = token_header or TOKEN_HEADER_DEFAULT
        self.retries = max(0, int(retries))
        self.proxy = normalize_proxy(proxy)
        self.validate_content = bool(validate_content)
        self.origin = resolve_origin(self.api_base)
        self._local = threading.local()  # 每线程一份 opener（urllib 的 opener 不宜跨线程共享）
        self._connection_mode = ""
        self._content_branch = ""

    # --- 连接方式（零配置自适应：直连 / 系统代理 / 指定代理）---
    @property
    def connection_mode(self) -> str:
        """当前连接方式：``direct`` / ``system`` / 显式代理地址（首次使用时自动探测并打印一行）。"""
        if not self._connection_mode:
            self.detect_connection()
        return self._connection_mode

    def _opener(self, mode: str):
        """按连接方式构造 opener（每线程缓存一份）：直连用空 ``ProxyHandler`` 绕开系统代理。"""
        cache = getattr(self._local, "openers", None)
        if cache is None:
            cache = {}
            self._local.openers = cache
        if mode not in cache:
            if mode == "direct":
                cache[mode] = urllib.request.build_opener(urllib.request.ProxyHandler({}))
            elif mode == "system":
                cache[mode] = urllib.request.build_opener()
            else:
                proxy_url = mode if "://" in mode else f"http://{mode}"
                handler = urllib.request.ProxyHandler({"http": proxy_url, "https": proxy_url})
                cache[mode] = urllib.request.build_opener(handler)
        return cache[mode]

    def _read(self, url: str, mode: str, timeout: Optional[float] = None) -> bytes:
        """单次请求（不做重试）：把网络与 HTTP 异常统一成带“可否重试”标记的 :class:`KnowledgeApiError`。"""
        request = urllib.request.Request(
            url, headers={self.token_header: self.token, "Accept": CONTENT_ACCEPT}
        )
        try:
            with self._opener(mode).open(request, timeout=timeout or self.timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise KnowledgeApiError(f"令牌无效或无权访问（HTTP {exc.code}）：{url}") from exc
            retryable = exc.code in (408, 425, 429) or exc.code >= 500
            raise KnowledgeApiError(f"HTTP {exc.code}：{url}", retryable=retryable) from exc
        except urllib.error.URLError as exc:
            raise KnowledgeApiError(f"网络不可达（{exc.reason}）：{url}", retryable=True) from exc
        except (TimeoutError, OSError) as exc:
            raise KnowledgeApiError(f"请求失败（{exc}）：{url}", retryable=True) from exc

    def _fetch(self, url: str, mode: str, retries: Optional[int] = None) -> bytes:
        """带退避重试的请求：瞬断（超时 / 连接重置 / 429 / 5xx）自动重试，确定性错误立即失败。"""
        attempts = self.retries if retries is None else max(0, int(retries))
        last: Optional[KnowledgeApiError] = None
        for attempt in range(attempts + 1):
            try:
                return self._read(url, mode)
            except KnowledgeApiError as exc:
                last = exc
                if not exc.retryable or attempt >= attempts:
                    break
                time.sleep(RETRY_BACKOFF_SECONDS * (2 ** attempt))
        raise last if last else KnowledgeApiError(f"请求失败：{url}")

    def detect_connection(self) -> str:
        """连接方式自适应（零配置）：对「直连」与「系统代理」各探测一次并选更快的那个，用户无需改 no_proxy 等设置。"""
        self._connection_mode = ""
        if self.proxy == "direct":
            note = "强制直连（--proxy direct）"
            self._connection_mode = "direct"
        elif self.proxy == "system":
            note = "跟随系统代理（--proxy system）"
            self._connection_mode = "system"
        elif self.proxy:
            note = f"指定代理 {self.proxy}"
            self._connection_mode = self.proxy
        else:
            probe_url = f"{self.api_base}{ASSETS_PATH}?limit=1&offset=0"
            measured: List[Tuple[float, str, str]] = []
            for mode, label in (("direct", "直连"), ("system", "系统代理")):
                started = time.monotonic()
                try:
                    self._read(probe_url, mode, timeout=min(self.timeout, CONNECT_PROBE_SECONDS))
                except KnowledgeApiError as exc:
                    safe_print(f"[{TOOL_NAME}] {label}探测失败：{exc}")
                    continue
                measured.append((time.monotonic() - started, mode, label))
            if not measured:
                raise KnowledgeApiError(
                    "直连与系统代理都无法访问知识库：请检查网络、地址与令牌，或用 --proxy 指定代理。"
                )
            measured.sort()
            elapsed, self._connection_mode, label = measured[0]
            if len(measured) > 1:
                direct = next((item for item in measured if item[1] == "direct"), None)
                best = measured[0]
                # 两种方式都可用且速度接近时优先直连：不要为了几毫秒去绕代理（内网直连更可靠、更稳）
                if direct and best[1] != "direct" and direct[0] <= max(best[0] * 2, best[0] + 0.2):
                    elapsed, self._connection_mode, label = direct
            both_ok = "，两种方式都可用、已取更快的" if len(measured) > 1 else ""
            note = f"{label}（探测 {elapsed:.2f} 秒{both_ok}）"
        safe_print(f"[{TOOL_NAME}] 连接方式：{note}")
        return self._connection_mode

    def _request(self, path: str, params: Optional[Dict[str, Any]] = None) -> bytes:
        """请求接口：相对路径自动补 ``api_base``、绝对 URL 原样使用，并自动退避重试。"""
        url = path if "://" in path else self.api_base + path
        if params:
            query = urllib.parse.urlencode({k: v for k, v in params.items() if v not in (None, "")})
            if query:
                url = f"{url}?{query}"
        return self._fetch(url, self.connection_mode)

    def _json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """取接口 JSON：兼容 ``{"success": true, "result": {...}}`` 与直接返回对象两种封装。"""
        body = self._request(path, params)
        try:
            payload = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            hint = ""
            if body.lstrip()[:1] == b"<":
                hint = "（收到的是 HTML 页面而非接口数据，接口根可能不对——请用 --api-base 指定正确的接口前缀）"
            raise KnowledgeApiError(f"响应不是合法 JSON：{path}{hint}") from exc
        if not isinstance(payload, dict):
            return {}
        if payload.get("success") is False:
            reason = pick(payload, ("message", "msg", "code")) or "未提供原因"
            raise KnowledgeApiError(f"接口返回失败：{path}（{reason}）")
        for key in ("result", "data"):
            inner = payload.get(key)
            if isinstance(inner, dict):
                return inner
        return payload

    @staticmethod
    def _items(payload: Dict[str, Any]) -> List[Dict[str, Any]]:
        """取列表型响应中的条目数组（items / list / records …）。"""
        raw = pick(payload, ITEM_KEYS, default=[]) or []
        return [item for item in raw if isinstance(item, dict)] if isinstance(raw, list) else []

    def list_workspaces(self) -> List[Dict[str, Any]]:
        """工作区清单（``--workspace`` 传名称时用于解析 workspaceId）。"""
        return self._items(self._json(WORKSPACES_PATH))

    def list_assets(self, offset: int, limit: int, workspace_id: Optional[str] = None) -> Dict[str, Any]:
        """文档列表一页（limit/offset 分页），返回 items / hasMore / totalCount。"""
        params: Dict[str, Any] = {"limit": limit, "offset": offset}
        if workspace_id:
            params["workspaceId"] = workspace_id
        return self._json(ASSETS_PATH, params)

    @staticmethod
    def has_more(page: Dict[str, Any], offset: int, item_count: int) -> bool:
        """是否还有下一页：接口给了 hasMore 就用它，否则按 totalCount 推算。"""
        more = page.get("hasMore")
        if isinstance(more, bool):
            return more
        total = 0
        for key in TOTAL_KEYS:
            try:
                total = int(page.get(key) or 0)
            except (TypeError, ValueError):
                total = 0
            if total:
                break
        return bool(total) and offset + item_count < total

    def list_versions(self, asset_id: str) -> List[Dict[str, Any]]:
        """某文档的全部版本；按接口返回顺序取第一个（即“第一份”）。"""
        if not asset_id:
            return []
        return self._items(self._json(f"{ASSETS_PATH}/{asset_id}/versions"))

    # --- 正文地址归一化（filePath 可能是完整 URL / 站点绝对路径 / 含接口前缀 / 相对路径）---
    def _content_url_candidates(self, raw_path: str) -> List[Tuple[str, str]]:
        """``filePath`` → 候选正文 URL 列表（``(写法标识, URL)``，按优先级）；首个成功的写法会被记住复用。"""
        raw = str(raw_path or "").strip()
        if not raw:
            return []
        if raw.startswith(("http://", "https://")):
            return [("absolute", raw)]  # 已是完整地址，原样使用
        origin = self.origin
        prefix = urllib.parse.urlsplit(self.api_base).path.rstrip("/")
        relative = raw.lstrip("/")
        candidates: List[Tuple[str, str]] = []

        def add(branch: str, url: str) -> None:
            if prefix:
                url = url.replace(prefix + prefix, prefix)  # 去掉重复前缀：/api/v1/api/v1/files/x.md
            if url and all(url != existing for _branch, existing in candidates):
                candidates.append((branch, url))

        if raw.startswith("/"):
            if prefix and raw.startswith(prefix + "/"):
                add("origin", origin + raw)  # filePath 已含接口前缀 → 站点根 + 原路径
            else:
                add("origin+prefix", origin + prefix + raw)  # 站点根 + 接口前缀 + filePath
            add("api", self.api_base + raw)  # 历史行为：接口根 + filePath
            add("origin", origin + raw)
        else:
            add("relative", f"{origin}/{relative}")
            add("api", f"{self.api_base}/{relative}")
            add("origin+prefix", f"{origin}{prefix}/{relative}")
        if self._content_branch:
            candidates.sort(key=lambda item: 0 if item[0] == self._content_branch else 1)
        return candidates

    def get_content(self, version: Dict[str, Any]) -> str:
        """取某个版本的正文（已转写好的 Markdown）。

        - ``filePath`` 自动归一化：完整 URL / 站点绝对路径 / 含接口前缀 / 相对路径都能取到，
          首个成功的写法会被记住并优先复用；
        - 响应体若是 JSON / HTML 报错页（HTTP 200 + 错误体），视为取不到正文并换下一个候选，
          避免把“接口错误”原样写盘、污染审查语料；
        - 单个请求瞬断（超时 / 连接重置 / 5xx）会退避重试。
        """
        path = str(pick(version, CONTENT_PATH_KEYS, default="") or "").strip()
        if not path:
            asset_id = pick(version, ASSET_ID_KEYS, default="")
            version_id = pick(version, VERSION_ID_KEYS, default="")
            path = f"{ASSETS_PATH}/{asset_id}/versions/{version_id}/content"
        candidates = self._content_url_candidates(path) or [("api", f"{self.api_base}{path}")]
        errors: List[str] = []
        rejected = False
        for branch, url in candidates:
            try:
                text = self._fetch(url, self.connection_mode).decode("utf-8", errors="replace")
            except KnowledgeApiError as exc:
                errors.append(f"{url} → {exc}")
                continue
            reason = looks_like_error_body(text) if self.validate_content else ""
            if reason:
                rejected = True
                errors.append(f"{url} → 疑似接口错误体（{reason}）")
                continue
            self._content_branch = branch
            return text
        detail = "；".join(errors[:3]) or f"无法取得正文：{path}"
        if rejected:
            raise ContentRejectedError(detail)
        raise KnowledgeApiError(detail)

    def resolve_workspace_id(self, workspace: str) -> str:
        """``--workspace`` 取值：纯数字按 workspaceId 直接使用，否则按工作区名称匹配。"""
        if workspace.isdigit():
            return workspace
        workspaces = self.list_workspaces()
        matches = [w for w in workspaces if str(pick(w, WORKSPACE_NAME_KEYS, default="")) == workspace]
        if not matches:
            names = sorted({str(pick(w, WORKSPACE_NAME_KEYS, default="")) for w in workspaces})
            raise ValueError(f"未找到工作区「{workspace}」；可选：{'、'.join(n for n in names if n)}")
        return str(pick(matches[0], WORKSPACE_ID_KEYS, default=""))


@dataclass
class FetchOutcome:
    """单篇文档的抓取结果（并发线程只产出它；计数统一由主线程按列表顺序完成）。"""

    display_path: str
    target: Path
    status: str = "failed"  # written / empty / oversize / error_body / no_version / api_error / write_error
    reason: str = ""

    @property
    def retryable(self) -> bool:
        """失败原因是否可能为瞬时问题（用于抓取结束后的补抓轮次判定）。

        只有网络/接口类失败才补抓：正文为空、超限、接口报错体这些是确定性结果，重抓也不会变。
        """
        return self.status in ("api_error", "no_version")


def _fetch_one(
    client: KnowledgeClient,
    display_path: str,
    target: Path,
    asset_id: str,
    max_file_size: int,
) -> FetchOutcome:
    """抓取单篇文档（并发执行体：只读网络、只写自己那一个文件，不碰共享可变状态，因此无需加锁）。"""
    try:
        versions = client.list_versions(asset_id)
        if not versions:
            return FetchOutcome(display_path, target, "no_version", "接口未返回版本信息，已跳过。")
        content = client.get_content(versions[0])  # 只取第一个版本
    except ContentRejectedError as exc:
        return FetchOutcome(display_path, target, "error_body", str(exc))
    except KnowledgeApiError as exc:
        return FetchOutcome(display_path, target, "api_error", str(exc))
    if not content.strip():
        return FetchOutcome(display_path, target, "empty", "正文为空，已跳过。")
    size = len(content.encode("utf-8"))
    if max_file_size and size > max_file_size:
        return FetchOutcome(
            display_path,
            target,
            "oversize",
            f"{size} 字节超过单篇上限 {max_file_size} 字节，未写入（可调大 --max-file-size）。",
        )
    try:
        write_text_file(target, content)
    except OSError as exc:
        return FetchOutcome(display_path, target, "write_error", f"写入失败（{exc}）")
    return FetchOutcome(display_path, target, "written")


def _run_fetch_tasks(
    client: KnowledgeClient,
    tasks: List[Tuple[str, Path, str]],
    max_file_size: int,
    workers: int,
    retries: int,
    bar: "Progress",
) -> Tuple[List[FetchOutcome], int]:
    """并发执行抓取任务，返回 ``(与 tasks 同序的结果, 补抓篇次)``。

    workers<=1 时顺序执行；并发时结果按下标回填、主线程按列表顺序统一计数，故不改变抓取结果；
    仍失败的文档补抓若干轮（上限 MAX_REFETCH_ROUNDS），兜住网络瞬断导致的漏取。
    """
    outcomes: List[Optional[FetchOutcome]] = [None] * len(tasks)
    pending = list(range(len(tasks)))
    refetched = 0
    finished = 0
    round_no = 0
    round_cap = max(0, min(int(retries), MAX_REFETCH_ROUNDS))
    while pending:
        round_no += 1
        if round_no > 1:
            time.sleep(RETRY_BACKOFF_SECONDS * (2 ** (round_no - 2)))
        jobs = [(index, tasks[index]) for index in pending]
        if workers > 1 and len(jobs) > 1:
            pool = concurrent.futures.ThreadPoolExecutor(max_workers=workers)
            try:
                futures = {
                    pool.submit(_fetch_one, client, dpath, target, asset_id, max_file_size): index
                    for index, (dpath, target, asset_id) in jobs
                }
                for future in concurrent.futures.as_completed(futures):
                    index = futures[future]
                    try:
                        outcome = future.result()
                    except concurrent.futures.CancelledError:
                        continue  # 用户中断：交由 KeyboardInterrupt 处理（已写入的文件会保留）
                    outcomes[index] = outcome
                    finished += 1
                    bar.update(finished, detail=outcome.display_path)
            finally:
                pool.shutdown(wait=False, cancel_futures=True)
        else:
            for index, (dpath, target, asset_id) in jobs:
                outcomes[index] = _fetch_one(client, dpath, target, asset_id, max_file_size)
                finished += 1
                bar.update(finished, detail=dpath)
        failed = [index for index in pending if outcomes[index] is not None and outcomes[index].retryable]
        if failed and round_no <= round_cap:
            refetched += len(failed)
            bar.update(finished, extra=f"补抓 {len(failed)} 篇（第 {round_no} 轮）")
            pending = failed
            continue
        break
    return [
        outcome or FetchOutcome(tasks[index][0], tasks[index][1], "api_error", "未执行")
        for index, outcome in enumerate(outcomes)
    ], refetched


def fetch_documents(
    api_base: str,
    token: str,
    output_dir: Path,
    limit: int = DEFAULT_LIMIT,
    workspace: Optional[str] = None,
    require_pass: bool = True,
    all_assets: bool = False,
    max_file_size: int = DEFAULT_MAX_FILE_SIZE,
    page_size: int = PAGE_SIZE,
    timeout: int = DEFAULT_TIMEOUT,
    token_header: str = TOKEN_HEADER_DEFAULT,
    workers: int = DEFAULT_WORKERS,
    retries: int = DEFAULT_RETRIES,
    proxy: Optional[str] = None,
    overwrite: bool = False,
    keep_raw_content: bool = False,
    show_progress: bool = True,
) -> Dict[str, Any]:
    """按抓取口径拉取文档并逐篇写入 output_dir，返回统计信息（不生成 JSON 文件）。

    默认：连接方式自适应、单请求退避重试 + 结束后补抓、正文并发（列表分页/过滤/去重仍单线程按序判定，结果与单线程一致）、
    断点续跑（未 overwrite 时跳过本地已存在且非空的 Markdown）。all_assets=True 时不做任何资产过滤。
    """
    client = KnowledgeClient(
        api_base, token, timeout, token_header, retries=retries, proxy=proxy, validate_content=not keep_raw_content
    )
    client.detect_connection()
    worker_count = max(1, min(MAX_WORKERS, int(workers or 1)))
    workspace_id = client.resolve_workspace_id(workspace) if workspace else None
    counters: Dict[str, int] = {
        "listed": 0,
        "written": 0,
        "skipped_duplicate": 0,
        "skipped_type": 0,
        "skipped_lifecycle": 0,
        "skipped_quality": 0,
        "skipped_oversize": 0,
        "skipped_empty": 0,
        "skipped_existing": 0,
        "skipped_error_body": 0,
        "retried": 0,
        "failed": 0,
    }
    warnings: List[str] = []
    written: List[str] = []
    seen_keys: set = set()
    total_count = 0

    existing: Set[str] = set()
    if not overwrite:
        existing = scan_existing_documents(output_dir)
        if existing:
            safe_print(
                f"[{TOOL_NAME}] 断点续跑：{output_dir} 下已有 {len(existing)} 篇 Markdown，已抓到的文档不再重复请求"
                f"（需要强制重抓时加 --overwrite）。"
            )

    # ---- 步骤 1：单线程按列表顺序做过滤 / 去重 / 断点续跑判定，得到本次要抓取的文档清单 ----
    tasks: List[Tuple[str, Path, str]] = []
    reached_limit = False
    offset = 0
    bar = Progress("列表遍历", 0, enabled=None if show_progress else False)
    while True:
        page = client.list_assets(offset, page_size, workspace_id)
        if not total_count:
            try:
                total_count = int(pick(page, TOTAL_KEYS, default=0) or 0)
            except (TypeError, ValueError):
                total_count = 0
            bar.total = max(total_count, 0)  # 首页拿到总数后，进度条改为按百分比 + 预计剩余显示
        items = client._items(page)
        if not items:
            break
        for item in items:
            counters["listed"] += 1
            bar.update(counters["listed"], detail=str(pick(item, DISPLAY_PATH_KEYS, default="") or ""))
            if not all_assets:
                if item.get(_ASSET_TYPE_FIELD) != _ASSET_TYPE_VALUE:
                    counters["skipped_type"] += 1
                    continue
                if item.get(_LIFECYCLE_FIELD) != _LIFECYCLE_VALUE:
                    counters["skipped_lifecycle"] += 1
                    continue
                if require_pass and item.get(_QUALITY_FIELD) != _QUALITY_VALUE:
                    counters["skipped_quality"] += 1
                    continue
            display_path = str(pick(item, DISPLAY_PATH_KEYS, default="") or "")
            key = normalize_doc_key(display_path)
            if key in seen_keys:
                counters["skipped_duplicate"] += 1  # 同一文档只抓第一份
                continue
            seen_keys.add(key)
            local_rel = build_relative_path(display_path)
            target = output_dir / local_rel
            if existing and key and local_rel.as_posix().lower() in existing:
                counters["skipped_existing"] += 1  # 断点续跑：本地已有，不再重复请求
                continue
            if limit and len(tasks) >= limit:
                reached_limit = True  # 只取列表顺序里的前 limit 篇可抓文档
                break
            tasks.append((display_path, target, str(pick(item, ASSET_ID_KEYS, default="") or "")))
        if reached_limit or not client.has_more(page, offset, len(items)):
            break
        offset += page_size
    bar.finish(extra=f"列表遍历 {counters['listed']:,} 条，本次待抓取 {len(tasks):,} 篇")

    # ---- 步骤 2：并发抓取正文，再按列表顺序统一计数（结果确定、可复现）----
    refetched = 0
    if tasks:
        bar2 = Progress("正文抓取", len(tasks), enabled=None if show_progress else False)
        outcomes, refetched = _run_fetch_tasks(
            client, tasks, max_file_size, worker_count, retries, bar2
        )
        for outcome in outcomes:
            if outcome.status == "written":
                counters["written"] += 1
                written.append(outcome.target.as_posix())
            elif outcome.status == "empty":
                counters["skipped_empty"] += 1
                warnings.append(f"{outcome.display_path}：{outcome.reason}")
            elif outcome.status == "oversize":
                counters["skipped_oversize"] += 1
                warnings.append(f"{outcome.display_path}：{outcome.reason}")
            elif outcome.status == "error_body":
                counters["skipped_error_body"] += 1
                warnings.append(f"{outcome.display_path}：{outcome.reason}")
            else:  # no_version / api_error / write_error
                counters["failed"] += 1
                warnings.append(f"{outcome.display_path}：{outcome.reason}")
        counters["retried"] = refetched
        bar2.finish(extra=f"写入 {counters['written']:,} 篇（失败 {counters['failed']:,} 篇）")

    return _result(
        client.api_base,
        output_dir,
        workspace_id,
        written,
        counters,
        warnings,
        total_count,
        connection=client.connection_mode,
        workers=worker_count,
    )


def _result(
    api_base: str,
    output_dir: Path,
    workspace_id: Optional[str],
    written: List[str],
    counters: Dict[str, int],
    warnings: List[str],
    total_count: int = 0,
    connection: str = "",
    workers: int = 0,
) -> Dict[str, Any]:
    """抓取结果（内存对象；本工具不写 JSON / HTML / XLSX）。"""
    return {
        "api_base": api_base,
        "output_dir": str(output_dir),
        "workspace_id": workspace_id,
        "documents": written,
        "counters": counters,
        "warnings": warnings,
        "total_count": int(total_count or 0),
        "connection": connection,
        "workers": workers,
    }


def print_summary(result: Dict[str, Any]) -> None:
    """把抓取统计打印到控制台（含“漏取风险”提示，不写入报告）。"""
    counters = result["counters"]
    total_count = result.get("total_count") or 0
    safe_print(f"[{TOOL_NAME}] 接口地址：{result['api_base']}")
    safe_print(f"[{TOOL_NAME}] 抓取目录：{result['output_dir']}")
    total_note = f"，接口报告总数 {total_count} 条" if total_count else ""
    safe_print(f"[{TOOL_NAME}] 列表遍历 {counters['listed']} 条{total_note}：写入 {counters['written']} 篇，跳过如下 ——")
    safe_print(
        f"[{TOOL_NAME}] 同一文档重复 {counters['skipped_duplicate']}、非文档类型 {counters['skipped_type']}、"
        f"非有效状态 {counters['skipped_lifecycle']}、质量未通过 {counters['skipped_quality']}、"
        f"正文为空 {counters['skipped_empty']}、超过单篇上限 {counters['skipped_oversize']}、"
        f"接口或写入失败 {counters['failed']}。"
    )
    extra_bits: List[str] = []
    if counters.get("skipped_existing"):
        extra_bits.append(f"断点续跑跳过已抓到 {counters['skipped_existing']} 篇")
    if counters.get("skipped_error_body"):
        extra_bits.append(f"接口错误体未写入 {counters['skipped_error_body']} 篇")
    if counters.get("retried"):
        extra_bits.append(f"失败补抓重试 {counters['retried']} 篇次")
    if extra_bits:
        safe_print(f"[{TOOL_NAME}] " + "、".join(extra_bits) + "。")
    if result.get("connection") or result.get("workers"):
        safe_print(f"[{TOOL_NAME}] 本次连接方式：{result.get('connection') or '未知'}；并发线程 {result.get('workers') or 1}。")
    if result["warnings"]:
        safe_print(f"[{TOOL_NAME}] 漏取风险提示（{len(result['warnings'])} 条，前 10 条）：")
        for message in result["warnings"][:10]:
            safe_print(f"[{TOOL_NAME}]   {message}")


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="fetch_knowledge.py",
        description="从在线知识库抓取已转写好的 Markdown 文档（作为质量审查与冲突比对的输入）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            '  python scripts/fetch_knowledge.py "https://kb.example.com/knowledge/asset" --token <token>\n'
            "  python scripts/fetch_knowledge.py <url> --token <token> --limit 0 --all-assets --workers 16\n"
            "默认只抓有效文档、同一文档只抓第一份；默认并发、断点续跑、自动重试与连接方式自适应。\n"
            "退出码：0=完成；1=网络或输入/输出错误；2=参数错误。"
        ),
    )
    parser.add_argument("url", help="知识库文档列表页地址（如 https://kb.example.com/knowledge/asset）")
    parser.add_argument(
        "--token",
        default=None,
        help=f"访问令牌（请求头默认 {TOKEN_HEADER_DEFAULT}）；也可用环境变量 {TOKEN_ENV} 提供",
    )
    parser.add_argument("--token-header", default=TOKEN_HEADER_DEFAULT, help=f"令牌请求头名（默认 {TOKEN_HEADER_DEFAULT}）")
    parser.add_argument(
        "--api-base",
        default=None,
        help=f"接口根地址（默认：站点根 + {API_PREFIX_DEFAULT}）",
    )
    parser.add_argument("--output", default=None, help=f"Markdown 写入目录；默认 {OUTPUT_SUBDIR.as_posix()}/<时间戳>/")
    parser.add_argument(
        "--limit", type=int, default=DEFAULT_LIMIT, help=f"最多抓取篇数（0 表示不限制，默认 {DEFAULT_LIMIT}）"
    )
    parser.add_argument("--workspace", default=None, help="限定工作区（工作区名称或 workspaceId）")
    parser.add_argument(
        "--any-quality", action="store_true", help=f"不检查质量状态（默认只抓 qualityStatus={_QUALITY_VALUE}）"
    )
    parser.add_argument(
        "--all-assets",
        action="store_true",
        help="抓取列表中的全部资产（不做资产类型 / 生命周期 / 质量状态过滤）",
    )
    parser.add_argument(
        "--max-file-size",
        type=int,
        default=DEFAULT_MAX_FILE_SIZE,
        help=f"单篇大小上限（字节，默认 {DEFAULT_MAX_FILE_SIZE}，与质量审查同一口径）",
    )
    parser.add_argument("--timeout", type=int, default=DEFAULT_TIMEOUT, help=f"接口超时秒数（默认 {DEFAULT_TIMEOUT}）")
    parser.add_argument(
        "--workers",
        type=int,
        default=DEFAULT_WORKERS,
        help=f"正文并发线程数（默认 {DEFAULT_WORKERS}，1 表示单线程，上限 {MAX_WORKERS}）",
    )
    parser.add_argument(
        "--retries",
        type=int,
        default=DEFAULT_RETRIES,
        help=f"单个请求失败后的重试次数（指数退避，默认 {DEFAULT_RETRIES}）；抓取结束后还会对失败文档补抓",
    )
    parser.add_argument(
        "--proxy",
        default=None,
        help="代理：auto（默认，自动探测直连与系统代理并选更快的）/ direct（强制直连）/ system（跟随系统代理）/ "
        "<host:port 或 URL>（显式代理）",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="忽略本地已存在的 Markdown，强制重新抓取（默认断点续跑：已抓到的文档不再请求）",
    )
    parser.add_argument(
        "--keep-raw-content",
        action="store_true",
        help="不校验正文（默认拒收 JSON / HTML 接口错误体，避免把接口错误写成文档污染审查语料）",
    )
    parser.add_argument("--no-progress", action="store_true", help="关闭进度条（默认在交互式终端显示）")
    return parser


def run(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    token = (args.token or os.environ.get(TOKEN_ENV) or "").strip()
    if not token:
        safe_print(f"[错误] 缺少访问令牌：请用 --token 传入，或设置环境变量 {TOKEN_ENV}。")
        return 2
    if args.limit < 0:
        safe_print(f"[错误] --limit 不能为负数：{args.limit}")
        return 2
    api_base = (args.api_base or "").strip()
    if not api_base:
        try:
            api_base = f"{resolve_origin(args.url)}{API_PREFIX_DEFAULT}"
        except ValueError as exc:
            safe_print(f"[错误] {exc}")
            return 2
    api_base = api_base.rstrip("/")

    output_dir = Path(args.output).expanduser() if args.output else make_timestamped_dir(OUTPUT_SUBDIR)
    try:
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return print_error(f"无法创建输出目录 {output_dir}：{exc}")

    safe_print(f"[{TOOL_NAME}] 输入地址：{args.url}")
    if args.limit == DEFAULT_LIMIT:
        safe_print(
            f"[{TOOL_NAME}] 提示：默认最多抓取 {DEFAULT_LIMIT} 篇（列表顺序里前 {DEFAULT_LIMIT} 篇可抓文档）；"
            f"抓取整个知识库请显式加 --limit 0。"
        )
    try:
        result = fetch_documents(
            api_base,
            token,
            output_dir,
            limit=args.limit,
            workspace=args.workspace,
            require_pass=not args.any_quality,
            all_assets=args.all_assets,
            max_file_size=args.max_file_size,
            timeout=args.timeout,
            token_header=args.token_header,
            workers=args.workers,
            retries=args.retries,
            proxy=args.proxy,
            overwrite=args.overwrite,
            keep_raw_content=args.keep_raw_content,
            show_progress=not args.no_progress,
        )
    except ValueError as exc:
        safe_print(f"[错误] {exc}")
        return 2
    except KnowledgeApiError as exc:
        return print_error(str(exc))
    print_summary(result)
    if not result["documents"]:
        safe_print(f"[{TOOL_NAME}] 未抓到任何文档，请核对地址、令牌与筛选条件。")
    return 0


def main() -> int:
    try:
        return run()
    except KeyboardInterrupt:
        safe_print(
            f"[{TOOL_NAME}] 已中断；已写入的 Markdown 会保留，重跑同一条命令即从断点继续（默认跳过已抓到的文档）。"
        )
        return 130


if __name__ == "__main__":
    sys.exit(main())
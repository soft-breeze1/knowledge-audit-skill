#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""fetch_knowledge.py — 从在线知识库抓取已转写好的 Markdown 文档（审查流程的输入缓存）。

输入：知识库文档列表页地址（如 ``https://kb.example.com/knowledge/asset``）+ 访问令牌。
默认接口契约（``<api-base>`` 默认为「页面地址的站点根 + ``/api/v1``」，用 ``--api-base`` 可整体改到任意前缀）：

    列表  GET  <api-base>/knowledge/assets?limit=&offset=&workspaceId=
    版本  GET  <api-base>/knowledge/assets/<assetId>/versions
    正文  GET  <api-base>/knowledge/assets/<assetId>/versions/<assetVersionId>/content

路径、筛选字段与响应字段名都集中在下方常量区，对接自己的后端时改常量或传 --api-base 即可。

抓取口径：
- 默认只抓「有效文档」：``assetType=DOCUMENT`` + ``lifecycleStatus=ACTIVE`` + ``qualityStatus=PASS``；
  ``--any-quality`` 放宽质量状态，``--all-assets`` 抓取列表中的全部资产（不做任何过滤）；
- **同一文档只抓第一份**：同一 ``displayPath`` 只保留列表中第一份，每份只在 ``versions`` 里取第一个版本；
- 正文原样写成 ``.md``（按 ``displayPath`` 保留目录结构，UTF-8、LF）；
- 正文为空、超过 ``--max-file-size``（默认 20 MB）或接口报错的资产不写入，属“漏取风险”，只在控制台提示。

用法：
    python scripts/fetch_knowledge.py "https://kb.example.com/knowledge/asset" --token <token>
    python scripts/fetch_knowledge.py <url> --token <token> --api-base "https://kb.example.com/api/v1"
    python scripts/fetch_knowledge.py <url> --token <token> --output input/kb --limit 500
    python scripts/fetch_knowledge.py <url> --token <token> --limit 0 --all-assets   # 全量抓取

``--token`` 也可由环境变量 ``KB_ACCESS_TOKEN`` 提供；``--limit`` 默认 200（``0`` 表示不限制）；
不指定 ``--output`` 时写入 ``output/_kb_source/<时间戳>/``。本步骤不生成 HTML / XLSX / JSON 文件，
抓取结果可直接交给 ``audit_markdown.py`` / ``detect_conflicts.py`` / ``run_all.py``。
退出码：0=完成；1=网络或输入/输出错误；2=参数错误。
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import make_timestamped_dir, print_error, safe_print, write_text_file  # noqa: E402

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
OUTPUT_SUBDIR = Path("output") / "_kb_source"
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
    """知识库接口调用失败（网络异常、HTTP 错误码、响应结构非预期）。"""


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
    segments[-1] = f"{segments[-1]}.md"
    return Path(*segments)


def normalize_doc_key(display_path: str) -> str:
    """同一文档判定键：路径归一（斜杠统一、去扩展名、去空白、小写），用于“只抓第一份”。"""
    text = _RE_PATH_SEPARATORS.sub("/", str(display_path or "")).strip().lower()
    stem, dot, ext = text.rpartition(".")
    if dot and stem and "/" not in ext:
        text = stem
    return text


class KnowledgeClient:
    """知识库接口客户端（仅需访问令牌，标准库实现）。"""

    def __init__(
        self,
        api_base: str,
        token: str,
        timeout: int = DEFAULT_TIMEOUT,
        token_header: str = TOKEN_HEADER_DEFAULT,
    ):
        self.api_base = api_base.rstrip("/")
        self.token = token
        self.timeout = timeout
        self.token_header = token_header or TOKEN_HEADER_DEFAULT

    def _request(self, path: str, params: Optional[Dict[str, Any]] = None) -> bytes:
        url = self.api_base + path
        if params:
            query = urllib.parse.urlencode({k: v for k, v in params.items() if v not in (None, "")})
            if query:
                url = f"{url}?{query}"
        request = urllib.request.Request(
            url, headers={self.token_header: self.token, "Accept": "application/json, text/markdown, text/plain, */*"}
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                return response.read()
        except urllib.error.HTTPError as exc:
            if exc.code in (401, 403):
                raise KnowledgeApiError(f"令牌无效或无权访问（HTTP {exc.code}）：{url}") from exc
            raise KnowledgeApiError(f"HTTP {exc.code}：{url}") from exc
        except urllib.error.URLError as exc:
            raise KnowledgeApiError(f"网络不可达（{exc.reason}）：{url}") from exc
        except OSError as exc:
            raise KnowledgeApiError(f"请求失败（{exc}）：{url}") from exc

    def _json(self, path: str, params: Optional[Dict[str, Any]] = None) -> Dict[str, Any]:
        """取接口 JSON：兼容 ``{"success": true, "result": {...}}`` 与直接返回对象两种封装。"""
        body = self._request(path, params)
        try:
            payload = json.loads(body.decode("utf-8"))
        except (UnicodeDecodeError, json.JSONDecodeError) as exc:
            raise KnowledgeApiError(f"响应不是合法 JSON：{path}") from exc
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

    def get_content(self, version: Dict[str, Any]) -> str:
        """取某个版本的正文（已转写好的 Markdown）；版本自带 filePath 时直接用它。"""
        path = str(pick(version, CONTENT_PATH_KEYS, default="") or "").strip()
        if not path:
            asset_id = pick(version, ASSET_ID_KEYS, default="")
            version_id = pick(version, VERSION_ID_KEYS, default="")
            path = f"{ASSETS_PATH}/{asset_id}/versions/{version_id}/content"
        return self._request(path).decode("utf-8", errors="replace")

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
) -> Dict[str, Any]:
    """按抓取口径拉取文档并逐篇写入 output_dir，返回统计信息（不生成 JSON 文件）。

    all_assets=True 时不做资产类型 / 生命周期 / 质量状态过滤（抓取列表中的全部资产）。
    """
    client = KnowledgeClient(api_base, token, timeout, token_header)
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
        "failed": 0,
    }
    warnings: List[str] = []
    written: List[str] = []
    seen_keys: set = set()
    total_count = 0

    offset = 0
    while True:
        page = client.list_assets(offset, page_size, workspace_id)
        if not total_count:
            try:
                total_count = int(pick(page, TOTAL_KEYS, default=0) or 0)
            except (TypeError, ValueError):
                total_count = 0
        items = client._items(page)
        if not items:
            break
        for item in items:
            counters["listed"] += 1
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
            try:
                versions = client.list_versions(str(pick(item, ASSET_ID_KEYS, default="") or ""))
                if not versions:
                    counters["failed"] += 1
                    warnings.append(f"{display_path}：接口未返回版本信息，已跳过。")
                    continue
                content = client.get_content(versions[0])  # 只取第一个版本
            except KnowledgeApiError as exc:
                counters["failed"] += 1
                warnings.append(f"{display_path}：{exc}")
                continue
            if not content.strip():
                counters["skipped_empty"] += 1
                warnings.append(f"{display_path}：正文为空，已跳过。")
                continue
            size = len(content.encode("utf-8"))
            if max_file_size and size > max_file_size:
                counters["skipped_oversize"] += 1
                warnings.append(
                    f"{display_path}：{size} 字节超过单篇上限 {max_file_size} 字节，未写入（可调大 --max-file-size）。"
                )
                continue
            target = output_dir / build_relative_path(display_path)
            try:
                write_text_file(target, content)
            except OSError as exc:
                counters["failed"] += 1
                warnings.append(f"{display_path}：写入失败（{exc}）")
                continue
            counters["written"] += 1
            written.append(target.as_posix())
            if limit and counters["written"] >= limit:
                return _result(client.api_base, output_dir, workspace_id, written, counters, warnings, total_count)
        if not client.has_more(page, offset, len(items)):
            break
        offset += page_size

    return _result(client.api_base, output_dir, workspace_id, written, counters, warnings, total_count)


def _result(
    api_base: str,
    output_dir: Path,
    workspace_id: Optional[str],
    written: List[str],
    counters: Dict[str, int],
    warnings: List[str],
    total_count: int = 0,
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
            '  python scripts/fetch_knowledge.py <url> --token <token> --api-base "https://kb.example.com/api/v1"\n'
            "  python scripts/fetch_knowledge.py <url> --token <token> --output input/kb --limit 500\n"
            "  python scripts/fetch_knowledge.py <url> --token <token> --workspace 示例工作区\n"
            "  python scripts/fetch_knowledge.py <url> --token <token> --limit 0 --all-assets\n"
            "默认只抓有效文档（文档类型 / 有效状态 / 质量通过），**同一文档只抓第一份**；\n"
            "--all-assets 抓取列表中的全部资产（不做过滤），配合 --limit 0 即全量抓取；\n"
            "抓取结果逐篇写成 .md（按 displayPath 保留目录结构），可直接交给 run_all.py 审查。\n"
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
        safe_print("[错误] 用户中断操作。")
        return 130


if __name__ == "__main__":
    sys.exit(main())
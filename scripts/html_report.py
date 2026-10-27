#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""html_report.py — 自包含 HTML 报告渲染（仅标准库；样式内联，无外部资源、无网络依赖）。

产出为单文件 HTML：双击即可打开，或直接作为附件发送；不引用任何 CDN/图片/脚本。
本模块只提供渲染原语，具体内容由调用方（audit_markdown.py / detect_conflicts.py）组织：

    kpi_cards(items)                     顶部指标卡片
    table(headers, rows, align=None)     通用表格
    details(summary, body_html, open_)   可折叠明细（默认折叠，保持页面简洁）
    section(title, body_html, note)      分节卡片
    build_html(...)                      统一外壳（页头/元信息/分节/页脚）
"""
from __future__ import annotations

import html
from typing import Any, Iterable, List, Optional, Sequence, Tuple

__all__ = ["esc", "kpi_cards", "table", "details", "section", "build_html"]

# 报告样式（内联在 HTML 中，不依赖任何外部资源）
_CSS = """
:root{
  --bg:#f5f6f8; --card:#ffffff; --line:#e3e6ea; --text:#1f2937; --muted:#6b7280;
  --accent:#2563eb; --accent-soft:#eff4ff; --warn:#b45309; --warn-soft:#fff7ed;
  --ok:#047857; --ok-soft:#ecfdf5;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--text);
  font:15px/1.65 -apple-system,"Segoe UI","Microsoft YaHei",Roboto,"Helvetica Neue",Arial,sans-serif;}
.wrap{max-width:1080px;margin:0 auto;padding:28px 18px 56px}
header.page{background:linear-gradient(135deg,#1e3a8a,#2563eb);color:#fff;border-radius:14px;
  padding:24px 26px;box-shadow:0 6px 18px rgba(30,58,138,.18)}
header.page h1{margin:0 0 6px;font-size:23px;letter-spacing:.3px}
header.page p{margin:0;opacity:.92;font-size:13.5px}
.meta{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:8px 20px;
  margin:18px 0 6px;font-size:13px;color:var(--muted)}
.meta b{color:var(--text);font-weight:600}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(158px,1fr));gap:12px;margin:18px 0}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .k{font-size:12.5px;color:var(--muted);margin-bottom:4px}
.kpi .v{font-size:22px;font-weight:600;letter-spacing:.4px}
.kpi.warn .v{color:var(--warn)} .kpi.ok .v{color:var(--ok)} .kpi.hot .v{color:var(--accent)}
section.card{background:var(--card);border:1px solid var(--line);border-radius:12px;
  padding:18px 20px;margin:16px 0;box-shadow:0 1px 2px rgba(16,24,40,.04)}
section.card>h2{margin:0 0 4px;font-size:17px}
section.card .note{color:var(--muted);font-size:12.5px;margin:2px 0 12px}
table{border-collapse:collapse;width:100%;font-size:13.5px;margin:6px 0 2px}
th,td{border:1px solid var(--line);padding:7px 9px;text-align:left;vertical-align:top;word-break:break-word}
th{background:#f8fafc;font-weight:600;white-space:nowrap}
tr:nth-child(even) td{background:#fcfdfe}
table td.num,table th.num{text-align:right;white-space:nowrap}
details{border:1px solid var(--line);border-radius:10px;margin:8px 0;background:#fff}
details>summary{cursor:pointer;padding:9px 12px;font-weight:600;font-size:13.5px;
  list-style:none;display:flex;gap:8px;align-items:center;flex-wrap:wrap}
details>summary::-webkit-details-marker{display:none}
details>summary::before{content:"▸";color:var(--muted);font-size:12px}
details[open]>summary::before{content:"▾"}
details>.body{padding:2px 12px 12px;border-top:1px solid var(--line)}
.badge{display:inline-block;padding:1px 8px;border-radius:999px;font-size:12px;
  background:var(--accent-soft);color:var(--accent);border:1px solid #dbe6ff}
.badge.warn{background:var(--warn-soft);color:var(--warn);border-color:#fde3c8}
.badge.ok{background:var(--ok-soft);color:var(--ok);border-color:#c9f0dd}
.badge.mute{background:#f3f4f6;color:var(--muted);border-color:var(--line)}
code{background:#f3f4f6;border-radius:4px;padding:1px 5px;font-size:12.5px;
  font-family:ui-monospace,Consolas,"Courier New",monospace}
footer.page{margin-top:22px;color:var(--muted);font-size:12px;text-align:center}
ul.plain{margin:6px 0;padding-left:20px}
"""


def esc(value: Any) -> str:
    """HTML 转义（None 视为空串）。"""
    return html.escape("" if value is None else str(value), quote=True)


def kpi_cards(items: Sequence[Tuple[str, Any, str]]) -> str:
    """渲染顶部指标卡片；items 为 (标题, 数值, 语气) 序列，语气 ∈ {"", hot, warn, ok}。"""
    parts: List[str] = ['<div class="kpis">']
    for label, value, tone in items:
        cls = f"kpi {tone}".strip()
        parts.append(
            f'<div class="{esc(cls)}"><div class="k">{esc(label)}</div>'
            f'<div class="v">{esc(value)}</div></div>'
        )
    parts.append("</div>")
    return "".join(parts)


def table(
    headers: Sequence[str],
    rows: Iterable[Sequence[Any]],
    align: Optional[Sequence[str]] = None,
) -> str:
    """渲染表格；align 为每列对齐方式（"num" 右对齐，其余左对齐）。"""
    aligns = list(align or [])
    parts: List[str] = ["<table><thead><tr>"]
    for idx, head in enumerate(headers):
        cls = ' class="num"' if idx < len(aligns) and aligns[idx] == "num" else ""
        parts.append(f"<th{cls}>{esc(head)}</th>")
    parts.append("</tr></thead><tbody>")
    for row in rows:
        parts.append("<tr>")
        for idx, cell in enumerate(row):
            cls = ' class="num"' if idx < len(aligns) and aligns[idx] == "num" else ""
            parts.append(f"<td{cls}>{esc(cell)}</td>")
        parts.append("</tr>")
    parts.append("</tbody></table>")
    return "".join(parts)


def details(summary: str, body_html: str, open_: bool = False) -> str:
    """渲染可折叠明细块（默认折叠，避免长报告一屏铺满）。"""
    open_attr = " open" if open_ else ""
    return f'<details{open_attr}><summary>{summary}</summary><div class="body">{body_html}</div></details>'


def section(title: str, body_html: str, note: Optional[str] = None) -> str:
    """渲染分节卡片。"""
    head = f"<h2>{esc(title)}</h2>"
    note_html = f'<div class="note">{esc(note)}</div>' if note else ""
    return f'<section class="card">{head}{note_html}{body_html}</section>'


def build_html(
    title: str,
    subtitle: str,
    meta_pairs: Sequence[Tuple[str, str]],
    sections: Sequence[str],
    footer_note: str = "",
) -> str:
    """拼装完整 HTML 文档（内联样式、自包含、无外部依赖）。"""
    meta_html = "".join(f"<div><b>{esc(k)}</b>：{esc(v)}</div>" for k, v in meta_pairs)
    footer_html = f'<footer class="page">{esc(footer_note)}</footer>' if footer_note else ""
    return (
        '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>\n<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
        f"<title>{esc(title)}</title>\n<style>{_CSS}</style>\n</head>\n<body>\n"
        '<div class="wrap">\n'
        f'<header class="page"><h1>{esc(title)}</h1><p>{esc(subtitle)}</p></header>\n'
        f'<div class="meta">{meta_html}</div>\n'
        + "\n".join(sections)
        + "\n" + footer_html + "\n</div>\n</body>\n</html>\n"
    )

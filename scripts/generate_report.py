#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""generate_report.py — knowledge-audit-skill 共享报告工具库。

为 audit_markdown.py / detect_conflicts.py 提供：
- 交付文件写出：HTML 报告（UTF-8、无 BOM、LF 换行）与 XLSX 清单（表头配色、列宽、
  单元格边框、冻结首行、自动筛选，Excel / WPS 双击即可打开）；不生成 Markdown / JSON 文件
- 时间戳与运行目录生成（每次任务一个带完整时间戳的子目录，互不覆盖）
- 文件展示路径、文本截断/计数、控制台安全输出等小工具

本模块不包含任何检测逻辑，仅供 scripts/ 下的脚本复用。
"""
from __future__ import annotations

import re
import shutil
import sys
import time
import zipfile
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

TOOL_VERSION = "1.0.5"


def now_utc_iso() -> str:
    """返回当前 UTC 时间的 ISO 8601 字符串（如 2026-10-22T09:30:00Z）。"""
    return datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")


def now_local_str() -> str:
    """返回本地时间的可读字符串（含 UTC 偏移）。"""
    now = datetime.now().astimezone()
    return now.strftime("%Y-%m-%d %H:%M:%S UTC%z")


def make_timestamped_dir(base_dir: Any, stamp: Any = None) -> Path:
    """在输出根目录下创建带完整时间戳的运行目录，避免多次运行的交付文件互相覆盖。

    - 目录名格式：``YYYYMMDD-HHMMSS``（本地时间，如 ``20261023-151204``）；
    - 同一秒内重复创建时追加 ``_2``、``_3``…，保证目录唯一；
    - base_dir 不存在时按需创建。
    """
    stamp = stamp or datetime.now().strftime("%Y%m%d-%H%M%S")
    base = Path(base_dir)
    base.mkdir(parents=True, exist_ok=True)
    run_dir = base / stamp
    suffix = 2
    while run_dir.exists():
        run_dir = base / f"{stamp}_{suffix}"
        suffix += 1
    run_dir.mkdir(parents=True, exist_ok=True)
    return run_dir


def write_text_file(path: Any, content: str):
    """以 UTF-8（无 BOM、LF 换行）写出文本文件，返回 Path。"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)
    return path


def relative_path(path: Any, input_path: Any) -> str:
    """文件在报告中的展示路径：输入为单个文件时用文件名，否则用相对输入目录的路径（/ 分隔）。"""
    path, input_path = Path(path), Path(input_path)
    if input_path.is_file():
        return path.name
    try:
        return path.relative_to(input_path).as_posix()
    except ValueError:
        return path.as_posix()


# ---------------------------------------------------------------------------
# XLSX 清单写出（Open XML = ZIP + XML，标准库手写，不需要 openpyxl 等第三方库）。
# 用 XLSX 而不是 CSV：CSV 是纯文本，带不了表头配色、列宽、边框与冻结行，打开时容易挤成一团。
# ---------------------------------------------------------------------------
_XLSX_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_XLSX_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'

# 视觉参数：表头深蓝底 + 白字加粗，隔行极浅蓝斑马纹，细灰边框
_HEADER_FILL = "FF2E75B6"
_BAND_FILL = "FFF1F6FC"
_LINE_COLOR = "FFB7C4D4"
_FONT_NAME = "Microsoft YaHei"  # 中英文混排时显示为“微软雅黑”

# cellXfs 索引：0=默认，其余按用途命名（详见 _styles_xml()）
_ST_HEADER = 1
_ST_BODY = 2
_ST_BODY_BAND = 3
_ST_NUM = 4
_ST_NUM_BAND = 5

_INT_OR_FLOAT_RE = re.compile(r"-?\d+(?:\.\d+)?")
# XML 1.0 不允许的字符：控制字符（制表 0x09 / 换行 0x0A / 回车 0x0D 除外）与 U+FFFE、U+FFFF
_ILLEGAL_XML_CHARS_RE = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ufffe\uffff]")


def _column_letters(index: int) -> str:
    """0 基列号 → Excel 列名（0→A、25→Z、26→AA）。"""
    letters = ""
    index += 1
    while index:
        index, remainder = divmod(index - 1, 26)
        letters = chr(ord("A") + remainder) + letters
    return letters


def _xml_text(value: Any) -> str:
    """转义为 XML 文本：去掉 XML 1.0 不允许的字符，再转义 & < > 与引号。"""
    text = "" if value is None else str(value)
    text = text.encode("utf-8", errors="replace").decode("utf-8")
    text = _ILLEGAL_XML_CHARS_RE.sub("", text)
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;")


def safe_sheet_name(name: str, fallback: str = "清单") -> str:
    """工作表名：去掉 Excel 不允许的字符与首尾单引号，长度截到 31 字符以内。"""
    cleaned = re.sub(r"[\\/*?\[\]:]", " ", str(name or "")).strip().strip("'").strip()
    return cleaned[:31] or fallback


def _xf(font_id: int, fill_id: int, border_id: int, alignment: str) -> str:
    return (
        f'<xf numFmtId="0" fontId="{font_id}" fillId="{fill_id}" borderId="{border_id}" xfId="0"'
        f' applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
        f'<alignment {alignment}/></xf>'
    )


def _styles_xml() -> str:
    """单元格样式表：表头（深蓝底白字居中）、正文、正文斑马纹、数字、数字斑马纹。"""
    return "".join(
        [
            _XML_DECL,
            f'<styleSheet xmlns="{_XLSX_NS}">',
            '<fonts count="2">',
            f'<font><sz val="10.5"/><color rgb="FF1F2937"/><name val="{_FONT_NAME}"/>'
            '<charset val="134"/></font>',
            f'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="{_FONT_NAME}"/>'
            '<charset val="134"/></font>',
            "</fonts>",
            '<fills count="4">',
            '<fill><patternFill patternType="none"/></fill>',
            '<fill><patternFill patternType="gray125"/></fill>',
            '<fill><patternFill patternType="solid">'
            f'<fgColor rgb="{_HEADER_FILL}"/><bgColor indexed="64"/></patternFill></fill>',
            '<fill><patternFill patternType="solid">'
            f'<fgColor rgb="{_BAND_FILL}"/><bgColor indexed="64"/></patternFill></fill>',
            "</fills>",
            '<borders count="2">',
            "<border><left/><right/><top/><bottom/><diagonal/></border>",
            "<border>"
            f'<left style="thin"><color rgb="{_LINE_COLOR}"/></left>'
            f'<right style="thin"><color rgb="{_LINE_COLOR}"/></right>'
            f'<top style="thin"><color rgb="{_LINE_COLOR}"/></top>'
            f'<bottom style="thin"><color rgb="{_LINE_COLOR}"/></bottom>'
            "<diagonal/></border>",
            "</borders>",
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
            '<cellXfs count="6">',
            '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>',
            _xf(1, 2, 1, 'horizontal="center" vertical="center" wrapText="1"'),
            _xf(0, 0, 1, 'vertical="top" wrapText="1"'),
            _xf(0, 3, 1, 'vertical="top" wrapText="1"'),
            _xf(0, 0, 1, 'horizontal="right" vertical="top"'),
            _xf(0, 3, 1, 'horizontal="right" vertical="top"'),
            "</cellXfs>",
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
            '<dxfs count="0"/>',
            "</styleSheet>",
        ]
    )


def _cell_xml(reference: str, style: int, value: Any, numeric: bool) -> str:
    """单个单元格：空值不写（保持表格开阔）；数字列写数值型，其余写内联字符串。"""
    if value is None or value == "":
        return ""
    text = str(value)
    if numeric and _INT_OR_FLOAT_RE.fullmatch(text):
        return f'<c r="{reference}" s="{style}"><v>{text}</v></c>'
    return f'<c r="{reference}" s="{style}" t="inlineStr"><is><t xml:space="preserve">{_xml_text(text)}</t></is></c>'


def _sheet_xml(
    header: Sequence[str],
    rows: Iterable[Sequence[Any]],
    widths: Optional[Sequence[int]],
    numeric_cols: Sequence[int],
    banded: bool,
) -> str:
    """工作表内容：首行冻结、表头配色居中、列宽、斑马纹与自动筛选。"""
    data_rows = [list(row) for row in rows]
    last_col = _column_letters(len(header) - 1)
    numeric = set(numeric_cols)
    parts = [
        _XML_DECL,
        f'<worksheet xmlns="{_XLSX_NS}">',
        f'<dimension ref="A1:{last_col}{len(data_rows) + 1}"/>',
        '<sheetViews><sheetView tabSelected="1" workbookViewId="0" showGridLines="0">'
        '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
        '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
        "</sheetView></sheetViews>",
        '<sheetFormatPr defaultRowHeight="18"/>',
    ]
    if widths:
        parts.append("<cols>")
        for idx, width in enumerate(widths):
            parts.append(f'<col min="{idx + 1}" max="{idx + 1}" width="{width}" customWidth="1"/>')
        parts.append("</cols>")
    parts.append('<sheetData><row r="1" ht="26" customHeight="1">')
    for idx, title in enumerate(header):
        parts.append(_cell_xml(f"{_column_letters(idx)}1", _ST_HEADER, title, numeric=False))
    parts.append("</row>")
    for row_idx, row in enumerate(data_rows, start=2):
        band = banded and row_idx % 2 == 0
        body_style = _ST_BODY_BAND if band else _ST_BODY
        num_style = _ST_NUM_BAND if band else _ST_NUM
        parts.append(f'<row r="{row_idx}">')
        for idx, value in enumerate(row):
            style = num_style if idx in numeric else body_style
            parts.append(_cell_xml(f"{_column_letters(idx)}{row_idx}", style, value, numeric=idx in numeric))
        parts.append("</row>")
    parts.append("</sheetData>")
    parts.append(f'<autoFilter ref="A1:{last_col}{len(data_rows) + 1}"/>')
    parts.append("</worksheet>")
    return "".join(parts)


def xlsx_parts(
    sheet_name: str,
    header: Sequence[str],
    rows: Iterable[Sequence[Any]],
    widths: Optional[Sequence[int]] = None,
    numeric_cols: Sequence[int] = (),
    banded: bool = True,
) -> List[Tuple[str, str]]:
    """组装 XLSX 包内的全部部件（包内文件名 → XML 内容），便于写出与检查。"""
    content_types = (
        _XML_DECL
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        + '<Default Extension="xml" ContentType="application/xml"/>'
        + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        + 'spreadsheetml.sheet.main+xml"/>'
        + '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-'
        + 'officedocument.spreadsheetml.worksheet+xml"/>'
        + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.'
        + 'spreadsheetml.styles+xml"/>'
        + "</Types>"
    )
    root_rels = (
        _XML_DECL
        + f'<Relationships xmlns="{_PKG_REL_NS}">'
        + f'<Relationship Id="rId1" Type="{_XLSX_REL_NS}/officeDocument" Target="xl/workbook.xml"/>'
        + "</Relationships>"
    )
    workbook = (
        _XML_DECL
        + f'<workbook xmlns="{_XLSX_NS}" xmlns:r="{_XLSX_REL_NS}">'
        + f'<sheets><sheet name="{_xml_text(safe_sheet_name(sheet_name))}" sheetId="1" r:id="rId1"/></sheets>'
        + "</workbook>"
    )
    workbook_rels = (
        _XML_DECL
        + f'<Relationships xmlns="{_PKG_REL_NS}">'
        + f'<Relationship Id="rId1" Type="{_XLSX_REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>'
        + f'<Relationship Id="rId2" Type="{_XLSX_REL_NS}/styles" Target="styles.xml"/>'
        + "</Relationships>"
    )
    return [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("xl/workbook.xml", workbook),
        ("xl/_rels/workbook.xml.rels", workbook_rels),
        ("xl/styles.xml", _styles_xml()),
        ("xl/worksheets/sheet1.xml", _sheet_xml(header, rows, widths, numeric_cols, banded)),
    ]


def write_xlsx_file(
    path: Any,
    sheet_name: str,
    header: Sequence[str],
    rows: Iterable[Sequence[Any]],
    widths: Optional[Sequence[int]] = None,
    numeric_cols: Sequence[int] = (),
    banded: bool = True,
):
    """写出 XLSX 清单文件，返回 Path。

    - ``widths`` 为各列宽度（字符数）：给足宽度，打开时不会挤成一团；
    - ``numeric_cols`` 中的列按数字右对齐（如“行号”）；
    - 首行冻结并带自动筛选，表头深蓝底白字，数据区隔行浅蓝斑马纹。
    """
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, content in xlsx_parts(sheet_name, header, rows, widths, numeric_cols, banded):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            zf.writestr(info, content)
    return path


def truncate_text(text: Any, limit: int = 200) -> str:
    """压缩空白并截断过长文本，明确标注已截断（不虚构内容）。"""
    if text is None:
        return ""
    s = " ".join(str(text).split())
    if len(s) <= limit:
        return s
    return s[:limit] + f"…（已截断，原文共 {len(s)} 字符）"


def window_text(text: Any, start: int, end: int, radius: int = 60) -> str:
    """以 ``[start, end)`` 为中心截取文本窗口（两侧各 ``radius`` 字符，超出部分标省略号）。

    用于“证据原文”定位：长行里命中位置靠后时，只截行首会把出问题的那段文字截掉，
    证据里看不到关键内容；以命中处为中心截取可保证“证据一定包含问题本身”。
    """
    s = "" if text is None else str(text)
    total = len(s)
    if total == 0:
        return ""
    start = max(0, min(int(start), total))
    end = max(start, min(int(end), total))
    left = max(0, start - radius)
    right = min(total, end + radius)
    window = s[left:right]
    return f"{'…' if left > 0 else ''}{window}{'…' if right < total else ''}"


def count_by(values: Iterable[Any]) -> Dict[str, int]:
    """统计可迭代对象中各值出现的次数（键统一转为字符串）。"""
    result: Dict[str, int] = {}
    for v in values:
        key = str(v)
        result[key] = result.get(key, 0) + 1
    return result


# --------------------------------------------------------------------------- 进度显示
# 纯标准库进度条：交互终端原地刷新一行，重定向 / CI 日志自动降级为分段整行输出；
# 只用色块字符（█ ░）与数字，不用 ANSI 转义，旧版 Windows 控制台也能正常显示。
_PROGRESS_FILL = "█"
_PROGRESS_EMPTY = "░"
_PROGRESS_MIN_INTERVAL = 0.1  # 秒：刷新节流，避免打印拖慢主流程
_PROGRESS_LOG_STEP = 5  # 非交互终端下每完成百分之几打印一行
_PROGRESS_LOG_EVERY_WHEN_UNKNOWN = 1000  # 总数未知（非交互终端）时每多少项打印一行


def format_duration(seconds: float) -> str:
    """秒 → ``HH:MM:SS``（用于进度条的已用时间与预计剩余时间）。"""
    total = int(max(0.0, seconds))
    return f"{total // 3600:02d}:{total % 3600 // 60:02d}:{total % 60:02d}"


def _ensure_utf8_stdout(stream: Any = None) -> None:
    """把输出流切到 UTF-8（中文与进度条色块字符在旧版 Windows 控制台下才不会报编码错误）。"""
    stream = stream or sys.stdout
    reconfigure = getattr(stream, "reconfigure", None)
    if reconfigure is None:
        return
    try:
        reconfigure(encoding="utf-8", errors="replace")
    except (ValueError, OSError):
        pass


def progress_enabled(stream: Any = None, override: Optional[bool] = None) -> bool:
    """是否显示动态进度条：优先取显式设定，否则要求输出流连着交互式终端。"""
    if override is not None:
        return bool(override)
    stream = stream or sys.stdout
    isatty = getattr(stream, "isatty", None)
    try:
        return bool(isatty and isatty())
    except (ValueError, OSError):
        return False


class Progress:
    """终端进度条（标准库实现，可用于 7 万+ 文件的长任务，全程只占一行）。

    用法::

        bar = Progress("质量审查", len(files))
        for index, fp in enumerate(files, 1):
            ...
            bar.update(index, detail=Path(fp).name, extra="已发现 3,201 处问题")
        bar.finish(extra="检查 74,240 个文件，发现 3,201 处问题")

    交互终端下的显示形态::

        [质量审查] ██████████░░░░░░░░░░ 58.3%  43,221/74,240  剩余约 00:14:02  已发现 3,201 处问题

    行为约定：

    - 交互终端：``\\r`` 原地刷新同一行，``finish()`` 补一个换行；
    - 非交互终端（重定向到文件、CI 日志）：不写 ``\\r``，每完成 ``log_step`` 个百分点打印一行；
    - ``total`` 为 0 / None 时退化为“已完成数量 + 用时”，不显示百分比与预计剩余；
    - 刷新有节流（默认 0.1 秒），不会因为打印而拖慢扫描；``update`` 在 100% 时不打印，
      最终汇总行统一由 ``finish()`` 给出。
    """

    def __init__(
        self,
        label: str,
        total: Optional[int] = None,
        *,
        stream: Any = None,
        enabled: Optional[bool] = None,
        width: Optional[int] = None,
        log_step: int = _PROGRESS_LOG_STEP,
    ):
        self.label = label
        self.total = max(0, int(total or 0))
        self.stream = stream or sys.stdout
        # enabled 三态：None=自动（交互终端显示动态条，重定向 / CI 降级为分段整行日志）；
        #              True=强制动态条；False=完全关闭（不打印任何进度）。
        if enabled is False:
            self.quiet = True
            self.enabled = False
        else:
            self.quiet = False
            self.enabled = progress_enabled(self.stream, enabled)
        self.log_step = max(1, int(log_step))
        columns = shutil.get_terminal_size((120, 25)).columns
        self.width = max(10, min(32, int(width) if width else columns - 62))
        self.started = time.monotonic()
        self.done = 0
        self._last_flush = 0.0
        self._last_logged = -self.log_step
        self._last_tail = ""
        if self.enabled:
            _ensure_utf8_stdout(self.stream)

    def _bar(self, ratio: float) -> str:
        filled = int(round(self.width * ratio))
        return _PROGRESS_FILL * filled + _PROGRESS_EMPTY * (self.width - filled)

    def _line(self, done: int, detail: str, extra: str, final: bool) -> str:
        elapsed = time.monotonic() - self.started
        parts = [f"[{self.label}]"]
        if self.total:
            ratio = max(0.0, min(1.0, done / self.total))
            parts.append(self._bar(ratio))
            parts.append(f"{ratio * 100:5.1f}%  {done:,}/{self.total:,}")
            if not final and 0 < done < self.total:
                parts.append(f"剩余约 {format_duration(elapsed / done * (self.total - done))}")
        else:
            parts.append(f"已完成 {done:,} 项")
        if final or not self.total:
            parts.append(f"用时 {format_duration(elapsed)}")
        tail = extra or detail
        if tail:
            parts.append(tail)
        return "  ".join(parts)

    def _write(self, text: str, newline: bool = False) -> None:
        suffix = "\n" if newline else ""
        try:
            self.stream.write(text + suffix)
            self.stream.flush()
        except UnicodeEncodeError:
            try:
                self.stream.write(text.encode("ascii", "replace").decode("ascii") + suffix)
            except (ValueError, OSError):
                pass
        except (ValueError, OSError, AttributeError):
            pass

    def update(self, done: int, detail: str = "", extra: str = "") -> None:
        """推进进度：``done`` 为已完成数量，``detail`` 为当前处理项，``extra`` 为附加说明。"""
        if self.quiet:
            return
        self.done = max(0, int(done))
        if extra:
            self._last_tail = extra
        tail = extra or self._last_tail
        if not self.enabled:
            if self.total:
                if self.done >= self.total:
                    return  # 100% 的汇总行交给 finish()
                ratio = self.done / self.total * 100.0
                if ratio - self._last_logged < self.log_step:
                    return
                self._last_logged = ratio
            elif self.done % _PROGRESS_LOG_EVERY_WHEN_UNKNOWN:
                return
            self._write(self._line(self.done, detail, tail, final=False), newline=True)
            return
        now = time.monotonic()
        if self.done < self.total and now - self._last_flush < _PROGRESS_MIN_INTERVAL:
            return
        self._last_flush = now
        self._write("\r" + self._line(self.done, detail, tail, final=False))

    def finish(self, detail: str = "", extra: str = "") -> None:
        """收尾：打印最终行（交互终端补换行，保证后续输出不会被覆盖）。"""
        if self.quiet:
            return
        line = self._line(self.done, detail if not extra else "", extra or detail or self._last_tail, final=True)
        if self.enabled:
            self._write("\r" + line + "\n")
        else:
            self._write(line, newline=True)


def safe_print(message: str) -> None:
    """向控制台安全打印：必要时把 stdout 切换为 UTF-8，失败时替换不可编码字符。"""
    try:
        if hasattr(sys.stdout, "reconfigure"):
            try:
                sys.stdout.reconfigure(encoding="utf-8", errors="replace")
            except (ValueError, OSError):
                pass
        print(message)
    except UnicodeEncodeError:
        print(str(message).encode("ascii", errors="replace").decode("ascii"))


def print_error(message: str) -> int:
    """打印错误信息并返回退出码 1（供 CLI main 使用）。"""
    safe_print(f"[错误] {message}")
    return 1

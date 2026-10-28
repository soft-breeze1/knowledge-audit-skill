#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""audit_markdown.py — 知识库 Markdown 文档解析与转写质量审查（第一阶段，只读审计）。

输入：Markdown 文件夹（递归扫描 `.md` / `.markdown`）或单个 Markdown 文件；不需要原始 DOCX/PDF/图片。
每次运行产出两份交付文件（写入同一个目录，不修改任何输入文件）：

    <报告基名>.html   单文件 HTML 报告（内联样式、无外部资源，双击即可打开或转发）
    <报告基名>.xlsx   逐行问题清单表格（文件、行号、问题类型、问题说明、证据原文）

检查口径 = 固定 6 项检查（口径见 CHECK_ITEMS）：
    编码可解码性（UTF-8） / 替换字符 U+FFFD / 编码乱码特征 / OCR 标记残留 / 占位符文本 / 空文件检测

每处问题都能定位到“文件 + 行号 + 证据原文”；文件读取失败、超过单文件大小上限属“漏审风险”，
不算质量问题，仅在控制台提示并列入 HTML 报告的运行信息。全流程不生成 Markdown / JSON 文件。

用法：
    python scripts/audit_markdown.py input/
    python scripts/audit_markdown.py input/example.md --output out

退出码：0=正常完成（可能发现问题）；1=输入/输出错误；2=参数或规则文件错误；3=发现问题且 --fail-on-issue。
"""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Dict, Iterator, List, Optional, Set, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import (  # noqa: E402
    TOOL_VERSION,
    Progress,
    count_by,
    make_timestamped_dir,
    now_local_str,
    now_utc_iso,
    print_error,
    relative_path,
    safe_print,
    truncate_text,
    window_text,
    write_text_file,
    write_xlsx_file,
)

from html_report import (  # noqa: E402
    build_html,
    details,
    esc,
    kpi_cards,
    section,
    table,
)

TOOL_NAME = "audit_markdown"
SUPPORTED_EXTENSIONS = (".md", ".markdown")
DEFAULT_MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB
# 默认排除的目录名（不区分大小写）；隐藏目录（以 . 开头）一律排除
DEFAULT_EXCLUDED_DIR_NAMES = {"__pycache__", "node_modules", "venv"}
# 本工具生成的报告文件基名（扫描目录时按文件名精确排除，防止把报告当作输入）
DEFAULT_REPORT_BASE_NAME = "知识库文档解析与转写质量问题报告"
REPORT_SUFFIXES = (".html", ".xlsx", ".md")  # 交付报告后缀：按文件名排除，避免把报告本身当作输入重复审查
# --- HTML 报告的呈现上限（保持简洁可读：高频同型问题按类型聚合呈现） ---
HTML_MAX_ISSUES_PER_GROUP = 5  # 每个「文件 + 问题类型」在 HTML 明细里最多逐条列出几条
HTML_MAX_NOISE_TYPES = 6  # 最多把几种“高频同一型问题”改为聚合呈现
NOISE_MIN_COUNT = 50  # “高频同型”的绝对数量门槛
NOISE_MIN_RATIO = 0.2  # “高频同型”占全部问题数的比例门槛
TOP_FILES_LIMIT = 10  # “问题最多的文件”列出条数
# 清单表格（XLSX）表头、列宽与数字列
SHEET_NAME = "问题清单"
SHEET_HEADER: Tuple[str, ...] = ("文件", "行号", "问题类型", "问题说明", "证据原文")
SHEET_WIDTHS: Tuple[int, ...] = (36, 8, 30, 46, 46)
SHEET_NUMERIC_COLS: Tuple[int, ...] = (1,)  # 行号按数字右对齐

# 检查口径：固定 6 项检查，分类即检查项（报告与清单表格的「问题类型」就是这里的名称）——
# 6 项全部计入「存在问题文件数 / 问题总数」；结构类写法不上报，读取失败 / 超限属“漏审风险”，只在控制台提示并列入报告运行信息。
CHECK_ITEMS: Tuple[Tuple[str, str, str], ...] = (
    (
        "encoding_error",
        "编码可解码性（UTF-8）",
        "文件能否按 UTF-8 解码；解码失败的字节就是损坏处，内容可能已丢失",
    ),
    (
        "replacement_char",
        "替换字符 U+FFFD",
        "解码失败时留下的占位字符，原内容已丢失",
    ),
    (
        "mojibake",
        "编码乱码特征",
        "乱码特征字符，如“锟斤拷”或 UTF-8 被按 Latin-1 打开后的形态（规则见 rules/ocr_patterns.json）",
    ),
    (
        "ocr_marker",
        "OCR 标记残留",
        "转写时留下的 [OCR]、[图片]、[无法识别] 一类标记（规则见 rules/ocr_patterns.json）",
    ),
    (
        "placeholder",
        "占位符文本",
        "TODO、待补充、此处省略 等未补全文本（规则见 rules/ocr_patterns.json）",
    ),
    (
        "empty_file",
        "空文件检测",
        "0 字节文件，可能是转写失败或占位文件",
    ),
)

# 检查项名称（报告与清单表格的「问题类型」列直接用这个名称）
CHECK_NAMES: Dict[str, str] = {type_id: name for type_id, name, _desc in CHECK_ITEMS}


def checks_summary() -> str:
    """6 项检查的一句话摘要（用于报告说明与控制台提示）。"""
    names = "、".join(CHECK_NAMES[type_id] for type_id, _name, _desc in CHECK_ITEMS)
    return f"共 {len(CHECK_ITEMS)} 项检查：{names}"


def check_items_scope() -> List[Dict[str, str]]:
    """检查项清单（报告“范围与限制说明”与运行信息用）：[{id, name, description}, ...]。"""
    return [{"id": type_id, "name": name, "description": desc} for type_id, name, desc in CHECK_ITEMS]


@dataclass
class Issue:
    """一条审查问题记录。"""

    issue_id: str
    file_path: str  # 相对输入根目录的路径（POSIX 风格）；单文件输入时为文件名
    absolute_path: str
    issue_type: str
    line_number: Optional[int]  # 1 起始；None 表示无法定位到具体行
    evidence: str
    explanation: str
    suggested_action: str

    def to_dict(self) -> Dict[str, Any]:
        return {
            "issue_id": self.issue_id,
            "file_path": self.file_path,
            "absolute_path": self.absolute_path,
            "issue_type": self.issue_type,
            "issue_type_label": CHECK_NAMES.get(self.issue_type, self.issue_type),
            "line_number": self.line_number,
            "evidence": self.evidence,
            "explanation": self.explanation,
            "suggested_action": self.suggested_action,
        }


@dataclass
class FileResult:
    """单个文件的审查结果。"""

    path: Path
    rel_path: str
    size: int = 0
    issues: List[Issue] = field(default_factory=list)
    skipped: bool = False  # 因超大/不可读而跳过内容级检查（不进问题统计，只在控制台提示并列入报告运行信息）
    skip_reason: str = ""


# 内置兜底规则：当 rules/ocr_patterns.json 缺失或不可用时使用（保持精简）。
BUILTIN_RULES: Dict[str, List[Dict[str, Any]]] = {
    "ocr_markers": [
        {
            "id": "ocr-bracket-en",
            "pattern": r"(?<!!)\[\s*OCR\s*\](?!\()",
            "flags": ["IGNORECASE"],
            "description": "英文方括号 OCR 标记残留（如 [OCR]）",
            "suggested_action": "对照原始文档，将 OCR 标记替换为真实内容，或删除该标记。",
        },
        {
            "id": "ocr-image-placeholder",
            "pattern": r"(?<!!)\[\s*(?:图片|图像|插图|照片|示意图)\s*[^\]]{0,20}\](?!\()",
            "description": "图片占位标记残留（如 [图片]；不匹配 ![...](...) 图片引用与 [...] (...) 链接语法）",
            "suggested_action": "对照原始文档补回真实图片引用，或删除该占位标记。",
        },
    ],
    "placeholders": [
        {
            "id": "todo",
            "pattern": r"\bTODO\b",
            "flags": ["IGNORECASE"],
            "description": "TODO 占位符",
            "suggested_action": "确认该处是有意留待补充，还是转写时遗漏了内容。",
        },
        {
            "id": "wait-supplement",
            "pattern": "待补充",
            "description": "“待补充”占位文本",
            "suggested_action": "确认该处是有意留空还是内容缺失。",
        },
    ],
    "mojibake": [
        {
            "id": "gbk-mojibake",
            "pattern": "锟斤拷|锟斤|烫烫烫|屯屯屯",
            "description": "GBK/UTF-8 编码转换乱码特征字符",
            "suggested_action": "该文本在编码转换中已损坏，需从原始文档重新转写此段内容。",
        },
    ],
}

# 规则类别 → issue_type 映射
CATEGORY_ISSUE_TYPE = {
    "ocr_markers": "ocr_marker",
    "placeholders": "placeholder",
    "mojibake": "mojibake",
}


@dataclass
class PatternRule:
    """一条已编译的文本异常规则。"""

    rule_id: str
    issue_type: str
    regex: Any  # re.Pattern
    description: str
    suggested_action: str


def _compile_flags(flag_names: Any) -> int:
    """把 flags 写法解析为 re 标志位。

    宽容写法：``["IGNORECASE", "MULTILINE"]``（推荐）、``"IGNORECASE|MULTILINE"``、
    ``"IGNORECASE, MULTILINE"``、单个 ``"IGNORECASE"`` 与 ``None`` 都接受。
    """
    if flag_names is None:
        names: List[Any] = []
    elif isinstance(flag_names, str):
        names = [part for part in re.split(r"[|,\s]+", flag_names) if part]
    elif isinstance(flag_names, (list, tuple, set, frozenset)):
        names = list(flag_names)
    else:
        raise ValueError(f"flags 写法不支持：{flag_names!r}（可用：IGNORECASE、MULTILINE）")
    flags = 0
    for name in names:
        name_upper = str(name).strip().upper()
        if not name_upper:
            continue
        if name_upper == "IGNORECASE":
            flags |= re.IGNORECASE
        elif name_upper == "MULTILINE":
            flags |= re.MULTILINE
        else:
            raise ValueError(f"不支持的正则标志：{name}（可用：IGNORECASE、MULTILINE）")
    return flags


def _parse_rules(raw_rules: Dict[str, Any], notes: List[str]) -> List[PatternRule]:
    """把 ``{类别: [规则, ...]}`` 编译为规则列表。

    单条规则有问题（缺 pattern、flags 写法不支持、正则编译失败）只跳过该条并把原因写入 notes，
    不影响同一类别里的其余规则，也不再让整条流水线中止。
    """
    rules: List[PatternRule] = []
    for category, entries in raw_rules.items():
        issue_type = CATEGORY_ISSUE_TYPE.get(str(category))
        if issue_type is None:
            notes.append(f"未识别的规则类别 {category}（可用：{'、'.join(sorted(CATEGORY_ISSUE_TYPE))}）")
            continue
        if not isinstance(entries, list):
            notes.append(f"规则类别 {category} 不是数组")
            continue
        for entry in entries:
            if not isinstance(entry, dict) or not str(entry.get("pattern") or "").strip():
                notes.append(f"规则类别 {category} 中有条目缺少 pattern")
                continue
            try:
                regex = re.compile(str(entry["pattern"]), _compile_flags(entry.get("flags")))
            except (re.error, ValueError) as exc:
                label = entry.get("id") or category
                notes.append(f"规则 {label} 不可用（{exc}）")
                continue
            rules.append(
                PatternRule(
                    rule_id=str(entry.get("id", category)),
                    issue_type=issue_type,
                    regex=regex,
                    description=str(entry.get("description", issue_type)),
                    suggested_action=str(entry.get("suggested_action", "对照原始文档核实该处内容。")),
                )
            )
    return rules


def load_rules(rules_path: Optional[Path], explicit: bool) -> Tuple[List[PatternRule], str]:
    """加载文本异常规则，返回 (规则列表, 来源说明)。

    宽容策略（一条规则写错不让整条流水线停摆）：
    - 单条规则缺 pattern / flags 写法不支持 / 正则编译失败 → 只跳过该条，原因写入来源说明；
    - 未识别的类别、非数组的类别 → 跳过并写入来源说明，不再静默丢弃；
    - 文件缺失 / 无法解析 / 一条规则都加载不到：默认路径回退内置兜底规则并在报告中注明原因；
      显式指定 ``--rules`` 时抛异常（CLI 以退出码 2 结束）。
    """
    notes: List[str] = []
    if rules_path is None or not Path(rules_path).is_file():
        if explicit:
            raise FileNotFoundError(f"规则文件不存在：{rules_path}")
        source = "内置兜底规则（未找到 rules/ocr_patterns.json）"
        data: Any = {"rules": BUILTIN_RULES}
    else:
        source = str(Path(rules_path))
        try:
            with open(Path(rules_path), "r", encoding="utf-8-sig") as f:
                data = json.load(f)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
            if explicit:
                raise ValueError(f"规则文件无法解析：{rules_path}（{exc}）")
            data = {"rules": BUILTIN_RULES}
            source = f"内置兜底规则（{rules_path} 无法解析：{exc}）"

    raw_rules: Any = data.get("rules", {}) if isinstance(data, dict) else {}
    if isinstance(data, dict) and not raw_rules and any(k in CATEGORY_ISSUE_TYPE for k in data):
        # 少写了 "rules" 外层、顶层直接是类别分组：按该写法解析，避免静默加载 0 条规则
        notes.append("未找到 rules 外层，已按“顶层直接是类别分组”解析")
        raw_rules = data
    if not isinstance(raw_rules, dict):
        if explicit:
            raise ValueError(f"规则文件格式错误：rules 应为按类别分组的对象（{rules_path}）")
        notes.append("rules 不是按类别分组的对象")
        raw_rules = {}

    rules = _parse_rules(raw_rules, notes)

    if not rules and "内置兜底规则" not in source:
        # 文件存在却一条规则都没加载到：宁可回退内置规则并写明原因，也不要静默少做 3 项检查
        reason = "；".join(notes) if notes else "rules 为空"
        if explicit:
            raise ValueError(f"规则文件未提供任何可用规则：{rules_path}（{reason}）")
        source = f"内置兜底规则（{rules_path} 未提供可用规则：{reason}）"
        notes = []
        rules = _parse_rules(BUILTIN_RULES, [])

    if notes:
        source = f"{source}（{'；'.join(notes)}）"
    return rules, source


# ---------- 基础正则 ----------
RE_FENCE = re.compile(r"^ {0,3}(`{3,}|~{3,})(.*)$")
# 替换字符的建议操作
_REPLACEMENT_CHAR_ACTION = "该字符是解码失败的占位符，内容已丢失；需对照原始文档重新转写该处。"


# 代码块外检查全部规则类别；代码块内只检查乱码类规则（避免把示例代码误报为占位符）
RULE_TYPES_OUTSIDE_FENCE = {"ocr_marker", "placeholder", "mojibake"}
RULE_TYPES_INSIDE_FENCE = {"mojibake"}


def _is_table_line(line: str) -> bool:
    s = line.strip()
    return len(s) >= 2 and s.startswith("|") and s.endswith("|")


def split_table_cells(line: str) -> List[str]:
    """按未转义的 | 拆分表格行单元格（去掉首尾管道符）。"""
    s = line.strip()
    if s.startswith("|"):
        s = s[1:]
    if s.endswith("|"):
        s = s[:-1]
    cells: List[str] = []
    cur: List[str] = []
    i = 0
    while i < len(s):
        ch = s[i]
        if ch == "\\" and i + 1 < len(s) and s[i + 1] == "|":
            cur.append("|")
            i += 2
            continue
        if ch == "|":
            cells.append("".join(cur))
            cur = []
            i += 1
            continue
        cur.append(ch)
        i += 1
    cells.append("".join(cur))
    return cells


def _is_separator_cell(cell: str) -> bool:
    return re.fullmatch(r"\s*:?-{1,}:?\s*", cell) is not None


def is_valid_separator_line(line: str) -> bool:
    """是否为合法的 GFM 表格分隔行（如 | --- | :---: |）。"""
    if not _is_table_line(line):
        return False
    cells = split_table_cells(line)
    return bool(cells) and all(_is_separator_cell(c) for c in cells) and any("-" in c for c in cells)


def iter_table_blocks(lines: List[str]) -> Iterator[Tuple[int, List[str], List[Tuple[int, List[str]]]]]:
    """迭代 Markdown 表格块（供属性抽取复用）。

    仅识别“首尾都有管道符”的 GFM 风格表格（表头行 + 紧随的合法分隔行）。
    产出：(表头行号(1 起始), 表头单元格列表, [(数据行行号, 数据行单元格列表), ...])
    """
    n = len(lines)
    i = 0
    while i < n:
        line = lines[i]
        if _is_table_line(line) and i + 1 < n and is_valid_separator_line(lines[i + 1]):
            header_cells = split_table_cells(line)
            rows: List[Tuple[int, List[str]]] = []
            j = i + 2
            while j < n and lines[j].strip() and _is_table_line(lines[j]):
                rows.append((j + 1, split_table_cells(lines[j])))
                j += 1
            yield (i + 1, header_cells, rows)
            i = j
            continue
        i += 1


def _describe_decode_error(raw: bytes, exc: UnicodeDecodeError) -> str:
    """生成编码错误的证据描述（只报告可观察事实，不虚构内容）。"""
    hints = []
    if raw.startswith(b"\xff\xfe"):
        hints.append("文件开头为 FF FE，疑似 UTF-16 LE 编码")
    elif raw.startswith(b"\xfe\xff"):
        hints.append("文件开头为 FE FF，疑似 UTF-16 BE 编码")
    null_count = raw.count(b"\x00")
    if null_count > max(10, len(raw) // 100):
        hints.append(f"含 {null_count} 个 0x00 字节，疑似 UTF-16 或二进制内容")
    head = " ".join(f"{b:02X}" for b in raw[:16])
    bad = raw[exc.start:exc.start + 1].hex() if exc.start < len(raw) else "?"
    base = f"文件前 16 字节：{head}；解码在第 {exc.start} 字节附近失败（无效字节 {bad}）"
    if hints:
        base += "；" + "；".join(hints)
    return base


def scan_file(
    path: Path,
    rel_path: str,
    rules: List[PatternRule],
    max_file_size: int = DEFAULT_MAX_FILE_SIZE,
    rule_match_cap: int = 100,
) -> FileResult:
    """扫描单个 Markdown 文件，返回审查结果（issue_id 由调用方统一编号）。"""
    result = FileResult(path=path, rel_path=rel_path)
    collected: List[Dict[str, Any]] = []

    def add(issue_type, line_number, evidence, explanation, suggested_action):
        collected.append(
            {
                "issue_type": issue_type,
                "line_number": line_number,
                "evidence": truncate_text(evidence),
                "explanation": explanation,
                "suggested_action": suggested_action,
            }
        )

    def finish() -> FileResult:
        result.issues = [
            Issue(
                issue_id="",
                file_path=rel_path,
                absolute_path=str(path),
                issue_type=item["issue_type"],
                line_number=item["line_number"],
                evidence=item["evidence"],
                explanation=item["explanation"],
                suggested_action=item["suggested_action"],
            )
            for item in collected
        ]
        return result

    # --- 读取与编码检查 ---
    # 读取失败 / 超过单文件大小上限 = “漏审风险”，不是质量问题：不计入统计，只在控制台提示并列入报告运行信息（见 run()）。
    try:
        result.size = path.stat().st_size
    except OSError as exc:
        result.skipped = True
        result.skip_reason = f"无法读取文件属性（{exc}）"
        return finish()

    if result.size == 0:
        add(
            "empty_file",
            None,
            "（空文件，0 字节）",
            "文件为空，可能是转写失败或占位文件。",
            "确认该文件是否应有内容；如有原始文档，请重新转写。",
        )
        return finish()

    if result.size > max_file_size:
        result.skipped = True
        result.skip_reason = f"超过单文件大小上限（{result.size} 字节 > {max_file_size} 字节，未做内容级检查）"
        return finish()

    try:
        raw = path.read_bytes()
    except OSError as exc:
        result.skipped = True
        result.skip_reason = f"无法读取文件内容（{exc}）"
        return finish()

    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        add(
            "encoding_error",
            None,
            _describe_decode_error(raw, exc),
            "文件无法按 UTF-8 解码，已改用替换模式继续检查（下文的“替换字符”问题即由此产生）。",
            "请将文件转为 UTF-8 编码保存，或从原始文档重新转写。",
        )
        text = raw.decode("utf-8", errors="replace")

    lines = text.splitlines()
    n = len(lines)

    # --- 逐行状态机（代码围栏状态）---
    fence_open = False
    fence_char = ""
    fence_len = 0
    rule_hit_counts: Dict[str, int] = defaultdict(int)
    rule_capped_counts: Dict[str, int] = defaultdict(int)
    rule_type_by_id = {r.rule_id: r.issue_type for r in rules}

    def run_rule_checks(line: str, line_no: int, allowed_types: Set[str]) -> None:
        for rule in rules:
            if rule.issue_type not in allowed_types:
                continue
            for _m in rule.regex.finditer(line):
                rule_hit_counts[rule.rule_id] += 1
                if rule_hit_counts[rule.rule_id] > rule_match_cap:
                    rule_capped_counts[rule.rule_id] += 1
                    continue
                add(
                    rule.issue_type,
                    line_no,
                    # 以命中位置为中心截取：长行里命中点靠后时，只截行首会把证据本身截掉
                    window_text(line, _m.start(), _m.end()),
                    f"命中规则 [{rule.rule_id}]：{rule.description}。"
                    f"命中片段：“{truncate_text(_m.group(0), 80)}”。",
                    rule.suggested_action,
                )

    def check_chars(line: str, line_no: int) -> None:
        """替换字符 U+FFFD 检查（解码失败的占位符，内容已丢失）。"""
        count = line.count(chr(0xFFFD))
        if count:
            first = line.index(chr(0xFFFD))
            add(
                "replacement_char",
                line_no,
                window_text(line, first, first + 1),
                f"发现 {count} 个替换字符 U+FFFD（通常意味着该处编码损坏或转写失败，内容已丢失）。",
                _REPLACEMENT_CHAR_ACTION,
            )

    # 表格解析工具保留在模块级（iter_table_blocks 供属性抽取复用），本工具不做表格结构检查

    i = 0
    while i < n:
        line = lines[i]
        line_no = i + 1
        stripped = line.strip()

        # 1) 代码围栏状态
        fence_m = RE_FENCE.match(line)
        opens_fence = bool(fence_m) and not (fence_m.group(1)[0] == "`" and "`" in fence_m.group(2))
        if fence_open:
            if (
                opens_fence
                and fence_m.group(1)[0] == fence_char
                and len(fence_m.group(1)) >= fence_len
                and not fence_m.group(2).strip()
            ):
                fence_open = False
            else:
                check_chars(line, line_no)
                run_rule_checks(line, line_no, RULE_TYPES_INSIDE_FENCE)
            i += 1
            continue
        if opens_fence:
            fence_open = True
            fence_char = fence_m.group(1)[0]
            fence_len = len(fence_m.group(1))
            i += 1
            continue

        # 2) 空行：跳过
        if not stripped:
            i += 1
            continue

        # 3) 替换字符 U+FFFD 检查
        check_chars(line, line_no)

        # 4) 可配置规则检查（OCR 标记 / 占位符 / 乱码特征）
        run_rule_checks(line, line_no, RULE_TYPES_OUTSIDE_FENCE)

        i += 1
    # --- 收尾检查 ---
    # 规则命中次数过多时：只逐条列出前若干处，其余在此汇总（整文件级、无具体行号）
    for rule_id, capped in rule_capped_counts.items():
        add(
            rule_type_by_id.get(rule_id, "placeholder"),
            None,
            f"规则 [{rule_id}] 在本文件命中超过 {rule_match_cap} 处",
            f"规则 [{rule_id}] 命中次数过多，仅逐条列出前 {rule_match_cap} 处，其余 {capped} 处未逐条列出。",
            "该文件问题密度较高，建议对照原始文档整体重新转写。",
        )
    return finish()


def collect_files(input_path: Path, output_dir: Path, report_base_name: str) -> List[Path]:
    """收集待扫描的 Markdown 文件（递归、排序、排除输出目录与无关目录）。

    输出目录只在“位于输入路径内部、且与输入路径不同”时整体排除；若输出目录是输入路径的祖先
    （如 `--output .`，或把报告输出到输入目录的上层），整体排除会把输入文档一并排除，
    此时只按报告文件名排除，避免出现“扫描 0 个文件”的静默漏检。
    """
    if input_path.is_file():
        return [input_path]
    resolved_out = output_dir.resolve()
    try:
        input_root = input_path.resolve()
    except OSError:
        input_root = input_path
    exclude_out_dir = resolved_out != input_root and resolved_out.is_relative_to(input_root)
    excluded_report_names = {f"{report_base_name}{suffix}" for suffix in REPORT_SUFFIXES}

    def dir_excluded(d: Path) -> bool:
        if d.name.startswith("."):
            return True
        if d.name.lower() in DEFAULT_EXCLUDED_DIR_NAMES:
            return True
        try:
            if exclude_out_dir and d.resolve() == resolved_out:
                return True
        except OSError:
            pass
        return False

    files: List[Path] = []
    for root, dirnames, filenames in os.walk(input_path):
        dirnames[:] = sorted(d for d in dirnames if not dir_excluded(Path(root) / d))
        for fn in sorted(filenames):
            if fn in excluded_report_names:
                continue
            fp = Path(root) / fn
            if fp.suffix.lower() not in SUPPORTED_EXTENSIONS:
                continue
            try:
                if exclude_out_dir and resolved_out in fp.resolve().parents:
                    continue
            except OSError:
                pass
            files.append(fp)
    return files


def compute_noise_types(by_type: Dict[str, int], total: int) -> List[str]:
    """挑出“高频同一型问题”类型：数量多且占比高，报告里改为按类型聚合呈现。

    真实知识库中“表格空单元格”“连续空行”这类问题常成百上千处（往往源于同一转写习惯），
    逐条铺开会把真正需要逐条核对的问题淹没；聚合呈现时给出数量与行号汇总，仍可逐处定位。
    """
    if total <= 0:
        return []
    ranked = [
        (count, type_id)
        for type_id, count in by_type.items()
        if count >= NOISE_MIN_COUNT and count / total >= NOISE_MIN_RATIO
    ]
    ranked.sort(key=lambda item: (-item[0], item[1]))
    return [type_id for _count, type_id in ranked[:HTML_MAX_NOISE_TYPES]]


def top_files_by_issues(files_info: List[Dict[str, Any]], limit: int = TOP_FILES_LIMIT) -> List[Dict[str, Any]]:
    """问题数最多的文件（报告“问题汇总”用）；数量相同时按文件路径排序，结果确定。"""
    ranked = sorted(files_info, key=lambda f: (-f["issue_count"], f["file_path"]))
    return ranked[:limit]


def group_issues_by_file(data: Dict[str, Any]) -> Dict[str, Dict[str, List[Dict[str, Any]]]]:
    """把 issues 组织成 {文件: {问题类型: [问题...]}}，供报告聚合呈现。"""
    by_file: Dict[str, Dict[str, List[Dict[str, Any]]]] = {}
    for issue in data["issues"]:
        by_file.setdefault(issue["file_path"], {}).setdefault(issue["issue_type"], []).append(issue)
    return by_file


def compress_line_refs(issues: List[Dict[str, Any]], limit: int = 12) -> str:
    """把多处行号压缩为“12、15-18”这样的简短引用；最多列出 limit 段。

    同一行上有多处问题时只列一次行号（问题处数另有统计），行号为空表示整文件级问题。
    """
    nums = sorted({i["line_number"] for i in issues if i["line_number"]})
    if not nums:
        return "（整文件级问题，无具体行号）"
    parts: List[str] = []
    start = prev = nums[0]
    for num in nums[1:]:
        if num == prev + 1:
            prev = num
            continue
        parts.append(f"{start}-{prev}" if start != prev else str(start))
        start = prev = num
    parts.append(f"{start}-{prev}" if start != prev else str(start))
    text = "、".join(parts[:limit])
    if len(parts) > limit:
        text += f"（共涉及 {len(nums)} 行）"
    return text


def build_report_data(
    input_path: Path,
    output_dir: Path,
    results: List[FileResult],
    rules_source: str,
    max_file_size: int,
    report_base_name: str,
) -> Dict[str, Any]:
    """汇总审查结果：统一编号、统计，组装报告数据（HTML 报告与清单表格共用）。"""
    # 问题统一编号（KA-0001 起）
    all_issues: List[Issue] = []
    for fr in results:
        fr.issues.sort(key=lambda x: (x.line_number is None, x.line_number or 0))
        all_issues.extend(fr.issues)

    for idx, issue in enumerate(all_issues, start=1):
        issue.issue_id = f"KA-{idx:04d}"

    by_type = count_by(i.issue_type for i in all_issues)
    noise_types = compute_noise_types(by_type, len(all_issues))

    files_info: List[Dict[str, Any]] = []
    for fr in results:
        files_info.append(
            {
                "file_path": fr.rel_path,
                "absolute_path": str(fr.path),
                "size_bytes": max(fr.size, 0),
                "issue_count": len(fr.issues),
            }
        )

    skipped_info: List[Dict[str, str]] = [
        {"file_path": fr.rel_path, "reason": fr.skip_reason} for fr in results if fr.skipped
    ]

    data: Dict[str, Any] = {
        "tool": TOOL_NAME,
        "version": TOOL_VERSION,
        "generated_at_utc": now_utc_iso(),
        "generated_at_local": now_local_str(),
        "input": str(input_path.resolve()),
        "output_dir": str(output_dir),
        "rules_source": rules_source,
        "scope": {
            "extensions": list(SUPPORTED_EXTENSIONS),
            "recursive": True,
            "max_file_size_bytes": max_file_size,
            "excluded": {
                "hidden_dirs": True,
                "dir_names": sorted(DEFAULT_EXCLUDED_DIR_NAMES),
                "output_dir": str(output_dir.resolve()),
                "report_names": [f"{report_base_name}{suffix}" for suffix in REPORT_SUFFIXES],
            },
            "checks": check_items_scope(),
            "notes": [
                "检查口径为 " + checks_summary() + "（各项口径见本报告“二、问题分布”）。",
                "报告与清单表格的「问题类型」就是上列检查项名称；6 项检查的问题全部计入"
                "「存在问题文件数 / 问题总数」，并逐行写入清单表格。",
                "仅检查扩展名为 .md / .markdown 的文件；其他格式未检查。",
                "表格列数与表头不一致、表格空单元格、空章节、代码块未闭合、疑似文本截断等结构类写法"
                "不在检查口径内，不作问题上报。",
                "表格与代码块不做结构检查，但其中的文本仍按上列检查项逐行检查；代码块内部只检查"
                "替换字符与编码乱码特征（避免把示例代码误报为占位符 / OCR 标记）。",
                "读取失败或超过单文件大小上限的文件属于“漏审风险”，不是质量问题：不计入"
                "「存在问题文件数 / 问题总数」，其数量与文件名在“五、附：本次运行信息”中列出，便于人工补审。",
                "本工具只读 Markdown 文本，不读取、不对照原始文档（DOCX/PDF/图片）；"
                "无法仅凭文本确认的问题在“说明”中如实写明，不作武断结论。",
                "高频同一型问题（如替换字符成批出现）在 HTML 报告中按类型聚合呈现，并附行号汇总，"
                "可据此在源文件里逐处定位核对；逐行的完整清单见同名 XLSX 表格文件。",
            ],
        },
        "files_scanned": len(results),
        "files_skipped": len(skipped_info),
        "skipped_files": skipped_info,
        "files_with_issues": sum(1 for fr in results if fr.issues),
        "issues_total": len(all_issues),
        "issues_by_type": by_type,
        "noise_issue_types": noise_types,
        "files_top": top_files_by_issues(files_info),
        "files": files_info,
        "issues": [i.to_dict() for i in all_issues],
    }
    return data


def build_sheet_rows(data: Dict[str, Any]) -> List[Tuple[Any, ...]]:
    """把审查结果展开为清单表格的数据行（一行 = 一处问题）。

    与 HTML 报告不同，表格不做聚合省略：每一处问题都单独一行，直接给出
    「文件 / 行号 / 问题类型 / 问题说明 / 证据原文」，便于直接筛出
    “哪个文档的什么地方出现了什么问题”；行号为空表示该问题覆盖整个文件、无法定位到具体行。

    6 项检查的问题全部写入本清单。
    """
    rows: List[Tuple[Any, ...]] = []
    for issue in data["issues"]:
        issue_type = str(issue.get("issue_type") or "")
        rows.append(
            (
                issue.get("file_path") or "",
                issue.get("line_number") if issue.get("line_number") else "",
                # 直接取 Issue.to_dict() 已算好的标签，同一映射不再算两遍
                issue.get("issue_type_label") or issue_type,
                issue.get("explanation") or "",
                issue.get("evidence") or "",
            )
        )
    return rows


def write_issue_sheet(path: Any, data: Dict[str, Any]):
    """把问题清单写为 XLSX 交付文件（表头配色、列宽、冻结首行、自动筛选），返回 Path。"""
    return write_xlsx_file(
        path,
        SHEET_NAME,
        SHEET_HEADER,
        build_sheet_rows(data),
        widths=SHEET_WIDTHS,
        numeric_cols=SHEET_NUMERIC_COLS,
    )


def render_html_report(data: Dict[str, Any]) -> str:
    """把审查结果渲染为单文件 HTML（样式内联、无外部资源，可直接打开或转发）。

    结构：先给审查概况与问题分布（按 6 项检查逐项列出数量与占比），再按文件分组列出问题明细
    （含行号、证据原文、说明与建议操作），然后是范围与限制说明、本次运行信息。
    为便于阅读，高频同型问题按类型聚合呈现（给数量与行号汇总），每个「文件 + 检查项」最多列出
    HTML_MAX_ISSUES_PER_GROUP 条；**逐行的完整清单见同名 XLSX 表格文件**。
    """
    scope = data["scope"]
    noise_types = set(data.get("noise_issue_types") or [])
    issues_total = data["issues_total"]
    kpis = kpi_cards(
        [
            ("扫描文件数", data["files_scanned"], ""),
            ("存在问题文件数", data["files_with_issues"], "hot"),
            ("问题总数", issues_total, "warn" if issues_total else "ok"),
        ]
    )
    profile_note = (
        f"检查口径为 {checks_summary()}；上列统计就是这 {len(CHECK_ITEMS)} 项检查检出的问题，"
        "全部逐行写入同名 XLSX 清单表格。"
    )
    sections = [section("一、审查概况", kpis, profile_note)]

    if issues_total:
        by_type = data["issues_by_type"]
        file_rows = [
            [f["file_path"], f["issue_count"]] for f in (data.get("files_top") or []) if f["issue_count"]
        ]
        check_rows: List[List[Any]] = [
            [
                name,
                by_type.get(type_id, 0),
                f"{by_type.get(type_id, 0) * 100.0 / issues_total:.1f}%",
                desc,
            ]
            for type_id, name, desc in CHECK_ITEMS
        ]
        body = (
            '<div class="note">问题最多的文件：</div>'
            + table(["文件", "问题数"], file_rows, align=["", "num"])
            + f'<div class="note">按检查项分布（共 {len(CHECK_ITEMS)} 项；数量为 0 表示该项未发现问题）：</div>'
            + table(["检查项", "数量", "占比", "口径说明"], check_rows, align=["", "num", "num", ""])
        )
        note_parts: List[str] = []
        if noise_types:
            labels = "、".join(CHECK_NAMES.get(t, t) for t in data["noise_issue_types"])
            note_parts.append(f"高频同一型问题（按类型聚合呈现，给数量与行号汇总）：{labels}。")
        sections.append(section("二、问题分布", body, "".join(note_parts) or None))

        file_blocks: List[str] = []
        for file_path, groups in group_issues_by_file(data).items():
            file_total = sum(len(v) for v in groups.values())
            inner: List[str] = []
            for type_id, issues in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])):
                label = CHECK_NAMES.get(type_id, type_id)
                if type_id in noise_types:
                    inner.append(
                        f'<div class="note"><span class="badge warn">{esc(label)}</span> '
                        f"{len(issues)} 处（行号：{esc(compress_line_refs(issues))}）。</div>"
                    )
                    continue
                rows = [
                    [
                        issue["issue_id"],
                        issue["line_number"] if issue["line_number"] else "整文件",
                        issue["evidence"],
                        issue["explanation"],
                        issue["suggested_action"],
                    ]
                    for issue in issues[:HTML_MAX_ISSUES_PER_GROUP]
                ]
                inner.append(
                    f'<div class="note"><span class="badge">{esc(label)}</span> {len(issues)} 处</div>'
                    + table(["编号", "行号", "证据", "说明", "建议操作"], rows, align=["", "num", "", "", ""])
                )
                if len(issues) > HTML_MAX_ISSUES_PER_GROUP:
                    inner.append(
                        f'<div class="note">另有 {len(issues) - HTML_MAX_ISSUES_PER_GROUP} 处同类问题'
                        f"（行号：{esc(compress_line_refs(issues[HTML_MAX_ISSUES_PER_GROUP:]))}）。</div>"
                    )
            file_blocks.append(details(f"{esc(file_path)}（{file_total} 处）", "".join(inner)))
        sections.append(
            section(
                "三、问题明细",
                "".join(file_blocks),
                f"按文件分组呈现（问题多的类型在前）；每个「文件 + 问题类型」最多列出 "
                f"{HTML_MAX_ISSUES_PER_GROUP} 条，点击文件名可展开。逐行完整清单见同名 XLSX 表格文件。",
            )
        )
    elif data["files_scanned"] == 0:
        sections.append(section("二、问题分布", '<div class="note">未找到可处理的 Markdown 文件。</div>', None))
    else:
        sections.append(
            section(
                "二、问题分布",
                f'<div class="note">本次共检查 {data["files_scanned"]} 个文件，未发现问题。</div>',
                None,
            )
        )

    notes = list(scope["notes"]) + [
        "同名清单表格（.xlsx）逐行列出全部问题：文件、行号、问题类型、问题说明、证据原文；"
        "「问题类型」就是本报告的检查项名称（如「替换字符 U+FFFD」「OCR 标记残留」）；"
        "行号为空表示该问题覆盖整个文件、无法定位到具体行。",
        "本次实际执行的检查项：" + checks_summary() + "。",
    ]
    sections.append(
        section(
            "四、范围与限制说明",
            '<ul class="plain">' + "".join(f"<li>{esc(n)}</li>" for n in notes) + "</ul>",
            None,
        )
    )

    # 五、附：本次运行信息（输入路径、生成时间、规则来源、检查范围、漏审风险）
    info_rows: List[Tuple[str, str]] = [
        ("输入路径", data["input"]),
        ("生成时间", data["generated_at_local"]),
        ("规则来源", data["rules_source"]),
        (
            "检查范围",
            f"扩展名 {'、'.join(scope['extensions'])}；递归扫描；"
            f"单文件上限 {scope['max_file_size_bytes']} 字节",
        ),
    ]
    if data.get("files_skipped"):
        skip_names = "；".join(
            f"{s['file_path']}（{s['reason']}）" for s in data["skipped_files"]
        )
        info_rows.append(
            ("漏审风险", f"{data['files_skipped']} 个文件未做内容级检查（不计入质量问题）：{skip_names}")
        )
    sections.append(
        section(
            "五、附：本次运行信息",
            table(["项目", "说明"], info_rows, align=["", ""]),
            "输入路径为绝对路径，可据此确认本次扫描的语料；单文件上限用于避免超大文件拖慢审查。"
            "读取失败 / 超限的文件仅列为“漏审风险”，不影响问题统计。",
        )
    )
    return build_html(
        "知识库文档解析与转写质量问题报告",
        "只列出 Markdown 文本层面可核查的问题（含证据原文与行号），归入固定 6 项检查；"
        "审查为只读操作，未修改任何输入文件。",
        [
            ("输入", data["input"]),
            ("生成时间", data["generated_at_local"]),
            ("工具", f"{data['tool']} v{data['version']}"),
        ],
        sections,
        "由 knowledge-audit-skill 自动生成（scripts/audit_markdown.py）；逐行完整清单见同名 XLSX 表格文件。",
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="audit_markdown.py",
        description="知识库 Markdown 文档解析与转写质量审查（第一阶段，只读审计）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/audit_markdown.py input/example.md\n"
            "  python scripts/audit_markdown.py input/                     # 输出到 output/<时间戳>/\n"
            "  python scripts/audit_markdown.py input/ --output out\n"
            "\n"
            "报告默认写入 output/ 下带完整时间戳的子目录；显式指定 --output 时直接写入该目录。\n"
            "每次产出两份交付文件：HTML 报告（便于阅读）+ XLSX 清单表格（逐行问题清单，Excel / WPS 直接打开，"
            "表头配色、列宽、冻结首行、自动筛选）；不生成 Markdown / JSON。\n"
            "检查口径固定为 6 项检查（编码可解码性（UTF-8）、替换字符 U+FFFD、编码乱码特征、OCR 标记残留、"
            "占位符文本、空文件检测）；问题全部计入「存在问题文件数 / 问题总数」并逐行写入清单表格；"
            "表格结构、空章节、代码块未闭合、疑似文本截断等写法不在检查口径内。\n"
            "退出码：0=正常完成（可能发现问题）；1=输入/输出错误；2=参数或规则文件错误；"
            "3=发现问题且指定了 --fail-on-issue。"
        ),
    )
    parser.add_argument("input", help="待审查的 Markdown 文件或目录")
    parser.add_argument(
        "--output",
        default=None,
        help="报告输出目录；不指定时在 output/ 下自动新建带时间戳的子目录（YYYYMMDD-HHMMSS）",
    )
    parser.add_argument(
        "--rules",
        default=None,
        help="文本异常规则 JSON 文件（默认：项目内 rules/ocr_patterns.json，缺失时用内置兜底规则）",
    )
    parser.add_argument(
        "--report-name",
        default=DEFAULT_REPORT_BASE_NAME,
        help="报告文件基名（默认：知识库文档解析与转写质量问题报告）",
    )
    parser.add_argument(
        "--max-file-size",
        type=int,
        default=DEFAULT_MAX_FILE_SIZE,
        help="单个文件大小上限（字节，默认 20MB），超过则跳过内容检查",
    )
    parser.add_argument(
        "--fail-on-issue",
        action="store_true",
        help="发现问题时以退出码 3 退出（供 CI 卡口）",
    )
    parser.add_argument("--no-progress", action="store_true", help="关闭进度条（默认在交互式终端显示）")
    return parser


def run(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    input_path = Path(args.input).expanduser()
    if not input_path.exists():
        return print_error(f"输入路径不存在：{input_path}")
    if input_path.is_file() and input_path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        return print_error(f"不支持的文件类型：{input_path}（当前仅支持 {'/'.join(SUPPORTED_EXTENSIONS)}）")

    output_dir = Path(args.output).expanduser() if args.output else Path("output")
    try:
        if not args.output:
            output_dir = make_timestamped_dir(output_dir)
        output_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return print_error(f"无法创建输出目录 {output_dir}：{exc}")
    safe_print(f"[audit_markdown] 输出目录：{output_dir.resolve()}")

    if args.rules:
        rules_path: Optional[Path] = Path(args.rules).expanduser()
    else:
        default_rules = SCRIPT_DIR.parent / "rules" / "ocr_patterns.json"
        rules_path = default_rules if default_rules.is_file() else None
    try:
        rules, rules_source = load_rules(rules_path, explicit=bool(args.rules))
    except (FileNotFoundError, ValueError) as exc:
        safe_print(f"[错误] {exc}")
        return 2

    files = collect_files(input_path, output_dir, args.report_name)
    if not files:
        safe_print(f"[提示] 未在 {input_path} 下找到可处理的 {'/'.join(SUPPORTED_EXTENSIONS)} 文件。")

    bar = Progress("质量审查", len(files), enabled=None if not args.no_progress else False)
    results: List[FileResult] = []
    issues_seen = 0
    for index, fp in enumerate(files, 1):
        try:
            result = scan_file(fp, relative_path(fp, input_path), rules, args.max_file_size)
        except Exception as exc:  # 单个文件出意外（文件被并发删除、异常字节等）不应中断几万篇的整批审查
            result = FileResult(
                path=fp,
                rel_path=relative_path(fp, input_path),
                skipped=True,
                skip_reason=f"扫描异常（{type(exc).__name__}：{exc}），未做内容级检查",
            )
        results.append(result)
        issues_seen += len(result.issues)
        bar.update(index, detail=result.rel_path, extra=f"已发现 {issues_seen:,} 处问题")
    bar.finish(extra=f"检查 {len(files):,} 个文件，发现 {issues_seen:,} 处问题")

    # 漏审风险提示：读取失败 / 超过大小上限的文件不做内容级检查，不计入质量问题，仅在控制台提示并列入 HTML 报告运行信息
    skipped = [fr for fr in results if fr.skipped]
    if skipped:
        safe_print(
            f"[audit_markdown] 漏审风险提示：{len(skipped)} 个文件未做内容级检查（不计入质量问题，仅在控制台提示并列入报告运行信息）："
        )
        for fr in skipped:
            safe_print(f"[audit_markdown]   - {fr.rel_path}：{fr.skip_reason}")

    data = build_report_data(
        input_path, output_dir, results, rules_source, args.max_file_size, args.report_name
    )

    # 写盘兜底：目录只读、磁盘写满、交付文件被 Excel/WPS 占用等都会抛 OSError，这里统一转成清晰错误而不是崩栈
    try:
        html_path = write_text_file(output_dir / f"{args.report_name}.html", render_html_report(data))
        safe_print(f"[audit_markdown] HTML 报告：{html_path}")
        sheet_path = write_issue_sheet(output_dir / f"{args.report_name}.xlsx", data)
    except OSError as exc:
        return print_error(f"无法写入交付文件（目录 {output_dir}）：{exc}")
    safe_print(
        f"[audit_markdown] XLSX 清单表格：{sheet_path}（{data['issues_total']} 行，一行一处问题；"
        f"检查项 {len(CHECK_ITEMS)} 项）"
    )

    safe_print(
        f"[audit_markdown] 扫描完成：检查 {data['files_scanned']} 个文件，"
        f"发现问题 {data['issues_total']} 处。"
    )
    # 自检（对账恒等式）：报告里的文件数 / 问题数必须与本次扫描逐一相等，避免统计口径漂移
    if data["files_scanned"] != len(results) or data["issues_total"] != issues_seen:
        safe_print(
            f"[audit_markdown] [自检] 统计口径异常：文件数 报告 {data['files_scanned']} 对 实际 {len(results)}、"
            f"问题数 报告 {data['issues_total']} 对 实际 {issues_seen}，请复核报告。"
        )
    else:
        safe_print(
            f"[audit_markdown] 自检通过：文件数 {len(results):,}，问题数 {issues_seen:,}，"
            f"跳过（未做内容级检查）{len(skipped):,} 个。"
        )
    if args.fail_on_issue and data["issues_total"] > 0:
        safe_print("[audit_markdown] 因发现问题且指定了 --fail-on-issue，以退出码 3 结束。")
        return 3
    return 0


def main() -> int:
    try:
        return run()
    except KeyboardInterrupt:
        safe_print("[错误] 用户中断操作。")
        return 130


if __name__ == "__main__":
    sys.exit(main())

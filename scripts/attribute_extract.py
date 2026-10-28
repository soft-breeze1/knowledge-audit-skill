#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""attribute_extract.py — 属性抽取引擎（库模块，无命令行入口；调用方为 detect_conflicts.py）。

规则式抽取（非 AI 推断），只识别可客观定位的模式，每条记录带来源定位（文件/行号/证据原文）：
- GFM 表格行：仅抽取“属性-取值”表——两列（| 属性名 | 取值 |），或从第 3 列起含独立“单位”列
  （如 | 参数 | 数值 | 单位 |）；多列对照表、故障排查表、备件/工具清单等（≥3 列且无单位列）整表跳过
- 列表项 / 正文行：“属性名：数值+单位”
- 代码块内部不抽取；复合取值（范围、多维尺寸）不抽取
- 型号（model）依次取自“型号”类属性值、一级标题、文件名中的型号词（须含字母，纯数字不认作型号）；
  识别不出时记 null，该文件的属性不参与跨文档比较
- 标识类字段（物料编码/零件号/图号/序号等）不抽取；额定/最大/最小等工况限定词作为 condition 单独记录
- 单位换算依赖 rules/attribute_aliases.json；未收录单位只记录原文，不猜测量纲

build_extraction() 返回结构化抽取结果（内存对象，不落盘）。
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import Progress, relative_path, truncate_text, window_text  # noqa: E402
from audit_markdown import RE_FENCE, collect_files, iter_table_blocks  # noqa: E402
# 型号类属性名（命中即按型号处理，取值为原始文本而非数值）
MODEL_ATTR_NAMES = {"型号", "产品型号", "设备型号", "销售型号", "型号规格", "机型", "model"}
# 型号属性名的中英混排尾缀（“型号 Model”“型号/Model”），见 is_model_attribute
RE_MODEL_ATTR_TAIL = re.compile(r"^(?:[\s/·\-_]+[A-Za-z]{1,10})?$")
# “属性名/单位”写法（真实规格书高频，如 回转速度/r/min、斗容/m³）：斜杠后须为形状合法的拉丁单位串
RE_ATTR_SLASH_UNIT = re.compile(
    r"^(?P<name>[^/]{1,16})\s*/\s*(?P<unit>[A-Za-zμ°℃%‰³²·][A-Za-z0-9μ°℃%‰³²·/\-]{0,11})$"
)
# 说明性词组：出现即视为正文语句而非参数名
PROSE_ATTR_WORDS = (
    "要求", "满足", "如下", "包括", "例如", "事项", "详见", "参见", "备注", "用于", "内容", "亮点", "优势", "特点",
)
# 表头黑名单词：版本/日期/目录/名称说明等表不是参数表（仅对无“单位”列的表生效）
NON_ATTRIBUTE_HEADER_WORDS = ("版本", "日期", "更新", "变更", "版权", "目录", "序号", "名称", "说明", "备注")
# 型号变型词：紧跟基础型号的变型标记（ABCDEF-100 HEV 与 ABCDEF-100 不是同一台机器，必须并入型号）
RE_MODEL_VARIANT = re.compile(
    r"^[\s\-_/]{0,2}(HEV|PHEV|EV|PRO|ECO|PLUS|LITE|STD|SMART)\b", re.IGNORECASE
)
# 型号词切分：下划线/斜杠/连续横线常用于拼接多个代号（ABCDEF-100_3104-----_ABCDEF-100）
RE_MODEL_SPLIT = re.compile(r"[_/]+|-{2,}")
# “数值+单位”形状的词不是型号（35mm、1.5m³、24kN·m 等）
RE_UNIT_LIKE_TOKEN = re.compile(
    r"^\d+(?:\.\d+)?(?:mm|cm|dm|km|kg|mg|ml|t|kw|w|v|a|l|hz|khz|mhz|mpa|kpa|pa|kn|nm|n|bar|rpm|s|h|m3|m2|%|m)$",
    re.IGNORECASE,
)
# 标识类属性名：物料编码/零件号等属标识符，跨文档比较无意义，抽取时跳过
IDENTIFIER_ATTR_NAMES = {
    "物料编码", "整机物料编码", "物料号", "零件号", "件号", "图号", "序号", "条码", "二维码",
}
# 说明性空话：常被误当作属性名，直接跳过
PROSE_ATTR_NAMES = {"例如", "注意事项", "备注", "提示"}
# 尾部括注工况（“行走速度（高）”“额定功率（ISO 14396）”）：作为 condition 记录，不参与别名合并
RE_TRAILING_QUALIFIER = re.compile(r"[（(](?P<q>[^（）()]{1,10})[）)]\s*$")
# 工况限定词：从属性名前缀剥离，作为 condition 记录，不参与别名合并
CONDITION_QUALIFIERS = ("额定", "最大", "最小", "常用")
RE_LIST_ITEM = re.compile(r"^\s*(?:[-*+]\s+|\d+[.、]\s+)(?P<body>.*)$")
RE_HEADING1 = re.compile(r"^#\s+(?P<title>.*)$")
# 型号词形：字母数字混合且含数字（如 XW-100、KA2000Pro）
RE_MODEL_TOKEN = re.compile(r"[A-Za-z0-9][A-Za-z0-9\-_/]*\d[A-Za-z0-9\-_/]*")
# 行内“属性：值”模式：属性名不含空格/常用标点（≤20 字符），值以数字开头
RE_ATTR_COLON = re.compile(
    r"(?P<attr>[^\s：:|，。；、！!？?（）()\[\]{}【】《》“”\"']{1,20})\s*[:：]\s*(?P<value>[+-]?\d[^，。；、：:]*)"
)
# 型号属性的行内模式（型号值不以数字开头，单独识别，避免与数值模式混淆）
RE_MODEL_COLON = re.compile(r"(?P<attr>型号|产品型号|设备型号|型号规格)\s*[:：]\s*(?P<value>\S.{0,30})")
# 取值后可接受的终止字符（出现其他字符视为取值不完整，放弃该条）
_VALUE_TERMINATORS = "，。；、,.;:： \t（（）)】]|"


def _clean_alias_map(raw: Any, notes: List[str]) -> Dict[str, List[str]]:
    """别名表归一化：值写成单个字符串（如 ``"整机重量"``）也按“一个同义词”处理。

    若不归一化，``for alias in aliases`` 会把字符串逐字符当成别名，静默产生错映射。
    """
    result: Dict[str, List[str]] = {}
    if not isinstance(raw, dict):
        return result
    for canonical, aliases in raw.items():
        if isinstance(aliases, str):
            values: List[Any] = [aliases]
        elif isinstance(aliases, (list, tuple, set, frozenset)):
            values = list(aliases)
        else:
            notes.append(f"别名 {canonical} 的同义词写法不支持（{type(aliases).__name__}），已跳过")
            continue
        cleaned = [str(item).strip() for item in values if str(item or "").strip()]
        if cleaned:
            result[str(canonical)] = cleaned
    return result


def _clean_unit_symbols(raw: Any, notes: List[str]) -> Dict[str, str]:
    """单位写法 → 标准符号：只保留“值是字符串”的条目。"""
    result: Dict[str, str] = {}
    if not isinstance(raw, dict):
        return result
    for written, canonical in raw.items():
        if isinstance(canonical, str) and canonical.strip():
            result[str(written)] = canonical.strip()
        else:
            notes.append(f"单位写法 {written} 的标准符号不是字符串，已跳过")
    return result


def _clean_units(raw: Any, notes: List[str]) -> Dict[str, Dict[str, Any]]:
    """量纲表归一化：条目须为 ``{canonical, factors}``；坏条目跳过，避免换算时抛 AttributeError / TypeError。"""
    result: Dict[str, Dict[str, Any]] = {}
    if not isinstance(raw, dict):
        return result
    for dimension, info in raw.items():
        if not isinstance(info, dict) or not isinstance(info.get("factors"), dict):
            notes.append(f"量纲 {dimension} 缺少 factors 对象，已跳过")
            continue
        factors: Dict[str, float] = {}
        for unit, factor in info["factors"].items():
            try:
                factors[str(unit)] = float(factor)
            except (TypeError, ValueError):
                notes.append(f"量纲 {dimension} 的换算因子 {unit} 不是数字，已跳过")
        if not factors:
            notes.append(f"量纲 {dimension} 没有可用的换算因子，已跳过")
            continue
        entry: Dict[str, Any] = {"factors": factors}
        canonical = info.get("canonical")
        if isinstance(canonical, str) and canonical.strip():
            entry["canonical"] = canonical.strip()
        result[str(dimension)] = entry
    return result


def load_alias_rules(rules_path: Optional[Path], explicit: bool) -> Tuple[Dict[str, Any], str]:
    """加载属性别名与单位规则，返回 (规则字典, 来源说明)。

    宽容策略：字段缺失 / 类型不对 / 嵌套写法不规范时按可用部分继续，原因写入来源说明，
    不让一条写错的规则中止整条流水线；显式指定 ``--alias-rules`` 且完全不可用时抛异常（退出码 2）。
    """
    if rules_path is None or not Path(rules_path).is_file():
        if explicit:
            raise FileNotFoundError(f"规则文件不存在：{rules_path}")
        return {}, "未找到规则文件（未做别名合并与单位换算）"
    try:
        with open(Path(rules_path), "r", encoding="utf-8-sig") as f:
            raw = json.load(f)
        source = str(Path(rules_path))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        if explicit:
            raise ValueError(f"规则文件无法解析：{rules_path}（{exc}）")
        return {}, f"规则文件无法解析（{exc}），未做别名合并与单位换算"

    notes: List[str] = []
    if not isinstance(raw, dict):
        if explicit:
            raise ValueError(f"规则文件格式错误：顶层应为对象（{rules_path}）")
        return {}, "规则文件格式错误（顶层应为对象），未做别名合并与单位换算"

    rules: Dict[str, Any] = {
        "attribute_aliases": _clean_alias_map(raw.get("attribute_aliases"), notes),
        "unit_symbols": _clean_unit_symbols(raw.get("unit_symbols"), notes),
        "units": _clean_units(raw.get("units"), notes),
    }
    for key in ("attribute_aliases", "unit_symbols", "units"):
        if key not in raw:
            notes.append(f"缺少字段 {key}")
    if explicit and not any(rules.values()):
        raise ValueError(f"规则文件不可用：{rules_path}（{'；'.join(notes) or '未提供有效字段'}）")
    if notes:
        source = f"{source}（{'；'.join(notes)}）"
    return rules, source


def build_reverse_alias_map(rules: Dict[str, Any]) -> Dict[str, str]:
    """构建 别名→标准名 映射（标准名也映射到自身）。"""
    reverse: Dict[str, str] = {}
    for canonical, aliases in rules.get("attribute_aliases", {}).items():
        reverse[str(canonical)] = str(canonical)
        for alias in aliases:
            reverse[str(alias)] = str(canonical)
    return reverse


def build_value_regex(rules: Dict[str, Any]) -> "re.Pattern[str]":
    """构建“数值+可选单位”正则：拉丁单位串，或规则表中收录的中文单位词（长词优先）。"""
    latin = r"[A-Za-zμ°℃%‰³²/·]+(?:\([A-Za-z0-9]+\))?"
    cjk_units = [
        str(u)
        for u in rules.get("unit_symbols", {}).keys()
        if re.search(r"[\u4e00-\u9fff]", str(u))
    ]
    for info in rules.get("units", {}).values():
        for u in info.get("factors", {}):
            if re.search(r"[\u4e00-\u9fff]", str(u)):
                cjk_units.append(str(u))
    cjk_units = sorted(set(cjk_units), key=len, reverse=True)
    unit_alt = latin
    if cjk_units:
        unit_alt = latin + "|" + "|".join(re.escape(u) for u in cjk_units)
    return re.compile(
        r"^\s*(?P<num>[+-]?\d+(?:\.\d+)?)\s*(?P<unit>" + unit_alt + r")?(?P<rest>.*)$"
    )


def normalize_unit(unit_raw: str, rules: Dict[str, Any]) -> Tuple[str, Optional[str], Optional[float]]:
    """单位归一化：返回 (归一化单位符号, 量纲, 到标准单位的换算因子)；未收录则量纲为 None。"""
    unit = (unit_raw or "").strip()
    if not unit:
        return "", None, None
    unit = rules.get("unit_symbols", {}).get(unit, unit)
    for dimension, info in rules.get("units", {}).items():
        factor = info.get("factors", {}).get(unit)
        if factor is not None:
            return unit, dimension, float(factor)
    return unit, None, None


def clean_attribute_name(raw: str) -> str:
    """清理属性名：去首尾空白/冒号/表格残留，并去掉人工编号前缀（“1. 整机重量”“C 最大挖掘深度”
    →“整机重量”“最大挖掘深度”）。

    尾部括注（如“行走速度（高）”）不在这里删除：它会在 normalize_attribute 中作为工况限定词
    单独记录，避免把不同工况下的取值混为一谈。
    """
    name = (raw or "").strip()
    name = re.sub(r"^(?:[0-9]{1,2}|[A-Za-z])[、.．)）]\s*", "", name)
    name = re.sub(r"^[A-Za-z]\s+(?=[\u4e00-\u9fff])", "", name)
    # 中英混排写法（“总高Overall height”→“总高”）：去掉紧跟中文的英文尾巴，但保留“/单位”后缀
    name = re.sub(r"(?<=[\u4e00-\u9fff])\s*[A-Za-z][A-Za-z\s\-]*$", "", name)
    return name.strip(" \t：:|")


def normalize_attribute(name: str, reverse_alias: Dict[str, str]) -> Tuple[str, str, str]:
    """属性名归一化：返回 (标准属性名, 工况限定词, 基础名)。

    工况限定词单独剥离为 condition，不参与别名合并，避免把不同工况下的取值误判为冲突：
    前缀式（额定/最大/最小/常用），以及尾部括注（“行走速度（高）”→“行走速度”，工况“高”）。
    """
    condition = ""
    base = name
    m = RE_TRAILING_QUALIFIER.search(name)
    if m:
        condition = m.group("q").strip()
        stripped = name[: m.start()].strip()
        base = stripped or name
    else:
        for qualifier in CONDITION_QUALIFIERS:
            if name.startswith(qualifier) and len(name) > len(qualifier):
                condition = qualifier
                base = name[len(qualifier):]
                break
    normalized = reverse_alias.get(base, base)
    return normalized, condition, base


def plausible_attribute(name: str, reverse_alias: Dict[str, str]) -> bool:
    """属性名合理性：非空、≤20 字符、非纯数字；含中文即认可，纯拉丁词须是已知别名。

    明显不是属性名的内容一律拒绝：说明性空话（“并满足以下要求”“如下”）、图片文件名或路径残留。
    """
    if not name or len(name) > 20:
        return False
    if name in PROSE_ATTR_NAMES:
        return False
    if any(word in name for word in PROSE_ATTR_WORDS):
        return False  # 正文语句片段（“…并满足以下要求”“…如下”），不是产品参数名
    if re.search(r"\.[A-Za-z]", name) or "/" in name or "\\" in name:
        return False  # 图片文件名/URL/路径残留（如 de-error.png)…）
    if re.fullmatch(r"[0-9.\-]+", name):
        return False
    if re.search(r"[\u4e00-\u9fff]", name):
        return True
    return name.lower() in {k.lower() for k in reverse_alias}


def is_model_attribute(name: str) -> bool:
    """型号类属性名判定：型号/产品型号/…，兼容“型号 Model”“型号/Model”等中英混排写法。"""
    lowered = name.lower()
    if lowered in MODEL_ATTR_NAMES:
        return True
    for candidate in MODEL_ATTR_NAMES:
        if lowered.startswith(candidate) and len(name) > len(candidate):
            return bool(RE_MODEL_ATTR_TAIL.match(name[len(candidate):]))
    return False


def split_attribute_unit(raw_name: str) -> Tuple[str, str]:
    """拆分“…属性名/单位”写法，返回 (属性名, 单位提示)。

    真实企业规格书大量使用这种表头（回转速度/r/min、斗容/m³、行走速度/km/h）；
    只有当斜杠后是形状合法的拉丁单位串时才拆分，避免把路径残留当成属性名。
    """
    m = RE_ATTR_SLASH_UNIT.match((raw_name or "").strip())
    if not m:
        return raw_name, ""
    return m.group("name").strip(), m.group("unit").strip()


def is_identifier_attribute(name: str) -> bool:
    """标识类字段判定：物料编码/零件号/图号/序号及其变体（…编码、…编号）都不是可比对的产品参数。"""
    return name in IDENTIFIER_ATTR_NAMES or name.endswith(("编码", "编号"))


def is_plausible_model_token(token: str) -> bool:
    """型号词合理性：须含 ≥2 个字母、非“数值+单位”写法（35mm/1.5m³）、非通用词。

    真实文档里“35mm”“1.5m³”常出现在型号列错位处，若当成型号会把不同产品的参数拉到一起比较。
    """
    token = (token or "").strip()
    if not token or token.lower() in {"model", "type", "no", "spec", "ip"}:
        return False
    if len(re.findall(r"[A-Za-z]", token)) < 2:
        return False
    return not RE_UNIT_LIKE_TOKEN.match(token)


def with_variant_suffix(model: str, tail: str) -> str:
    """把型号后面的变型词并入型号（“ABCDEF-100 HEV”与“ABCDEF-100”不是同一台机器）。

    真实资料常把同一基础型号的变型写成“型号 + 变型词”（HEV/PRO/ECO/PLUS…）；
    若只取基础型号，会把不同变型机的参数拉到一起比较。
    """
    m = RE_MODEL_VARIANT.match(tail or "")
    return f"{model} {m.group(1).upper()}" if m else model


def find_model_token(text: str) -> Optional[str]:
    """在文本中查找型号词：字母数字混合且含数字，并要求至少含两个字母。

    纯数字（年份 2024、编码 803013104）与“数值+单位”（35mm）不认作型号；复合串按下划线/
    斜杠/连续横线切分后取第一个合法片段（ABCDEF-100_3104-----_ABCDEF-100 → ABCDEF-100）；紧跟的
    变型词并入型号（ABCDEF-100 HEV → ABCDEF-100 HEV）。
    """
    source = text or ""
    for m in RE_MODEL_TOKEN.finditer(source):
        parts = [p.strip("-") for p in RE_MODEL_SPLIT.split(m.group(0))]
        for idx, part in enumerate(parts):
            if not is_plausible_model_token(part):
                continue
            if idx == len(parts) - 1:
                return with_variant_suffix(part, source[m.end(): m.end() + 14])
            return part
    return None


def is_non_attribute_table(header_cells: List[str]) -> bool:
    """是否为“非参数表”：版本/日期/更新记录、目录、名称说明等表，整表跳过。

    真实规格书常把版本履历表、变更记录表与参数表混排，这类表没有可比对的量化参数；
    仅在表格没有独立“单位”列时生效，含“单位”列的表照常抽取。
    """
    return any(any(word in cell for word in NON_ATTRIBUTE_HEADER_WORDS) for cell in header_cells)


# 取值列表头词：判断“单位”列在取值列之前还是之后（真实资料两种写法都存在）
VALUE_HEADER_WORDS = ("参数", "数值", "值", "规格", "数据")


def is_value_header(cell: str) -> bool:
    """表头单元格是否为“取值/数值”列（用于判断“单位”列相对取值列的位置）。"""
    return any(word in cell for word in VALUE_HEADER_WORDS)


def find_value_column_groups(header_cells: List[str]) -> List[Tuple[int, int, Optional[int]]]:
    """找出表中所有“属性-取值(-单位)”列组，返回 (属性列, 取值列, 单位列或 None)。

    - 两列表（| 属性 | 取值 |）：一组 (0, 1, None)；
    - 含独立“单位”列的表：按表头语义判断单位列在取值之后还是之前——
      `| 项目 | 参数 | 单位 |`（单位在取值之后），或 `| 项目内容 | 单位 | 参数 |`（单位在取值
      之前，真实折页/单页资料常见）；左右并排的对比表可同时命中多组；
    - 其他表（无“单位”列的多列表、故障排查表、备件清单）返回空列表，由调用方整表跳过。
    """
    if len(header_cells) == 2:
        return [(0, 1, None)]
    groups: List[Tuple[int, int, Optional[int]]] = []
    for idx, cell in enumerate(header_cells):
        if idx < 1 or "单位" not in cell:
            continue
        if idx + 1 < len(header_cells) and is_value_header(header_cells[idx + 1]):
            groups.append((idx - 1, idx + 1, idx))  # 单位列在取值列之前
        elif idx >= 2:
            groups.append((idx - 2, idx - 1, idx))  # 常规：属性 | 取值 | 单位
    return groups


def extract_table_unit_hint(header_cells: List[str]) -> str:
    """从表头提取整表单位提示（如“单位mm”“单位：/km/h”），供排版型表格的相邻标签-数值对使用。"""
    for cell in header_cells:
        if "单位" not in cell:
            continue
        m = re.search(
            r"单位\s*[:：]?\s*(?P<unit>[A-Za-zμ°℃%‰³²·][A-Za-zμ°℃%‰³²·/\-]{0,11}|[\u4e00-\u9fff]{1,4})",
            cell,
        )
        if m:
            return m.group("unit").strip()
    return ""


def iter_adjacent_value_pairs(cells: List[str]) -> List[Tuple[str, str]]:
    """排版型表格的“相邻标签-数值”对：先压掉空单元格，再取相邻两格（标签, 数值）。

    真实折页/单页对比表把多个参数并排放在一行里（| A 总长 | 7825 | 操作重量/kg | 14600 |），
    无法用“属性-取值-单位”列组描述，只能按“标签紧跟数值”的客观形态抽取。
    """
    compact = [c.strip() for c in cells if c.strip()]
    return [(compact[idx], compact[idx + 1]) for idx in range(len(compact) - 1)]


def parse_value_text(value_text: str, value_regex: "re.Pattern[str]") -> Optional[Dict[str, Any]]:
    """解析取值文本：数值+可选单位；取值不完整（日期、范围、多维尺寸等）时返回 None。"""
    m = value_regex.match(value_text)
    if not m or not m.group("num"):
        return None
    rest = m.group("rest") or ""
    if rest:
        if rest[0] not in _VALUE_TERMINATORS:
            return None
        if rest.strip().startswith(("×", "x", "X", "~", "～")):
            return None  # 范围或多维尺寸：只取局部会失真，放弃该条
    num = float(m.group("num"))
    if num.is_integer():
        num = int(num)
    unit = (m.group("unit") or "").strip()
    end = m.end("unit") if unit else m.end("num")
    return {
        "value_number": num,
        "unit_raw": unit,
        "value_raw": value_text[:end].strip(),
    }


def compute_fenced_lines(lines: List[str]) -> Set[int]:
    """返回代码块内的行下标集合（0 起始，含围栏行本身），判定逻辑与 audit_markdown 保持一致。"""
    fenced: Set[int] = set()
    fence_open = False
    fence_char = ""
    fence_len = 0
    for idx, line in enumerate(lines):
        m = RE_FENCE.match(line)
        opens = bool(m) and not (m.group(1)[0] == "`" and "`" in m.group(2))
        if fence_open:
            fenced.add(idx)
            if (
                opens
                and m.group(1)[0] == fence_char
                and len(m.group(1)) >= fence_len
                and not m.group(2).strip()
            ):
                fence_open = False
            continue
        if opens:
            fenced.add(idx)
            fence_open = True
            fence_char = m.group(1)[0]
            fence_len = len(m.group(1))
    return fenced


EVIDENCE_LIMIT = 160  # 证据原文的展示上限（字符数）


def _evidence_text(evidence_line: str, record: Dict[str, Any]) -> str:
    """证据原文：窄行原样保留；宽行以“属性名 / 取值”出现处为中心截取。

    宽表格行（如左右并排多组的规格书表）里靠后的“属性-取值”若只截行首，
    证据里就看不到被抽取的内容；以命中处为中心截取可保证关键内容一定在证据里。
    """
    line = evidence_line or ""
    if len(" ".join(line.split())) <= EVIDENCE_LIMIT:
        return truncate_text(line, EVIDENCE_LIMIT)
    for needle in (record.get("attribute_raw"), record.get("value_raw")):
        text = str(needle or "").strip()
        if not text:
            continue
        pos = line.find(text)
        if pos >= 0:
            return truncate_text(window_text(line, pos, pos + len(text), radius=50), EVIDENCE_LIMIT)
    return truncate_text(line, EVIDENCE_LIMIT)


def _append_record(
    entry: Dict[str, Any], record: Dict[str, Any], line_no: int, context: str, evidence_line: str
) -> None:
    record["line_number"] = line_no
    record["context"] = context
    record["evidence"] = _evidence_text(evidence_line, record)
    entry["attributes"].append(record)


def _handle_candidate(
    entry: Dict[str, Any],
    attr_raw: str,
    value_text: str,
    line_no: int,
    context: str,
    evidence_line: str,
    rules: Dict[str, Any],
    value_regex: "re.Pattern[str]",
    reverse_alias: Dict[str, str],
    table_unit_hint: str = "",
) -> None:
    """处理一个候选“属性=取值”：识别型号类属性或解析数值+单位，成功则记录。

    table_unit_hint 为整表单位提示（表头里的“单位mm”），仅在属性名自身不带单位时兜底使用。
    """
    name = clean_attribute_name(attr_raw)
    name, unit_hint = split_attribute_unit(name)
    if not unit_hint:
        unit_hint = table_unit_hint
    if not name or is_identifier_attribute(name):
        return
    value_text = (value_text or "").strip()
    # 型号类属性：取原始文本中的型号词，不做数值解析
    if is_model_attribute(RE_TRAILING_QUALIFIER.sub("", name).strip() or name):
        model_value = find_model_token(value_text)
        if not model_value:
            return  # 取值中不含型号词（多为表头错位或编号），不记作型号，也不产生记录
        if entry["model"] is None:
            entry["model"] = model_value
            entry["model_source"] = f"型号属性（第 {line_no} 行）"
        _append_record(
            entry,
            {
                "kind": "model",
                "attribute_raw": name,
                "attribute_normalized": "型号",
                "condition": "",
                "value_raw": value_text[:40],
                "model_value": model_value,
            },
            line_no,
            context,
            evidence_line,
        )
        return
    if not plausible_attribute(name, reverse_alias):
        return
    parsed = parse_value_text(value_text, value_regex)
    if parsed is not None and not parsed["unit_raw"] and unit_hint:
        # “属性名/单位”表头：取值列只写数字，单位来自表头（回转速度/r/min → 12 r/min）
        parsed = parse_value_text(f"{value_text} {unit_hint}", value_regex)
    if parsed is None:
        return
    normalized, condition, base = normalize_attribute(name, reverse_alias)
    unit_norm, dimension, factor = normalize_unit(parsed["unit_raw"], rules)
    canonical_unit = ""
    canonical_value = None
    if dimension is not None and factor is not None:
        canonical_unit = str(rules.get("units", {}).get(dimension, {}).get("canonical", ""))
        canonical_value = round(parsed["value_number"] * factor, 10)
        if canonical_value == int(canonical_value):
            canonical_value = int(canonical_value)
    _append_record(
        entry,
        {
            "kind": "numeric",
            "attribute_raw": name,
            "attribute_base": base,
            "attribute_normalized": normalized,
            "condition": condition,
            "value_raw": parsed["value_raw"],
            "value_number": parsed["value_number"],
            "unit_raw": parsed["unit_raw"],
            "unit_normalized": unit_norm,
            "dimension": dimension,
            "canonical_value": canonical_value,
            "canonical_unit": canonical_unit,
        },
        line_no,
        context,
        evidence_line,
    )


def extract_from_file(
    path: Path,
    rel_path: str,
    rules: Dict[str, Any],
    value_regex: "re.Pattern[str]",
    reverse_alias: Dict[str, str],
) -> Dict[str, Any]:
    """抽取单个文件：表格 → 行内/列表 → 一级标题/文件名型号回退。无法解码的文件如实记录并跳过。"""
    entry: Dict[str, Any] = {
        "file_path": rel_path,
        "absolute_path": str(path),
        "model": None,
        "model_source": None,
        "decode_error": None,
        "attributes": [],
    }
    try:
        raw = path.read_bytes()
    except OSError as exc:
        entry["decode_error"] = f"无法读取文件：{exc}"
        return entry
    try:
        text = raw.decode("utf-8-sig")
    except UnicodeDecodeError as exc:
        entry["decode_error"] = f"文件无法按 UTF-8 解码（{exc}），已跳过内容抽取。"
        return entry
    lines = text.splitlines()
    fenced = compute_fenced_lines(lines)
    table_lines: Set[int] = set()

    # 1) 表格：仅抽取“属性-取值”表（两列，或从第 3 列起含独立“单位”列）
    for header_no, header_cells, rows in iter_table_blocks(lines):
        if (header_no - 1) in fenced:
            continue  # 代码块内的“表格”是示例内容，不抽取
        table_lines.add(header_no)
        table_lines.add(header_no + 1)
        for row_no, _cells in rows:
            table_lines.add(row_no)
        groups = [
            g
            for g in find_value_column_groups(header_cells)
            if not is_non_attribute_table([header_cells[g[0]]])
        ]
        table_unit_hint = extract_table_unit_hint(header_cells)
        if not groups and not table_unit_hint:
            continue  # 非“属性-取值”结构（多列对照表、故障排查表、工具清单等）整表跳过
        for row_no, cells in rows:
            if groups:
                for attr_i, value_i, unit_i in groups:
                    if value_i >= len(cells):
                        continue
                    value_text = cells[value_i]
                    if unit_i is not None and unit_i < len(cells) and cells[unit_i].strip():
                        value_text = (cells[value_i].strip() + " " + cells[unit_i].strip()).strip()
                    _handle_candidate(
                        entry, cells[attr_i] if attr_i < len(cells) else "", value_text, row_no,
                        "表格", lines[row_no - 1], rules, value_regex, reverse_alias,
                    )
                continue
            # 排版型表格（折页/单页对比表）：同行相邻的“标签-数值”对
            for attr_text, pair_value in iter_adjacent_value_pairs(cells):
                if is_non_attribute_table([attr_text]):
                    continue  # 版本/日期/序号等非参数标签（排版型表格里也常出现）
                if not pair_value[:1].isdigit() and pair_value[:1] not in "+-":
                    continue  # 只认“数值紧跟标签”的客观形态，避免把标签-标签当成参数
                _handle_candidate(
                    entry, attr_text, pair_value, row_no, "表格", lines[row_no - 1],
                    rules, value_regex, reverse_alias, table_unit_hint,
                )

    # 2) 行内与列表：“属性：数值+单位”
    for idx, line in enumerate(lines):
        line_no = idx + 1
        if idx in fenced:
            continue
        if line_no in table_lines:
            continue
        li = RE_LIST_ITEM.match(line)
        body = li.group("body") if li else line
        context = "列表" if li else "正文"
        for m in RE_MODEL_COLON.finditer(body):
            _handle_candidate(
                entry, m.group("attr"), m.group("value"), line_no, context, line,
                rules, value_regex, reverse_alias,
            )
        for m in RE_ATTR_COLON.finditer(body):
            _handle_candidate(
                entry, m.group("attr"), m.group("value"), line_no, context, line,
                rules, value_regex, reverse_alias,
            )

    # 3) 型号回退：依次取一级标题、文件名中的型号词
    if entry["model"] is None:
        for idx, line in enumerate(lines):
            if idx in fenced:
                continue
            hm = RE_HEADING1.match(line)
            if hm:
                token = find_model_token(hm.group("title"))
                if token:
                    entry["model"] = token
                    entry["model_source"] = f"一级标题（第 {idx + 1} 行）"
                break
    if entry["model"] is None:
        token = find_model_token(path.stem)
        if token:
            entry["model"] = token
            entry["model_source"] = "文件名"
    entry["attributes"] = _dedupe_records(entry["attributes"])
    return entry


def _dedupe_records(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """同一行重复命中的同一条记录只保留一次（如同一行被多个模式重复捕获）。

    去重键含行号/属性名/工况/取值：不同行的同名属性仍各自保留。
    """
    seen: Set[Tuple[Any, ...]] = set()
    unique: List[Dict[str, Any]] = []
    for rec in records:
        key = (
            rec.get("line_number"),
            rec.get("kind"),
            rec.get("attribute_raw"),
            rec.get("condition"),
            rec.get("value_raw"),
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append(rec)
    return unique


def build_extraction(
    input_path: Path,
    rules_path: Optional[Path] = None,
    output_dir: Optional[Path] = None,
    report_base_name: str = "",
    show_progress: bool = True,
) -> Dict[str, Any]:
    """遍历输入（单个 `.md` 或目录）抽取属性记录，返回结构化结果（内存对象，供直接比对）。

    - rules_path 为 None 时使用项目内 rules/attribute_aliases.json；显式传入但不可用则抛异常；
    - output_dir / report_base_name 用于扫描时排除输出目录与报告文件本身，避免把报告当输入重复抽取；
    - show_progress=False 关闭进度条（非交互终端会自动降级为分段整行日志）。
    """
    explicit = rules_path is not None
    if rules_path is None:
        default_rules = SCRIPT_DIR.parent / "rules" / "attribute_aliases.json"
        rules_path = default_rules if default_rules.is_file() else None
    rules, rules_source = load_alias_rules(rules_path, explicit=explicit)
    reverse_alias = build_reverse_alias_map(rules)
    value_regex = build_value_regex(rules)
    files = collect_files(input_path, output_dir or input_path, report_base_name)
    bar = Progress("属性抽取", len(files), enabled=None if show_progress else False)
    entries: List[Dict[str, Any]] = []
    for index, fp in enumerate(files, 1):
        entry = extract_from_file(fp, relative_path(fp, input_path), rules, value_regex, reverse_alias)
        entries.append(entry)
        bar.update(index, detail=str(entry.get("file_path") or fp.name))
    bar.finish(extra=f"抽取 {len(files):,} 个文件的属性记录")
    return {
        "input": str(input_path.resolve()),
        "rules_source": rules_source,
        "files": entries,
        "summary": {
            "files_scanned": len(entries),
            "files_with_model": sum(1 for e in entries if e["model"]),
            "files_with_decode_error": sum(1 for e in entries if e["decode_error"]),
            "numeric_records": sum(1 for e in entries for a in e["attributes"] if a["kind"] == "numeric"),
            "model_records": sum(1 for e in entries for a in e["attributes"] if a["kind"] == "model"),
        },
    }


if __name__ == "__main__":
    print(
        "attribute_extract 是库模块（供 detect_conflicts.py 调用）；"
        "请运行 python scripts/detect_conflicts.py 或 scripts/run_all.py。"
    )

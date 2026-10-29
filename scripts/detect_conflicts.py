#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""detect_conflicts.py — 同一型号同类属性冲突比对（规则式抽取，非语义推断）。

读取 Markdown（目录或单个文件），内部完成属性抽取后按（型号, 标准属性名, 工况）分组比对；
只列出取值不一致且证据来自 ≥2 个不同文件的分组，并给出各来源文件、行号、取值与证据原文。
产出两份交付文件（HTML 报告 + XLSX 逐行清单），不修改输入、不生成 MD/JSON。
退出码：0=完成；1=输入/输出错误；2=规则文件错误。
"""

from __future__ import annotations

import argparse
import math
import sys
from collections import defaultdict
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import (  # noqa: E402
    TOOL_VERSION,
    build_html,
    details,
    esc,
    kpi_cards,
    make_timestamped_dir,
    now_local_str,
    now_utc_iso,
    print_error,
    safe_print,
    section,
    table,
    write_text_file,
    write_xlsx_file,
)
from attribute_extract import build_extraction  # noqa: E402
from audit_markdown import SUPPORTED_EXTENSIONS  # noqa: E402

TOOL_NAME = "detect_conflicts"
DEFAULT_REPORT_BASE_NAME = "知识库同类属性冲突审查报告"
REL_TOL = 1e-9
ABS_TOL = 1e-12


def _values_equal(a: Any, b: Any) -> bool:
    return math.isclose(float(a), float(b), rel_tol=REL_TOL, abs_tol=ABS_TOL)


def compare_values(members: List[Tuple[Dict[str, Any], Dict[str, Any]]]) -> Tuple[bool, str]:
    """比较同组内多条取值，返回 (是否不一致, 比较口径说明)。members 至少 2 条。"""
    dims = {rec.get("dimension") for _, rec in members}
    if None not in dims and len(dims) == 1:
        canonical_unit = members[0][1].get("canonical_unit") or ""
        values = [rec.get("canonical_value") for _, rec in members]
        if all(v is not None for v in values):
            same = all(_values_equal(values[0], v) for v in values[1:])
            return (not same), f"各方单位均收录于规则表，已换算为 {canonical_unit} 后比较。"
    units = {rec.get("unit_normalized") or "" for _, rec in members}
    if len(units) == 1:
        unit = next(iter(units))
        values = [rec.get("value_number") for _, rec in members]
        if all(v is not None for v in values):
            same = all(_values_equal(values[0], v) for v in values[1:])
            if unit:
                basis = f"单位 {unit} 未收录于规则表，按相同单位的数值字面比较。"
            else:
                basis = "各方均未标注单位，按数值字面比较（若原文实际单位不同，此口径可能不成立）。"
            return (not same), basis
    values_text = [str(rec.get("value_raw") or "").strip() for _, rec in members]
    basis = "各方单位不一致或缺失、无法换算，按原始取值字面比较。"
    return (len(set(values_text)) > 1), basis


def _dedupe_members(
    members: List[Tuple[Dict[str, Any], Dict[str, Any]]],
) -> List[Tuple[Dict[str, Any], Dict[str, Any]]]:
    """去掉同一 (文件, 行号, 取值) 的重复记录：同一处证据重复出现不构成对照。"""
    seen = set()
    unique: List[Tuple[Dict[str, Any], Dict[str, Any]]] = []
    for fentry, rec in members:
        key = (
            str(fentry.get("file_path") or ""),
            rec.get("line_number"),
            str(rec.get("value_raw") or ""),
            rec.get("canonical_value"),
        )
        if key in seen:
            continue
        seen.add(key)
        unique.append((fentry, rec))
    return unique


def build_conflicts(
    data: Dict[str, Any],
) -> Tuple[List[Dict[str, Any]], Dict[str, int]]:
    """按（型号, 标准属性名, 工况）分组比对，返回 (不一致分组列表, 统计)。

    列为冲突的门槛：同一「型号+属性名+工况」下取值不一致，且证据来自 **≥2 个不同文件**。
    同一文件内重复出现的记录（同一处证据、或同文档不同章节复述）只算单一来源，不报冲突，
    避免把真实文档里的重复表述、表格重复渲染误报为跨文档冲突。
    """
    group_map: Dict[Tuple[str, str, str], List[Tuple[Dict[str, Any], Dict[str, Any]]]] = defaultdict(list)
    single_source = 0
    no_model = 0
    for fentry in data["files"]:
        model = str(fentry.get("model") or "").strip()
        for rec in fentry.get("attributes", []):
            if rec.get("kind") != "numeric":
                continue
            if not model:
                no_model += 1
                continue
            # 严格同一性：型号按原文精确匹配（只去首尾空白），型号不同即视为不同产品
            key = (
                model,
                str(rec.get("attribute_normalized") or ""),
                str(rec.get("condition") or ""),
            )
            group_map[key].append((fentry, rec))

    conflicts: List[Dict[str, Any]] = []
    groups_compared = 0
    groups_consistent = 0
    groups_single_file = 0
    records_compared = 0
    records_single_file = 0
    for key, members in group_map.items():
        members = _dedupe_members(members)
        if len(members) < 2:
            single_source += len(members)
            continue
        if len({str(f.get("file_path") or "") for f, _r in members}) < 2:
            # 同一文件内重复出现：只有一处来源，不构成“跨来源取值不一致”
            groups_single_file += 1
            records_single_file += len(members)
            continue
        groups_compared += 1
        records_compared += len(members)
        is_conflict, basis = compare_values(members)
        if not is_conflict:
            groups_consistent += 1
            continue
        conflicts.append(
            {
                "model": key[0],
                "attribute": key[1],
                "condition": key[2],
                "comparison_basis": basis,
                "members": [_serialize_member(f, r) for f, r in members],
            }
        )
    conflicts.sort(key=lambda g: (g["model"], g["attribute"], g["condition"]))
    for i, g in enumerate(conflicts, 1):
        g["group_id"] = f"AC-{i:04d}"
    summary = {
        "groups_compared": groups_compared,
        "groups_conflict": len(conflicts),
        "groups_consistent": groups_consistent,
        "groups_single_file": groups_single_file,
        "records_compared": records_compared,
        "records_single_source": single_source,
        "records_single_file": records_single_file,
        "records_no_model": no_model,
    }
    return conflicts, summary


def _serialize_member(fentry: Dict[str, Any], rec: Dict[str, Any]) -> Dict[str, Any]:
    """把 (文件条目, 属性记录) 压缩为报告用的来源记录（保留定位与证据）。"""
    return {
        "file": fentry.get("file_path", ""),
        "line_number": rec.get("line_number"),
        "context": rec.get("context", ""),
        "attribute_raw": rec.get("attribute_raw", ""),
        "value_raw": rec.get("value_raw", ""),
        "value_number": rec.get("value_number"),
        "unit_normalized": rec.get("unit_normalized", ""),
        "canonical_value": rec.get("canonical_value"),
        "canonical_unit": rec.get("canonical_unit", ""),
        "evidence": rec.get("evidence", ""),
    }


def build_report_data(input_path: Path, extraction: Dict[str, Any]) -> Dict[str, Any]:
    """汇总比对结果为报告数据（HTML 报告与清单表格共用）。"""
    conflicts, summary = build_conflicts(extraction)
    summary["files_scanned"] = extraction["summary"]["files_scanned"]
    summary["numeric_records_total"] = extraction["summary"]["numeric_records"]
    return {
        "tool": TOOL_NAME,
        "version": TOOL_VERSION,
        "generated_at_utc": now_utc_iso(),
        "generated_at_local": now_local_str(),
        "input_path": str(input_path.resolve()),
        "rules_source": extraction["rules_source"],
        "summary": summary,
        "conflicts": conflicts,
        "scope_notes": [
            "只列出「相同属性但描述不一致」的分组；描述一致的分组不列入报告（数量见“比对概况”）。",
            "同一性要求：只有「型号 + 属性名 + 工况」三者完全一致的记录才会相互比较；型号按原文精确匹配，"
            "不做模糊归类——“XCT50G5-1(增强版)” 与 “XCT50G5-1v” 属不同型号，不比较、也不会列入报告。",
            "未识别型号的记录，以及只有单一来源的记录不参与比较，数量见“比对概况”。",
            "列入门槛：同一「型号+属性名+工况」下取值不一致（即「相同属性但描述不一致」），且证据必须来自 **≥2 个不同文件**；"
            "同一文件内重复出现的记录（同一处证据、同文档不同章节复述）只算单一来源，不列入。",
            "抽取为规则式（非语义理解）：只识别“属性-取值”表（两列，或含独立“单位”列），以及列表/正文的"
            "“属性：数值+单位”模式；代码块、多列对照表、故障排查表、清单表不抽取。",
            "本工具只读取 Markdown 文本，不读取、不对照原始文档（DOCX/PDF/图片）。",
        ],
    }


# 冲突清单表格（XLSX）表头：一行 = 一处来源取值，同一组冲突的各来源共用同一「冲突编号」，
# 按「冲突编号」即可筛出同一处冲突涉及哪几个文档。
SHEET_NAME = "冲突清单"
SHEET_HEADER: Tuple[str, ...] = (
    "冲突编号",
    "型号",
    "属性",
    "描述",
    "来源文档",
    "行号",
    "证据原文",
)
SHEET_WIDTHS: Tuple[int, ...] = (12, 18, 22, 30, 34, 8, 46)
SHEET_NUMERIC_COLS: Tuple[int, ...] = (5,)  # 行号按数字右对齐


def _describe_value(member: Dict[str, Any]) -> str:
    """把「原始取值 + 换算后取值」合并为一句描述；未发生单位换算时只给原始取值。"""
    raw = str(member.get("value_raw") or "").strip()
    canonical = member.get("canonical_value")
    if canonical is None:
        return raw
    canon = f"{canonical} {member.get('canonical_unit') or ''}".strip()
    if not raw or raw == canon:
        return canon
    return f"{raw}（换算后 {canon}）"


def build_sheet_rows(data: Dict[str, Any]) -> List[Tuple[Any, ...]]:
    """把「相同属性但描述不一致」的明细展开为清单表格的数据行（一行 = 一处来源取值）。

    同一组「相同属性但描述不一致」的所有来源共用同一个「冲突编号」；组内来源按「文档、行号」排序，
    便于人工逐份打开文档核对。
    """
    rows: List[Tuple[Any, ...]] = []
    for group in data["conflicts"]:
        members = sorted(
            group["members"], key=lambda m: (str(m.get("file") or ""), m.get("line_number") or 0)
        )
        for member in members:
            rows.append(
                (
                    group["group_id"],
                    group["model"],
                    group["attribute"],
                    _describe_value(member),
                    member.get("file") or "",
                    member.get("line_number") if member.get("line_number") else "",
                    member.get("evidence") or "",
                )
            )
    return rows


def write_conflict_sheet(path: Any, data: Dict[str, Any]):
    """把冲突清单写为 XLSX 交付文件（表头配色、列宽、冻结首行、自动筛选），返回 Path。"""
    return write_xlsx_file(
        path,
        SHEET_NAME,
        SHEET_HEADER,
        build_sheet_rows(data),
        widths=SHEET_WIDTHS,
        numeric_cols=SHEET_NUMERIC_COLS,
    )


def render_html_report(data: Dict[str, Any]) -> str:
    """把比对结果渲染为单文件 HTML（样式内联、无外部资源，可直接打开或转发）。

    重点呈现“哪几个文档的什么属性不一致”：每组「相同属性但描述不一致」一条可展开明细，组内逐条给出
    来源文件、行号、原始取值、换算后取值与证据原文；逐行清单另见同名 XLSX 表格文件。
    """
    s = data["summary"]
    conflicts = data["conflicts"]
    uncompared = (
        s.get("records_single_source", 0) + s.get("records_single_file", 0) + s.get("records_no_model", 0)
    )
    kpis = kpi_cards(
        [
            ("扫描文件数", s["files_scanned"], ""),
            ("含相同属性数量", s["groups_compared"], "hot"),
            ("相同属性但描述不一致", s["groups_conflict"], "warn" if s["groups_conflict"] else "ok"),
        ]
    )
    overview_rows = [
        ("输入路径", data["input_path"]),
        ("属性别名/单位规则", data["rules_source"]),
        ("比对时间", data["generated_at_local"]),
    ]
    overview_note = (
        f"数值属性记录 {s.get('numeric_records_total', 0)} 条，其中参与比对 {s['records_compared']} 条；"
        f"含相同属性的分组 {s['groups_compared']} 个，其中描述一致 {s['groups_consistent']} 个（未列出）；"
        f"未参与比对的记录 {uncompared} 条"
        f"（单一来源 {s.get('records_single_source', 0)} 条、同一文件内重复 {s.get('records_single_file', 0)} 条、"
        f"未识别型号 {s.get('records_no_model', 0)} 条）。"
    )
    sections = [section("一、比对概况", kpis + table(["项目", "说明"], overview_rows), overview_note)]

    detail_blocks: List[str] = []
    for group in conflicts:
        rows = []
        for member in group["members"]:
            rows.append(
                [
                    member.get("file", ""),
                    member.get("line_number", ""),
                    _describe_value(member),
                    member.get("evidence", ""),
                ]
            )
        head = (
            f'<span class="badge">{esc(group["group_id"])}</span> 型号 <b>{esc(group["model"])}</b> · '
            f'<b>{esc(group["attribute"])}</b>'
            f'<span class="badge mute">{len(group["members"])} 条来源</span>'
        )
        body = (
            f'<div class="note">比较口径：{esc(group["comparison_basis"])}</div>'
            + table(
                ["来源文档", "行号", "描述", "证据原文"],
                rows,
                align=["", "num", "", ""],
            )
        )
        detail_blocks.append(details(head, body, open_=len(conflicts) <= 5))
    if not conflicts:
        detail_blocks.append('<div class="note">未发现同一型号下相同属性但描述不一致的情况。</div>')
    sections.append(
        section(
            "二、相同属性但描述不一致明细",
            "".join(detail_blocks),
            f"共 {len(conflicts)} 组「相同属性但描述不一致」（按「型号 + 属性」列出各来源文档与取值）；"
            "明细默认折叠，点击标题展开。逐行清单见同名 XLSX 表格文件。",
        )
    )

    notes = list(data.get("scope_notes", [])) + [
        "同名清单表格（.xlsx）逐行列出每组冲突的每个来源：冲突编号、型号、属性、描述、"
        "来源文档、行号、证据原文；按「冲突编号」即可筛出同一处冲突涉及哪几个文档。",
    ]
    notes_html = '<ul class="plain">' + "".join(f"<li>{esc(n)}</li>" for n in notes) + "</ul>"
    sections.append(section("三、范围与限制说明", notes_html, None))
    return build_html(
        "知识库同类属性冲突审查报告",
        "只列出同一型号下「相同属性但描述不一致」的分组，并指明各来源文档与行号；审查为只读操作，未修改任何输入文档。",
        [("工具", f"{data['tool']} v{data['version']}"), ("生成时间", data["generated_at_local"])],
        sections,
        "由 knowledge-audit-skill 自动生成（scripts/detect_conflicts.py）；结果供人工复核参考。",
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="detect_conflicts.py",
        description="同一型号同类属性冲突检测（第二阶段）：直接读取知识库 Markdown，内部完成属性抽取后比对",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/detect_conflicts.py input/\n"
            "  python scripts/detect_conflicts.py input/ --output output\n"
            "不指定 --output 时写入 output/<时间戳>/；每次产出 HTML 报告 + XLSX 清单表格，不生成 Markdown / JSON。\n"
            "退出码：0=完成；1=输入/输出错误；2=规则文件错误。"
        ),
    )
    parser.add_argument("input", help="待比对的 Markdown 文件或目录（递归扫描）")
    parser.add_argument(
        "--output",
        default=None,
        help="报告输出目录；不指定时在 output/ 下自动新建带时间戳的子目录（YYYYMMDD-HHMMSS）",
    )
    parser.add_argument(
        "--alias-rules",
        default=None,
        help="属性别名与单位规则 JSON（默认：项目内 rules/attribute_aliases.json）",
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
    safe_print(f"[detect_conflicts] 输出目录：{output_dir.resolve()}")

    rules_path = Path(args.alias_rules).expanduser() if args.alias_rules else None
    try:
        extraction = build_extraction(
            input_path,
            rules_path,
            output_dir=output_dir,
            report_base_name=DEFAULT_REPORT_BASE_NAME,
            show_progress=not args.no_progress,
        )
    except (OSError, ValueError) as exc:
        safe_print(f"[错误] {exc}")
        return 2
    for entry in extraction["files"]:
        if entry["decode_error"]:
            safe_print(f"[detect_conflicts] [提示] {entry['file_path']}：{entry['decode_error']}")

    data = build_report_data(input_path, extraction)
    s = data["summary"]
    # 写盘兜底：目录只读、磁盘写满、交付文件被 Excel/WPS 占用等都会抛 OSError，这里统一转成清晰错误而不是崩栈
    try:
        html_path = write_text_file(output_dir / f"{DEFAULT_REPORT_BASE_NAME}.html", render_html_report(data))
        safe_print(f"[detect_conflicts] HTML 报告：{html_path}")
        sheet_path = write_conflict_sheet(output_dir / f"{DEFAULT_REPORT_BASE_NAME}.xlsx", data)
    except OSError as exc:
        return print_error(f"无法写入交付文件（目录 {output_dir}）：{exc}")
    safe_print(f"[detect_conflicts] XLSX 清单表格：{sheet_path}（{s['groups_conflict']} 组相同属性但描述不一致）")
    safe_print(
        f"[detect_conflicts] 比对完成：扫描 {s['files_scanned']} 个文件，抽取数值属性记录 "
        f"{s['numeric_records_total']} 条，含相同属性的分组 {s['groups_compared']} 个，"
        f"其中相同属性但描述不一致 {s['groups_conflict']} 个（已列出）。"
    )
    return 0


def main() -> int:
    try:
        return run()
    except KeyboardInterrupt:
        safe_print("[错误] 用户中断操作。")
        return 130


if __name__ == "__main__":
    sys.exit(main())

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run_all.py — 一条命令跑完整条审查流水线（全自动、无需人工介入）。

输入（二选一）：
- 已转写为 Markdown 的知识库文件夹或单个 `.md` 文档；或
- 一个在线知识库地址（`--url` + `--token`）：先自动抓取知识库中已转写好的 Markdown，再进入审查流程。

依次执行（任一步骤失败即停止，已生成的报告不受影响）：
0. fetch_knowledge.py    在线知识库抓取（仅 `--url` 时执行）→ 逐篇 Markdown（输入缓存，不是交付文件）
1. audit_markdown.py     文档解析与转写质量审查 → HTML 报告 + XLSX 问题清单
2. detect_conflicts.py   同一型号同类属性冲突比对 → HTML 报告 + XLSX 冲突清单

每次运行都在输出根目录下新建带完整时间戳的子目录（`YYYYMMDD-HHMMSS`，同秒重复追加 `_2`、`_3`），
4 个交付文件（2 个 HTML + 2 个 XLSX）全部放在该目录内；在线抓取的 Markdown 写入
`output/_kb_source/<时间戳>/`（输入缓存，不计入交付文件）。

用法：
    python scripts/run_all.py input/
    python scripts/run_all.py "D:/知识库" --output output
    python scripts/run_all.py --url "https://kb.example.com/knowledge/asset" --token <x-access-token>
    python scripts/run_all.py --url <url> --token <token> --limit 0 --all-assets   # 全量抓取后审查
    python scripts/run_all.py input/ --fail-on-issue       # 发现问题时退出码 3，便于 CI 卡口

退出码：0=完成；1=输入/输出错误；2=规则或参数错误；3=完成但发现问题（仅 --fail-on-issue 时）。
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import List, Optional, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_markdown  # noqa: E402
import detect_conflicts  # noqa: E402
import fetch_knowledge  # noqa: E402
from generate_report import make_timestamped_dir, safe_print  # noqa: E402


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="run_all.py",
        description="一条命令跑完整条审查流水线：在线知识库抓取（可选）→ 转写质量审查 → 同一型号属性冲突比对",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/run_all.py input/\n"
            "  python scripts/run_all.py input/ --output output --fail-on-issue\n"
            "  python scripts/run_all.py --url \"https://kb.example.com/knowledge/asset\" --token <x-access-token>\n"
            "  python scripts/run_all.py --url <url> --token <token> --api-base \"https://kb.example.com/api/v1\"\n"
            "  python scripts/run_all.py --url <url> --token <token> --limit 0 --all-assets\n"
            "\n"
            "每次运行在 --output 下新建带时间戳的子目录，2 个 HTML 报告 + 2 个 XLSX 清单表格写入其中；\n"
            "用 --url 时先在线抓取已转写好的 Markdown（默认只抓有效文档、同一文档只抓第一份；\n"
            "--all-assets 抓取列表中的全部资产），抓取结果写入 <输出根目录>/_kb_source/<时间戳>/。\n"
            "退出码：0=完成；1=输入/输出错误；2=规则或参数错误；3=完成但发现问题（仅 --fail-on-issue）。"
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=None,
        help="待审查的 Markdown 文件或目录（与 --url 二选一）",
    )
    parser.add_argument(
        "--url",
        default=None,
        help="在线知识库列表页地址（如 https://kb.example.com/knowledge/asset）；给出后先抓取 Markdown 再审查",
    )
    parser.add_argument(
        "--token",
        default=None,
        help=f"知识库访问令牌（请求头默认 {fetch_knowledge.TOKEN_HEADER_DEFAULT}）；也可用环境变量 {fetch_knowledge.TOKEN_ENV}",
    )
    parser.add_argument(
        "--token-header",
        default=fetch_knowledge.TOKEN_HEADER_DEFAULT,
        help=f"令牌请求头名（默认 {fetch_knowledge.TOKEN_HEADER_DEFAULT}）",
    )
    parser.add_argument(
        "--api-base",
        default=None,
        help=f"接口根地址（默认：站点根 + {fetch_knowledge.API_PREFIX_DEFAULT}，可指向自己的服务）",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=fetch_knowledge.DEFAULT_LIMIT,
        help=f"在线抓取最多下载篇数（0 表示不限制，默认 {fetch_knowledge.DEFAULT_LIMIT}）",
    )
    parser.add_argument("--workspace", default=None, help="在线抓取限定工作区（工作区名称或 workspaceId）")
    parser.add_argument(
        "--any-quality",
        action="store_true",
        help="在线抓取不检查质量状态（默认只抓质量通过的文档）",
    )
    parser.add_argument(
        "--all-assets",
        action="store_true",
        help="在线抓取列表中的全部资产（不做资产类型 / 生命周期 / 质量状态过滤）",
    )
    parser.add_argument(
        "--kb-output",
        default=None,
        help="在线抓取的 Markdown 写入目录（默认 <输出根目录>/_kb_source/<时间戳>）",
    )
    parser.add_argument(
        "--output",
        default="output",
        help="输出根目录；每次运行在其下新建带时间戳的子目录（默认：output）",
    )
    parser.add_argument(
        "--report-name",
        default=audit_markdown.DEFAULT_REPORT_BASE_NAME,
        help="质量审查报告基名（默认：知识库文档解析与转写质量问题报告）",
    )
    parser.add_argument(
        "--max-file-size",
        type=int,
        default=audit_markdown.DEFAULT_MAX_FILE_SIZE,
        help="单文件大小上限（字节，默认 20MB），超过则跳过内容检查",
    )
    parser.add_argument("--rules", default=None, help="文本异常规则 JSON（默认：rules/ocr_patterns.json）")
    parser.add_argument("--alias-rules", default=None, help="属性别名与单位规则 JSON（默认：rules/attribute_aliases.json）")
    parser.add_argument("--fail-on-issue", action="store_true", help="发现问题时以退出码 3 结束")
    return parser


def crawl_online_input(args: argparse.Namespace, output_root: Path) -> Tuple[Optional[Path], int]:
    """--url 模式：先在线抓取知识库中已转写好的 Markdown，返回（抓取目录, 退出码）；成功时退出码为 0。"""
    token = (args.token or os.environ.get(fetch_knowledge.TOKEN_ENV) or "").strip()
    if not token:
        safe_print(f"[错误] 缺少访问令牌：请用 --token 传入，或设置环境变量 {fetch_knowledge.TOKEN_ENV}。")
        return None, 2
    try:
        api_base = (args.api_base or "").strip() or f"{fetch_knowledge.resolve_origin(args.url)}{fetch_knowledge.API_PREFIX_DEFAULT}"
    except ValueError as exc:
        safe_print(f"[错误] {exc}")
        return None, 2
    api_base = api_base.rstrip("/")
    kb_dir = Path(args.kb_output).expanduser() if args.kb_output else make_timestamped_dir(output_root / "_kb_source")
    try:
        kb_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        safe_print(f"[错误] 无法创建抓取目录 {kb_dir}：{exc}")
        return None, 1

    safe_print(f"[run_all] === 0/3 在线知识库抓取 ===（{args.url}）")
    try:
        crawled = fetch_knowledge.fetch_documents(
            api_base,
            token,
            kb_dir,
            limit=args.limit,
            workspace=args.workspace,
            require_pass=not args.any_quality,
            all_assets=args.all_assets,
            max_file_size=args.max_file_size,
            token_header=args.token_header,
        )
    except ValueError as exc:
        safe_print(f"[错误] {exc}")
        return None, 2
    except fetch_knowledge.KnowledgeApiError as exc:
        safe_print(f"[错误] {exc}")
        return None, 1
    fetch_knowledge.print_summary(crawled)
    if not crawled["documents"]:
        safe_print("[run_all] 未抓到任何文档，流水线中止。")
        return None, 1
    return kb_dir, 0


def run(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    if bool(args.url) == bool(args.input):
        safe_print("[错误] 请二选一：用 --url 给出在线知识库列表页地址，或给出本地 Markdown 文件/目录。")
        return 2
    output_root = Path(args.output).expanduser()
    if args.url:
        input_path, code = crawl_online_input(args, output_root)
        if input_path is None:
            return code
    else:
        input_path = Path(args.input).expanduser()
        if not input_path.exists():
            safe_print(f"[错误] 输入路径不存在：{input_path}")
            return 1
    try:
        run_dir = make_timestamped_dir(output_root)
    except OSError as exc:
        safe_print(f"[错误] 无法创建输出目录 {output_root}：{exc}")
        return 1

    planned = audit_markdown.collect_files(input_path, run_dir, args.report_name)
    safe_print(f"[run_all] 输入：{input_path.resolve()}")
    safe_print(f"[run_all] 交付目录（带完整时间戳）：{run_dir}")
    safe_print(f"[run_all] 待审查 Markdown 文件：{len(planned)} 个")
    steps = [
        (
            "1/3 文档质量审查",
            audit_markdown.run,
            [str(input_path), "--output", str(run_dir), "--report-name", args.report_name]
            + ["--max-file-size", str(args.max_file_size)]
            + (["--rules", args.rules] if args.rules else [])
            + (["--fail-on-issue"] if args.fail_on_issue else []),
        ),
        (
            "2/3 属性冲突比对",
            detect_conflicts.run,
            [str(input_path), "--output", str(run_dir)]
            + (["--alias-rules", args.alias_rules] if args.alias_rules else []),
        ),
    ]
    worst = 0
    for label, step_fn, step_argv in steps:
        safe_print(f"[run_all] === {label} ===")
        code = step_fn(step_argv)
        if code not in (0, 3):
            safe_print(f"[run_all] {label} 失败（退出码 {code}），流水线中止；已生成的报告仍可查看。")
            return code
        if code == 3:
            worst = 3

    safe_print(f"[run_all] 流水线完成，交付目录：{run_dir}")
    deliverables = sorted(fp for fp in run_dir.glob("*") if fp.is_file())
    safe_print(f"[run_all] 交付文件共 {len(deliverables)} 个（HTML 报告 + XLSX 清单表格）：")
    for fp in deliverables:
        safe_print(f"[run_all]   {fp.name}（{fp.stat().st_size} 字节）")
    if not deliverables:
        safe_print("[run_all]   未生成交付文件，请检查输入路径与上方日志。")
    if worst == 3:
        safe_print("[run_all] 已发现问题且指定了 --fail-on-issue，以退出码 3 结束。")
    return worst


def main() -> int:
    try:
        return run()
    except KeyboardInterrupt:
        safe_print("[错误] 用户中断操作。")
        return 130


if __name__ == "__main__":
    sys.exit(main())

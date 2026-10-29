#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""run_all.py — 一条命令跑完整条审查流水线（全自动）。

输入三选一：Markdown 目录、单个 .md、或用 --url 指定在线知识库列表页（先抓取再审查；不给输入时默认用 input/）。
依次执行：在线抓取（仅 --url）→ 转写质量审查 → 同一型号属性冲突比对 → 交付自检。
每次运行在输出根目录下新建时间戳子目录，固定产出 4 个交付文件（2 HTML + 2 XLSX）；在线抓取结果写入
<输出根目录>/_kb_source/<url哈希>/ 作为输入缓存，不计入交付文件。退出码：0=完成；1=输入/输出错误；
2=规则或参数错误；3=完成但发现问题（仅 --fail-on-issue）。在线抓取默认并发、断点续跑、失败补抓、连接自适应。
"""
from __future__ import annotations

import argparse
import hashlib
import os
import sys
import time
from pathlib import Path
from typing import List, Optional, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import audit_markdown  # noqa: E402
import check_deliverables  # noqa: E402
import detect_conflicts  # noqa: E402
import fetch_knowledge  # noqa: E402
from generate_report import format_duration, make_timestamped_dir, safe_print  # noqa: E402

DEFAULT_INPUT_DIR = "input"  # 无参运行时的默认输入目录（仓库内置演示语料就在 input/）
EXPECT_HTML = 2  # 交付物：2 个 HTML 报告（质量审查 + 冲突审查）
EXPECT_XLSX = 2  # 交付物：2 个 XLSX 清单表格


def file_sha256(path: Path) -> str:
    """交付文件指纹（只打印到控制台，便于事后核对文件是否被改动；不额外写任何文件）。"""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="run_all.py",
        description="一条命令跑完整条审查流水线：在线知识库抓取（可选）→ 转写质量审查 → 同一型号属性冲突比对",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/run_all.py input/\n"
            '  python scripts/run_all.py --url "https://kb.example.com/knowledge/asset" --token <token>\n'
            "  python scripts/run_all.py input/ --fail-on-issue     # 发现问题时退出码 3\n"
            "退出码：0=完成；1=输入/输出错误；2=规则或参数错误；3=完成但发现问题（仅 --fail-on-issue）。"
        ),
    )
    parser.add_argument("input", nargs="?", default=None,
                        help="待审查的 Markdown 文件或目录（与 --url 二选一；都不给时默认用 input/）")
    parser.add_argument("--output", default="output", help="输出根目录；每次运行在其下新建时间戳子目录（默认 output）")
    parser.add_argument("--report-name", default=audit_markdown.DEFAULT_REPORT_BASE_NAME,
                        help="质量审查报告基名（默认：知识库文档解析与转写质量问题报告）")
    parser.add_argument("--max-file-size", type=int, default=audit_markdown.DEFAULT_MAX_FILE_SIZE,
                        help="单文件大小上限（字节，默认 20MB），超过则跳过内容检查")
    parser.add_argument("--rules", default=None, help="文本异常规则 JSON（默认 rules/ocr_patterns.json）")
    parser.add_argument("--alias-rules", default=None, help="属性别名与单位规则 JSON（默认 rules/attribute_aliases.json）")
    parser.add_argument("--allow-empty", action="store_true", help="输入为空时也生成报告（默认以退出码 1 结束）")
    parser.add_argument("--fail-on-issue", action="store_true", help="发现问题时以退出码 3 结束")
    parser.add_argument("--no-progress", action="store_true", help="关闭进度条")

    online = parser.add_argument_group("在线知识库抓取（--url）")
    online.add_argument("--url", default=None, help="在线知识库列表页地址（如 https://kb.example.com/knowledge/asset）")
    online.add_argument("--token", default=None,
                        help=f"访问令牌（请求头默认 {fetch_knowledge.TOKEN_HEADER_DEFAULT}）；也可用环境变量 {fetch_knowledge.TOKEN_ENV}")
    online.add_argument("--token-header", default=fetch_knowledge.TOKEN_HEADER_DEFAULT, help="令牌请求头名")
    online.add_argument("--api-base", default=None,
                        help=f"接口根地址（默认：站点根 + {fetch_knowledge.API_PREFIX_DEFAULT}）")
    online.add_argument("--limit", type=int, default=fetch_knowledge.DEFAULT_LIMIT,
                        help=f"最多抓取篇数（0 不限，默认 {fetch_knowledge.DEFAULT_LIMIT}）")
    online.add_argument("--workspace", default=None, help="限定工作区（名称或 workspaceId）")
    online.add_argument("--any-quality", action="store_true", help="放宽质量状态（默认只抓质量通过的文档）")
    online.add_argument("--all-assets", action="store_true", help="抓取列表内全部资产（不做任何过滤）")
    online.add_argument("--kb-output", default=None, help="抓取 Markdown 的写入目录（默认 <输出根目录>/_kb_source/<url哈希>）")
    online.add_argument("--workers", type=int, default=fetch_knowledge.DEFAULT_WORKERS,
                        help=f"正文并发线程数（默认 {fetch_knowledge.DEFAULT_WORKERS}，1 为单线程）")
    online.add_argument("--retries", type=int, default=fetch_knowledge.DEFAULT_RETRIES, help="单请求失败重试次数")
    online.add_argument("--proxy", default=None, help="代理方式：auto（默认）/ direct / system / <host:port 或 URL>")
    online.add_argument("--overwrite", action="store_true", help="忽略本地缓存强制重抓")
    online.add_argument("--strict-fetch", action="store_true", help="任一文档漏取即中止流水线")
    online.add_argument("--timeout", type=int, default=fetch_knowledge.DEFAULT_TIMEOUT, help="接口超时秒数")
    online.add_argument("--keep-raw-content", action="store_true", help="不校验正文（默认拒收 JSON/HTML 接口错误体）")
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
    if args.kb_output:
        kb_dir = Path(args.kb_output).expanduser()
    else:
        # 固定目录（按 url 派生）而非每次新建时间戳目录：否则断点续跑永远没有可复用的历史文件、--overwrite/续跑形同虚设
        kb_slug = hashlib.sha1((args.url or "").encode("utf-8")).hexdigest()[:8]
        kb_dir = output_root / "_kb_source" / f"kb-{kb_slug}"
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
        return None, 2
    except fetch_knowledge.KnowledgeApiError as exc:
        safe_print(f"[错误] {exc}")
        return None, 1
    fetch_knowledge.print_summary(crawled)
    if not crawled["documents"]:
        safe_print("[run_all] 未抓到任何文档，流水线中止。")
        return None, 1
    # 失败门控：只有“一篇都没抓到”才中止；个别文档漏取（网络瞬断、超大文件等）只提示，不停摆
    # —— 否则一次网络抖动就能挡住后面几万篇文档的审查（需要强一致时用 --strict-fetch）
    failed = crawled["counters"].get("failed", 0)
    if failed:
        safe_print(
            f"[run_all] 抓取有 {failed} 篇未取到（已在上方列入“漏取风险”）：流水线继续，审查照常进行；"
            f"需要补抓可重跑同一条命令（默认断点续跑，只补缺口）。"
        )
        if args.strict_fetch:
            safe_print("[run_all] 已指定 --strict-fetch：按“任一漏取即中止”处理，流水线停止。")
            return None, 1
    return kb_dir, 0


def resolve_input(args: argparse.Namespace, output_root: Path) -> Tuple[Optional[Path], int]:
    """确定本次运行的输入：``--url`` 先在线抓取；输入与 ``--url`` 互斥；都不给时默认用 ``input/``。"""
    if args.url:
        if args.input:
            safe_print("[错误] --url 与输入路径只能二选一。")
            return None, 2
        return crawl_online_input(args, output_root)
    if args.input:
        input_path = Path(args.input).expanduser()
    else:
        input_path = Path(DEFAULT_INPUT_DIR).expanduser()
        if not input_path.exists():
            safe_print(
                "[错误] 请给出输入：待审查的 Markdown 文件或目录，或用 --url 给出在线知识库地址"
                f"（当前目录下也没有默认输入目录 {DEFAULT_INPUT_DIR}/）。"
            )
            return None, 2
        safe_print(f"[run_all] 未指定输入，使用默认输入目录 {DEFAULT_INPUT_DIR}/。")
    if not input_path.exists():
        safe_print(f"[错误] 输入路径不存在：{input_path}")
        return None, 1
    return input_path, 0


def run(argv: Optional[List[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    output_root = Path(args.output).expanduser()
    input_path, code = resolve_input(args, output_root)
    if input_path is None:
        return code
    try:
        run_dir = make_timestamped_dir(output_root)
    except OSError as exc:
        safe_print(f"[错误] 无法创建输出目录 {output_root}：{exc}")
        return 1

    planned = audit_markdown.collect_files(input_path, run_dir, args.report_name)
    safe_print(f"[run_all] 输入：{input_path.resolve()}")
    safe_print(f"[run_all] 交付目录（带完整时间戳）：{run_dir}")
    safe_print(f"[run_all] 待审查 Markdown 文件：{len(planned)} 个")
    if not planned and not args.allow_empty:
        # 空语料会生成“0 处问题”的报告，容易被误读成“知识库没问题”，因此默认以退出码 1 结束
        safe_print(
            f"[run_all] [警告] 未在 {input_path} 下匹配到任何 Markdown"
            f"（{'/'.join(audit_markdown.SUPPORTED_EXTENSIONS)}）：空语料生成“0 处问题”的报告容易被误读为"
            "“知识库没问题”，本次以退出码 1 结束。"
        )
        safe_print("[run_all] 确认要用空语料出报告时，请加 --allow-empty 重跑。")
        try:
            run_dir.rmdir()  # 不留空交付目录，避免干扰 check_deliverables 的“最新运行目录”判定
        except OSError:
            pass
        return 1
    steps = [
        (
            "1/3 文档质量审查",
            audit_markdown.run,
            [str(input_path), "--output", str(run_dir), "--report-name", args.report_name]
            + ["--max-file-size", str(args.max_file_size)]
            + (["--rules", args.rules] if args.rules else [])
            + (["--no-progress"] if args.no_progress else [])
            + (["--fail-on-issue"] if args.fail_on_issue else []),
        ),
        (
            "2/3 属性冲突比对",
            detect_conflicts.run,
            [str(input_path), "--output", str(run_dir)]
            + (["--alias-rules", args.alias_rules] if args.alias_rules else [])
            + (["--no-progress"] if args.no_progress else []),
        ),
    ]
    worst = 0
    for label, step_fn, step_argv in steps:
        safe_print(f"[run_all] === {label} ===")
        started = time.monotonic()
        code = step_fn(step_argv)
        safe_print(f"[run_all] {label} 用时 {format_duration(time.monotonic() - started)}（退出码 {code}）。")
        if code not in (0, 3):
            safe_print(f"[run_all] {label} 失败（退出码 {code}），流水线中止；已生成的报告仍可查看。")
            return code
        if code == 3:
            worst = 3

    safe_print(f"[run_all] 流水线完成，交付目录：{run_dir}")
    try:
        problems, files = check_deliverables.inspect(run_dir, EXPECT_HTML, EXPECT_XLSX)
    except OSError as exc:
        problems, files = [f"无法读取交付目录：{exc}"], []
    if problems:
        safe_print("[run_all] [警告] 交付物自检未通过：" + "；".join(problems))
        safe_print(f"[run_all]   目录内现有文件：{files}")
    else:
        safe_print(
            f"[run_all] 交付物自检通过：{len(files)} 个文件"
            f"（{EXPECT_HTML} 个 HTML 报告 + {EXPECT_XLSX} 个 XLSX 清单表格）。"
        )
    deliverables = sorted(fp for fp in run_dir.glob("*") if fp.is_file())
    safe_print(f"[run_all] 交付文件共 {len(deliverables)} 个（HTML 报告 + XLSX 清单表格）：")
    for fp in deliverables:
        try:
            safe_print(f"[run_all]   {fp.name}（{fp.stat().st_size} 字节，SHA256 {file_sha256(fp)[:12]}…）")
        except OSError as exc:  # 收尾统计读不到文件（被占用 / 被移走）只提示，不影响本次运行的退出码
            safe_print(f"[run_all]   {fp.name}（无法读取文件属性：{exc}）")
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

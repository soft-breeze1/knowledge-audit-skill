#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_deliverables.py — 校验一次流水线的交付物是否为 2 个 HTML 报告 + 2 个 XLSX 清单。

CI 卡口与本地自检共用；不依赖 shell 特性（Windows/Linux 行为一致），失败时发出 GitHub Actions 注解。
退出码：0=交付物齐全；1=没有运行目录或交付文件不符。
"""
from __future__ import annotations

import argparse
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import safe_print  # noqa: E402

# 运行目录命名见 generate_report.make_timestamped_dir：20261026-153012（同名时追加 _2、_3）
RUN_DIR_RE = re.compile(r"^\d{8}-\d{6}(?:_\d+)?$")


def escape_command_value(message: str) -> str:
    """按 GitHub Actions 工作流命令规则转义注解内容（避免 % / 换行被解析成控制字符）。"""
    return message.replace("%", "%25").replace("\r", "%0D").replace("\n", "%0A")


def annotate(level: str, message: str) -> None:
    """发出 GitHub Actions 注解（::error:: / ::notice::）；在本地运行时只是一行普通输出。"""
    safe_print(f"::{level}::{escape_command_value(message)}")


def run_dir_sort_key(path: Path) -> Tuple[str, int]:
    """运行目录排序键：先按时间戳，再按同名后缀的数字大小。

    目录名形如 ``20261026-153012`` 或同名冲突时的 ``20261026-153012_2`` / ``_10``。
    直接按字符串排序会把 ``_10`` 排在 ``_2`` 前面（字典序），导致“最新运行目录”判错；
    这里把后缀解析成整数比较，保证 ``_10`` 晚于 ``_2``。
    """
    base, sep, suffix = path.name.partition("_")
    return (base, int(suffix) if sep and suffix.isdigit() else 0)


def find_run_dirs(output_root: Path) -> List[Path]:
    """返回输出根目录下的运行目录，按时间戳、再按同名后缀数字升序排序。"""
    try:
        entries = list(output_root.iterdir())
    except OSError:
        return []
    return sorted(
        (path for path in entries if path.is_dir() and RUN_DIR_RE.match(path.name)),
        key=run_dir_sort_key,
    )


def inspect(run_dir: Path, expect_html: int, expect_xlsx: int) -> Tuple[List[str], List[str]]:
    """检查运行目录内的交付文件，返回 (问题清单, 实际文件名清单)；问题为空表示通过。"""
    try:
        files = sorted(path.name for path in run_dir.iterdir() if path.is_file())
    except OSError as exc:
        return [f"无法读取目录内容：{exc}"], []
    html = [name for name in files if name.lower().endswith(".html")]
    xlsx = [name for name in files if name.lower().endswith(".xlsx")]
    other = [name for name in files if name not in html and name not in xlsx]
    problems = []
    if len(html) != expect_html:
        problems.append(f"HTML 报告应为 {expect_html} 个，实际 {len(html)} 个")
    if len(xlsx) != expect_xlsx:
        problems.append(f"XLSX 清单应为 {expect_xlsx} 个，实际 {len(xlsx)} 个")
    if other:
        problems.append(f"存在非交付文件：{other}")
    return problems, files


def parse_args(argv: Optional[List[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        prog="check_deliverables.py",
        description="校验一次流水线的交付物（2 个 HTML 报告 + 2 个 XLSX 清单）。",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "示例：\n"
            "  python scripts/check_deliverables.py\n"
            "  python scripts/run_all.py input/ && python scripts/check_deliverables.py\n"
        ),
    )
    parser.add_argument("--output", default="output", help="输出根目录（默认 output）")
    parser.add_argument("--expect-html", type=int, default=2, help="期望的 HTML 报告数量（默认 2）")
    parser.add_argument("--expect-xlsx", type=int, default=2, help="期望的 XLSX 清单数量（默认 2）")
    return parser.parse_args(argv)


def run() -> int:
    args = parse_args()
    output_root = Path(args.output)
    run_dirs = find_run_dirs(output_root)
    if not run_dirs:
        if output_root.is_dir():
            try:
                detail = f"{output_root} 下现有条目：{sorted(path.name for path in output_root.iterdir())}"
            except OSError as exc:
                detail = f"无法读取目录 {output_root}：{exc}"
        else:
            detail = f"目录不存在：{output_root}"
        annotate("error", f"未找到运行目录（形如 20261026-153012 的时间戳子目录）；{detail}")
        return 1
    run_dir = run_dirs[-1]
    if len(run_dirs) > 1:
        safe_print(f"[提示] {output_root} 下有 {len(run_dirs)} 个运行目录，只校验最新的一个：{run_dir}")
    problems, files = inspect(run_dir, args.expect_html, args.expect_xlsx)
    if problems:
        for problem in problems:
            annotate("error", f"交付目录 {run_dir}：{problem}")
        safe_print(f"[失败] 交付目录 {run_dir} 实际有 {len(files)} 个文件：{files}")
        return 1
    annotate("notice", f"交付目录 {run_dir} 校验通过，共 {len(files)} 个文件：{files}")
    return 0


def main() -> int:
    try:
        return run()
    except KeyboardInterrupt:
        safe_print("[错误] 用户中断操作。")
        return 130


if __name__ == "__main__":
    sys.exit(main())
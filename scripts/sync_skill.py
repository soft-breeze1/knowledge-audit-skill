#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""sync_skill.py — 将项目根目录的 SKILL.md 同步到两个 Agent 的项目级发现目录。

依据官方规范：
- Claude Code 项目级 Skill 路径：.claude/skills/<skill-name>/SKILL.md
- Codex 项目级 Skill 路径（Git 仓库内）：.agents/skills/<skill-name>/SKILL.md
- 两者均遵循 Agent Skills 开放标准（agentskills.io），
  frontmatter 仅使用标准字段（name/description/license/compatibility/metadata/allowed-tools），
  因此同一份 SKILL.md 可以安全复用；本脚本用“复制 + 校验”方式维护两份副本，
  防止内容漂移。

用法：
    python scripts/sync_skill.py            # 执行同步（复制到两个发现目录）
    python scripts/sync_skill.py --check    # 仅校验一致性（不一致时退出码 1）

校验内容：两个发现目录的 SKILL.md 与根目录真源逐字节一致；且 SKILL.md 的 ``metadata.version`` 与
``scripts/generate_report.py`` 的 ``TOOL_VERSION`` 一致（避免版本号在两处漂移）。
"""
from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path
from typing import List, Optional

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import TOOL_VERSION, safe_print  # noqa: E402

PROJECT_ROOT = SCRIPT_DIR.parent
SOURCE = PROJECT_ROOT / "SKILL.md"
SKILL_NAME = "knowledge-audit"
# SKILL.md frontmatter 里的版本号（本仓库真源是 generate_report.TOOL_VERSION，这里校验两处不漂移）
VERSION_RE = re.compile(r"^\s*version:\s*[\"']?([0-9][0-9A-Za-z.\-]*)[\"']?\s*$", re.MULTILINE)
TARGETS = [
    PROJECT_ROOT / ".claude" / "skills" / SKILL_NAME / "SKILL.md",
    PROJECT_ROOT / ".agents" / "skills" / SKILL_NAME / "SKILL.md",
]


def sha256_of(path: Path) -> Optional[str]:
    """文件 SHA256；读取失败（权限、被占用等）返回 None，由调用方给出清晰错误而不是崩栈。"""
    try:
        return hashlib.sha256(path.read_bytes()).hexdigest()
    except OSError:
        return None


def skill_declared_version() -> Optional[str]:
    """取 SKILL.md frontmatter 里声明的版本号（``metadata.version``）；读不到返回 None。"""
    try:
        text = SOURCE.read_text(encoding="utf-8")
    except OSError:
        return None
    m = VERSION_RE.search(text)
    return m.group(1) if m else None


def check_version() -> List[str]:
    """版本一致性检查：SKILL.md 的 metadata.version 应与 generate_report.TOOL_VERSION 相同。"""
    declared = skill_declared_version()
    if declared is None:
        return [f"SKILL.md 未找到 metadata.version（报告页脚与运行信息用的是 TOOL_VERSION={TOOL_VERSION}）"]
    if declared != TOOL_VERSION:
        return [
            f"版本号不一致：SKILL.md metadata.version={declared}，"
            f"generate_report.TOOL_VERSION={TOOL_VERSION}"
        ]
    return []


def sync(check_only: bool = False) -> int:
    if not SOURCE.is_file():
        safe_print(f"[错误] 未找到源文件：{SOURCE}（请先在项目根目录创建 SKILL.md）")
        return 1
    source_hash = sha256_of(SOURCE) if SOURCE.is_file() else None
    if source_hash is None:
        safe_print(f"[错误] 无法读取源文件：{SOURCE}（请检查文件权限或是否被占用）")
        return 1
    failures = []
    for target in TARGETS:
        rel = target.relative_to(PROJECT_ROOT)
        if target.is_file() and sha256_of(target) == source_hash:
            safe_print(f"[一致] {rel}")
            continue
        if check_only:
            safe_print(f"[不一致] {rel}（缺失或内容与源文件不同）")
            failures.append(target)
            continue
        try:
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(SOURCE.read_bytes())
        except OSError as exc:
            safe_print(f"[错误] 无法写入 {rel}：{exc}")
            failures.append(target)
            continue
        safe_print(f"[已同步] SKILL.md -> {rel}")
    if failures:
        safe_print(f"[错误] {len(failures)} 个目标与源文件不一致，请执行 python scripts/sync_skill.py 同步。")
    version_problems = check_version()
    for problem in version_problems:
        safe_print(f"[错误] {problem}（两处需同步修改）")
    if failures or version_problems:
        return 1
    safe_print("[完成] SKILL.md 同步/校验结束（副本内容与版本号均一致）。")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="同步 SKILL.md 到 Claude Code 与 Codex 的项目级发现目录"
    )
    parser.add_argument(
        "--check", action="store_true", help="仅校验一致性（副本内容 + 版本号），不修改任何文件"
    )
    args = parser.parse_args(argv)
    return sync(check_only=args.check)


if __name__ == "__main__":
    sys.exit(main())

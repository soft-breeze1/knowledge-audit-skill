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
"""
from __future__ import annotations

import argparse
import hashlib
import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

from generate_report import safe_print  # noqa: E402

PROJECT_ROOT = SCRIPT_DIR.parent
SOURCE = PROJECT_ROOT / "SKILL.md"
SKILL_NAME = "knowledge-audit"
TARGETS = [
    PROJECT_ROOT / ".claude" / "skills" / SKILL_NAME / "SKILL.md",
    PROJECT_ROOT / ".agents" / "skills" / SKILL_NAME / "SKILL.md",
]


def sha256_of(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sync(check_only: bool = False) -> int:
    if not SOURCE.is_file():
        safe_print(f"[错误] 未找到源文件：{SOURCE}（请先在项目根目录创建 SKILL.md）")
        return 1
    source_hash = sha256_of(SOURCE)
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
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(SOURCE.read_bytes())
        safe_print(f"[已同步] SKILL.md -> {rel}")
    if failures:
        safe_print(f"[错误] {len(failures)} 个目标与源文件不一致，请执行 python scripts/sync_skill.py 同步。")
        return 1
    safe_print("[完成] SKILL.md 同步/校验结束。")
    return 0


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(
        description="同步 SKILL.md 到 Claude Code 与 Codex 的项目级发现目录"
    )
    parser.add_argument(
        "--check", action="store_true", help="仅校验一致性，不修改任何文件"
    )
    args = parser.parse_args(argv)
    return sync(check_only=args.check)


if __name__ == "__main__":
    sys.exit(main())

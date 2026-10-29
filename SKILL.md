---
name: knowledge-audit
license: MIT
metadata:
  version: "1.1.0"
  author: knowledge-audit-skill contributors
  compatibility: "Python 3.9+，仅标准库；Windows / macOS / Linux"
description: 企业知识库质量审查（Agent Skill，Python 3.9+、纯标准库）。输入三选一：已转写为 Markdown 的知识库文件夹、单个 .md 文档，或在线知识库文档列表页地址（配访问令牌，自动抓取后审查）。一条命令完成两类客观审查并产出带证据定位（文件 + 行号 + 证据原文）的交付物：① 转写质量审查（固定 6 项检查）；② 同一型号同类属性取值冲突检测（型号 + 属性 + 工况一致且证据来自 ≥2 个文件，支持单位换算与属性别名合并）。输出 2 份 HTML 报告 + 2 份 XLSX 清单表格，不生成 Markdown / JSON、不修改输入。用于知识库体检、文档转写问题排查、多文档参数一致性比对。
---

# 企业知识库质量审查（knowledge-audit）

对 Markdown 文档做两类**客观、可复核**的审查，每类各产出一份 HTML（人读）+ 一份 XLSX 清单（逐行可筛选），共 4 个文件；**不修改输入、不生成 Markdown / JSON**。每次运行写入 `output/<YYYYMMDD-HHMMSS>/` 时间戳子目录，一次任务一个目录、互不覆盖。

## 何时使用

- 用户提供 Markdown 文档或目录，要求做质量检查 / 排查转写问题；
- 用户怀疑多份产品文档中同一型号同一参数（功率、重量、电压等）互相矛盾，要求比对；
- 用户要求对知识库做体检 / 审计 / 一致性检查，或给一个**知识库网址** + `x-access-token` 要求先抓取再审查。

## 快速开始

在项目根目录用 Python 3.9+ 运行（无第三方依赖）。默认一条命令跑完整流程（不给输入时用 `input/`）：

```bash
python scripts/run_all.py                         # 审查 input/（内置演示语料）
python scripts/run_all.py <目录或单个.md>
python scripts/run_all.py --url "<知识库列表页地址>" --token "<x-access-token>"
```

输入三选一：单个 `.md`/`.markdown` 文件、Markdown 目录（递归扫描，支持中文路径），或 `--url`（配合 `--token` 或环境变量 `KB_ACCESS_TOKEN`，先自动抓取再审查）。分步执行：

```bash
python scripts/fetch_knowledge.py "<地址>" --token "<token>"   # ⓪ 在线抓取（可选）
python scripts/audit_markdown.py <目录或.md>                    # ① 转写质量审查
python scripts/detect_conflicts.py <目录或.md>                  # ② 属性冲突比对
```

## 执行主链路

输入判定 →（可选）在线抓取 → 转写质量审查 → 属性冲突比对 → 交付自检。任一步失败即中止（已生成的报告仍可查看）；三个阶段在交互式终端显示进度条，重定向到文件 / CI 日志时自动降级为分段整行输出。

## 输入与输出

- **输入**：只需已转写好的 Markdown；不需要原始 DOCX/PDF/图片，也不核对原始文件。
- **交付物**（4 个，均写入本次时间戳子目录）：
  - `知识库文档解析与转写质量问题报告.html` / `.xlsx`
  - `知识库同类属性冲突审查报告.html` / `.xlsx`
- HTML 单文件、样式内联，双击即开可转发；XLSX 由标准库生成（表头深蓝底白字、列宽、边框、隔行斑马纹、首行冻结并开启自动筛选），Excel / WPS 双击可开。
- **XLSX 列**：
  - 质量清单：`文件 / 行号 / 问题类型 / 问题说明 / 证据原文`（问题类型即 6 项检查项名称，逐行一行一处）；
  - 冲突清单：`冲突编号 / 型号 / 属性 / 描述 / 来源文档 / 行号 / 证据原文`（按「冲突编号」可筛出同一处冲突涉及哪几个文档）。
- 在线抓取的 Markdown 属**输入缓存**，写入 `<输出根目录>/_kb_source/<url哈希>/`，与交付目录分开、不计入交付文件；默认断点续跑，中断后重跑只补缺口。
- **退出码**：0=完成；1=输入/输出错误（含“未匹配到任何 Markdown”，`--allow-empty` 可放开）；2=规则或参数错误；3=完成但发现问题（仅 `--fail-on-issue`）。

## 常用参数

质量与冲突：`--output`（输出根目录）、`--report-name`、`--rules`（文本异常规则）、`--alias-rules`（别名与单位规则）、`--max-file-size`（默认 20MB）、`--allow-empty`、`--fail-on-issue`、`--no-progress`。

在线抓取（`--url`，run_all 与 fetch_knowledge 通用）：`--token` / `KB_ACCESS_TOKEN`、`--token-header`（默认 `x-access-token`）、`--api-base`（默认站点根 + `/api/v1`）、`--limit`（默认 200，0 不限）、`--workspace`、`--any-quality`、`--all-assets`、`--workers`（默认 8）、`--retries`（默认 3）、`--proxy`（`auto`/`direct`/`system`/地址）、`--overwrite`、`--timeout`、`--keep-raw-content`；`--strict-fetch`（仅 run_all：任一漏取即中止）。

## 检查口径（规则式，非语义推断）

**转写质量（固定 6 项检查，分类即检查项）**：编码可解码性（UTF-8） / 替换字符 U+FFFD / 编码乱码特征 / OCR 标记残留 / 占位符文本 / 空文件检测。6 项问题**全部**计入「存在问题文件数 / 问题总数」并逐行写入清单；表格列数/空单元格、空章节、代码块未闭合、疑似截断、图片引用、标题格式等**结构性写法不上报**；代码块内只查替换字符与乱码特征。读取失败或超大小上限属“漏审风险”，不计入统计，只在控制台提示并写入 HTML 运行信息。

**属性冲突**：分组键「型号 + 属性 + 工况」完全一致才比较；型号按原文精确匹配并保留变型词（`ABCDEF-100 HEV` ≠ `ABCDEF-100`）；**取值不一致且证据来自 ≥2 个不同文件**才列入。抽取只认两列表、“属性-取值-单位”列组表（单位列在前或后、支持左右并排多组）、排版表相邻“标签-数值”对、正文 `属性：数值+单位`；单位按 `rules/attribute_aliases.json` 换算、属性按同义词合并；标识类字段（物料编码/零件号/序号）、故障排查表、备件清单、版本履历表、范围/多维复合取值不抽取；未识别型号或单一来源不参与比对（报告只给数量）。

## 结果解读

两个阶段都**只列问题、不设分级**，每条结论附文件 + 行号 + 证据原文；无法仅凭文本确认的情形在“说明”中如实写明，不武断下结论。HTML 中高频同型问题按类型聚合（给数量与行号汇总，每个「文件 + 问题类型」最多逐条列 5 处），逐行完整清单在 XLSX。汇报可先给概览、再列重点（附行号）、最后给建议，或直接指路同目录的 HTML / XLSX。不要虚构报告里没有的内容。

## 规则与边界

- 规则可配置：`rules/ocr_patterns.json`（OCR 标记 / 占位符 / 乱码特征，pattern 为正则）、`rules/attribute_aliases.json`（属性别名、单位换算）。缺失时回退内置精简规则；显式指定但不可用则退出码 2。
- **零依赖、零配置**：仅标准库；并发、重试、断点续跑、直连/系统代理自适应在默认值下即“快且稳”，无需改 `no_proxy` 等环境变量。
- 只审查 Markdown 文本；输入只读；扫描时自动排除输出目录与报告文件本身。
- **唯一真源**：根目录 `SKILL.md`；`.claude/skills/` 与 `.agents/skills/` 下为同步副本，改动后执行 `python scripts/sync_skill.py`（`--check` 校验）。

<div align="center">

# knowledge-audit-skill

**知识库质量审查 · Agent Skill**

一条命令查两类问题：文档转写质量、同型号属性冲突<br/>
结论均带证据定位（文件 / 行号 / 证据原文）· 只读输入 · 零第三方依赖

[![Python](https://img.shields.io/badge/Python-3.9%2B-3772ab?logo=python&logoColor=white)](https://www.python.org/)
[![Dependencies](https://img.shields.io/badge/%E4%BE%9D%E8%B5%96-0-2ea44f)](#快速开始)
[![License](https://img.shields.io/badge/License-MIT-16a34a)](LICENSE)
[![Agent Skills](https://img.shields.io/badge/Agent%20Skills-%E5%BC%80%E6%94%BE%E6%A0%87%E5%87%86-2563eb)](https://agentskills.io)

</div>

> 输入三选一：Markdown 文件夹 · 单个 `.md` 文档 · 在线知识库地址（配访问令牌自动抓取）。
> 每次运行固定产出 **4 个文件**（2 份 HTML 报告 + 2 份 XLSX 清单），不修改输入、不生成 Markdown / JSON。

## 目录

- [特性](#特性)
- [快速开始](#快速开始)
- [检查什么](#检查什么)
- [交付物](#交付物)
- [在线抓取（可选）](#在线抓取可选)
- [安装](#安装)
- [命令行](#命令行)
- [口径](#口径)
- [规则配置](#规则配置)
- [CI](#ci)
- [常见问题](#常见问题)
- [版本](#版本)

## 特性

- **两类客观审查**：① 转写质量（固定 6 项检查）；② 同一型号同类属性取值冲突（支持单位换算与属性别名合并）。
- **证据可复核**：每条结论都定位到「文件 + 行号 + 证据原文」。
- **只读、零依赖**：仅用 Python 标准库，不修改任何输入文档。
- **开箱即用**：一条命令跑完整流程，自动生成带样式的 HTML / XLSX 交付物。
- **在线可用**：给 `--url` + `--token` 即先抓取知识库存量 Markdown 再审查（并发、断点续跑、失败补抓、连接自适应）。

## 快速开始

输入是**已经转写好的 Markdown**，不需要原始 DOCX / PDF / 图片。

```bash
# 下载/克隆本仓库后，在仓库根目录执行：

python scripts/run_all.py          # 不给输入时默认用 input/（内置演示语料）
# 结果在 output/<时间戳>/，打开两份 HTML 即可
```

常用变体：

```bash
python scripts/run_all.py input/ --fail-on-issue                 # 发现问题时退出码 3（CI 卡口）
python scripts/run_all.py --url "<知识库地址>" --token "<token>"   # 在线知识库：先抓取再审查
python scripts/audit_markdown.py input/                          # 只跑 ① 转写质量审查
python scripts/detect_conflicts.py input/                        # 只跑 ② 属性冲突比对
```

`input/` 是虚构的演示语料：`example.md` 覆盖各类转写缺陷，`products/` 下两份 `机器ABCDEF-100` 留了一处冲突。预期结果见 [`input/README.md`](input/README.md)。

## 检查什么

| 常见问题 | 对应审查 |
| --- | --- |
| 转成 Markdown 后出现乱码、`U+FFFD`、`[OCR]`、`待补充`、空文件 | **① 转写质量审查**：固定 6 项，逐行给出文件、行号、证据 |
| 两份手册写了同一台机器的“操作重量”，一份 12300 kg、一份 12400 kg | **② 属性冲突检测**：按「型号 + 属性 + 工况」比对，列出全部来源与取值 |

- 只报问题、不设分级，性质判断交给人；
- 规则与单位换算在 `rules/*.json`，报告会写明检查项、规则来源与生成时间；
- 冲突只认来自 **≥2 个不同文件** 的不一致，同一文件内重复不算。

## 交付物

每次运行在 `output/<YYYYMMDD-HHMMSS>/` 固定产出 4 个文件，一次任务一个目录、互不覆盖：

| 文件 | 内容 |
| --- | --- |
| `知识库文档解析与转写质量问题报告.html` | 概况、6 项检查分布、按文件分组的问题明细（行号 + 证据） |
| `知识库文档解析与转写质量问题报告.xlsx` | 逐行清单：文件、行号、问题类型、问题说明、证据原文 |
| `知识库同类属性冲突审查报告.html` | 比对概况，每组冲突的来源文档、行号、取值描述、证据 |
| `知识库同类属性冲突审查报告.xlsx` | 逐行清单：冲突编号、型号、属性、描述、来源文档、行号、证据原文 |

HTML 单文件、样式内联，双击即开可直接转发；XLSX 表头着色、列宽、斑马纹、首行冻结并开启自动筛选。

## 在线抓取（可选）

给 `--url` + `--token` 时，先抓取知识库中已转写好的 Markdown，再进入审查。抓取结果作为**输入缓存**写入 `output/_kb_source/kb-<hash>/`，不计入交付文件；重跑默认断点续跑，只补缺口。

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `--limit` | `200` | 最多抓多少篇，`0` 不限 |
| `--workers` | `8` | 正文并发线程数（1 = 单线程，上限 32） |
| `--retries` | `3` | 单请求失败重试次数（指数退避），结束后再补抓几轮 |
| `--proxy` | `auto` | 自动选更快的直连 / 系统代理；也可 `direct` / `system` / 代理地址 |
| `--overwrite` | 关闭 | 忽略本地缓存强制重抓 |
| `--all-assets` | 关闭 | 抓列表内全部资产，不做类型 / 生命周期 / 质量过滤 |
| `--strict-fetch` | 关闭 | （`run_all.py` 提供）任一漏取即中止；默认漏取只提示，审查照常 |
| `--api-base` | `<站点>/api/v1` | 接口根地址（前端与接口不同端口时用这个指定，如 `http://host:8080/admin/api/v1`） |

列表分页、资产过滤与「同一文档只抓第一份」按列表顺序单线程判定，并发只用于取正文，结果与单线程一致。接口返回 200 但正文是 JSON / HTML 报错页时会被拒收并单独计数。

## 安装

本技能遵循 [Agent Skills 开放标准](https://agentskills.io/specification)：一个目录即一个技能，至少含 `SKILL.md`；本技能还需要同级的 `scripts/` 与 `rules/`。

**安装 = 把仓库根目录的 `SKILL.md` + `scripts/` + `rules/` 放进目标技能目录。** 仓库里 `.claude/skills/knowledge-audit/` 只是同步好的 `SKILL.md` 副本，直接拿它并不完整，请按下面的命令组装。

各客户端的技能目录：

| 使用方式 | 技能目录（放入 `SKILL.md` + `scripts/` + `rules/`） |
| --- | --- |
| CCSwitch（Windows 图形化） | `%USERPROFILE%\.claude\skills\knowledge-audit\` |
| Claude Code · 个人级（所有项目） | `~/.claude/skills/knowledge-audit/` |
| Claude Code · 项目级 | `<你的项目>/.claude/skills/knowledge-audit/` |
| Codex · 项目级 | `<你的项目>/.agents/skills/knowledge-audit/` |

### 1. 先组装技能目录

**Windows PowerShell**（在仓库根目录执行，装成 Claude Code 个人级为例）：

```powershell
$dst = "$env:USERPROFILE\.claude\skills\knowledge-audit"
New-Item -ItemType Directory -Force -Path "$dst\scripts","$dst\rules" | Out-Null
Copy-Item SKILL.md  "$dst\SKILL.md"
Copy-Item scripts\*.py "$dst\scripts\"
Copy-Item rules\*.json "$dst\rules\"
```

**macOS / Linux / Git Bash**：

```bash
DST=~/.claude/skills/knowledge-audit
mkdir -p "$DST/scripts" "$DST/rules"
cp SKILL.md "$DST/"
cp scripts/*.py "$DST/scripts/"
cp rules/*.json "$DST/rules/"
```

> 装到别的目录（项目级 / Codex）时，把 `$dst` / `DST` 换成上表对应的路径即可。

### 2. CCSwitch

CCSwitch 的「Skills」页管理的就是本地技能目录（默认 `%USERPROFILE%\.claude\skills\`）。

1. 按上一步组装到该目录；
2. 回到 CCSwitch → Skills 页 → 点右上角 **「添加 Skill」「添加」**（选“从本地目录”指向该目录），或点 **「检查更新」** 让它自动识别；
3. 在列表里给 `knowledge-audit` 勾选要启用的客户端（Claude Code / Claude Desktop 等）。

### 3. Claude Code

- **个人级（所有项目可用）**：组装到 `~/.claude/skills/knowledge-audit/`；
- **项目级（只对该项目生效）**：组装到 `<项目>/.claude/skills/knowledge-audit/`，随项目一起提交。

启动 `claude` 后即可用自然语言触发，例如：

> 帮我用 knowledge-audit 审查 `input/` 这批 Markdown 的转写质量

也可以显式调用 `/knowledge-audit <目录>`。

### 4. Codex

把技能组装到 Git 仓库内的 `.agents/skills/knowledge-audit/`（Codex 会扫描仓库里的 `.agents/skills/`）；个人级可放到 `~/.agents/skills/knowledge-audit/`（以你的 Codex 配置为准）。之后在会话中描述任务即可触发。

### 验证与升级

- **验证**：技能目录里应能看到 `SKILL.md`、`scripts/`、`rules/` 三样；进入该目录运行 `python scripts/run_all.py <一个含 .md 的目录>` 能产出 4 个交付文件即安装成功。
- **升级**：仓库更新后，重新执行「先组装技能目录」的复制命令即可（`git pull` 后覆盖 `SKILL.md`、`scripts/`、`rules/`）。

> 根目录 `SKILL.md` 是唯一真源，`.claude/`、`.agents/` 下是同步副本。改完执行 `python scripts/sync_skill.py`，校验用 `python scripts/sync_skill.py --check`。

## 命令行

`run_all.py`（推荐入口）：

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `input` | `input/` | 待审查的 Markdown 文件或目录（与 `--url` 二选一） |
| `--output` | `output` | 输出根目录，每次运行在其下新建时间戳子目录 |
| `--report-name` | `知识库文档解析与转写质量问题报告` | 质量审查报告基名 |
| `--rules` / `--alias-rules` | `rules/ocr_patterns.json` / `rules/attribute_aliases.json` | 规则文件 |
| `--max-file-size` | 20 MB | 单文件上限，超过则跳过内容级检查 |
| `--fail-on-issue` | 关闭 | 发现问题时以退出码 3 结束 |
| `--allow-empty` | 关闭 | 目录内无 Markdown 时也出报告（默认退出码 1，避免空语料被误读成“没问题”） |

单阶段脚本：

| 脚本 | 作用 |
| --- | --- |
| `audit_markdown.py` | 只跑转写质量审查 |
| `detect_conflicts.py` | 只跑属性冲突比对 |
| `fetch_knowledge.py` | 只跑在线抓取 |
| `sync_skill.py` | 同步 / 校验 `SKILL.md` 副本（`--check` 仅校验） |
| `check_deliverables.py` | 校验运行目录里是 2 HTML + 2 XLSX（CI 与本地自检共用） |

退出码：`0` 完成；`1` 输入 / 输出错误（含网络失败、未匹配到 Markdown）；`2` 参数或规则错误；`3` 完成但发现问题（仅 `--fail-on-issue`）。

## 口径

**转写质量（固定 6 项）**：编码可解码性（UTF-8）、替换字符 U+FFFD、编码乱码特征、OCR 标记残留、占位符文本、空文件检测。6 项问题全部计入统计并逐行写入清单。表格结构、空章节、代码块未闭合、疑似截断等结构性写法不上报；代码块内只查替换字符与乱码特征。读取失败或超过大小上限的文件属“漏审风险”，不计入统计，只在控制台与报告运行信息里提示。

**属性冲突**：分组键为「型号 + 属性 + 工况」完全一致才比较，型号按原文精确匹配并保留变型词（`ABCDEF-100 HEV` ≠ `ABCDEF-100`）；列为冲突需取值不一致且来自 ≥2 个不同文件。抽取是规则式的（非语义理解）：只认两列表、“属性-取值-单位”列组表、正文 `属性：数值+单位`；标识类字段、故障排查表、备件清单、版本履历表、范围 / 多维复合取值不抽取。

## 规则配置

两份都是普通 JSON，改完即生效，文件缺失或无法解析时回退内置精简规则。

```json
// rules/ocr_patterns.json —— pattern 为 Python re 正则，flags 可选 IGNORECASE / MULTILINE
{ "rules": {
  "ocr_markers": [ { "id": "ocr-bracket-en", "pattern": "(?<!!)\\[\\s*OCR\\s*\\](?!\\()", "flags": ["IGNORECASE"] } ],
  "placeholders": [ { "id": "wait-supplement", "pattern": "待补充" } ],
  "mojibake":     [ { "id": "gbk-mojibake", "pattern": "锟斤拷|烫烫烫" } ]
} }
```

```json
// rules/attribute_aliases.json —— 标准属性名 → 同义词；单位写法 → 标准符号与换算因子
{ "attribute_aliases": { "重量": ["整机重量", "操作重量", "weight"] },
  "unit_symbols": { "吨": "t", "千瓦": "kW", "转/分": "r/min" },
  "units": { "mass": { "canonical": "kg", "factors": { "kg": 1.0, "t": 1000.0 } } } }
```

用 `--rules` / `--alias-rules` 指定自定义规则文件；显式指定但不可用时退出码 2。

## CI

`--fail-on-issue` 把“发现问题”变成可卡口的构建失败。仓库自带 [`.github/workflows/knowledge-audit.yml`](.github/workflows/knowledge-audit.yml)，在 Ubuntu / Windows × Python 3.9 / 3.12 上跑演示语料，校验 `SKILL.md` 副本一致性与交付文件数量，并上传报告。

```yaml
- run: python scripts/sync_skill.py --check
- run: python scripts/run_all.py input/ --fail-on-issue
- run: python scripts/check_deliverables.py
```

在线模式需把令牌（如 `KB_ACCESS_TOKEN`）配成仓库 Secret。

## 常见问题

**会改我的文档吗？要装依赖吗？能审 DOCX / PDF 吗？**
不改、不装依赖。输入全程只读，只用标准库。本技能只审**已经转写好的 Markdown**，不解析原始 DOCX / PDF / 图片，也不承诺“转写内容与原始文档逐字一致”。

**同一份文档里重复的参数会被当冲突吗？**
不会，冲突要求证据来自 ≥2 个不同文件。

**我明明看到不一致，报告里却没有？**
按顺序自查：型号是否与原文完全一致；属性名是否归一到同一标准名（`rules/attribute_aliases.json`）；工况限定词（额定 / 最大 / 最小 / 尾部括注）是否一致；证据是否来自 ≥2 个文件。仍不确定时，用两个 `.md` 组成最小语料单独跑一次最容易定位。

## 版本

| 版本 | 变化 |
| --- | --- |
| `1.1.0` | 精简重构：合并报告渲染到 `generate_report.py`、去除冗余代码与冗长注释；在线抓取参数在 `--help` 中分组；修正文档中过期的清单列说明；移除演示与抓取缓存等无关产物 |
| `1.0.6` | 属性冲突排除故障排查表并精简报告列（去工况、取值合并为「描述」、文档改「来源文档」）；质量问题清单去掉「建议操作」列；在线抓取目录固定化使断点续跑生效，接口根报错更明确、抓取文件名去双扩展名；XLSX / HTML 报告样式美化 |
| `1.0.5` | 写盘失败与单文件异常不再中断整批；规则加载宽容化并在报告注明跳过原因；证据按命中处截取；支持零参数与空语料门控；校验 SKILL 版本号一致 |
| `1.0.4` | 清理冗余代码，修复进度条与漏审风险展示，精简 SKILL 描述 |
| `1.0.3` | 精简 README，补充版本记录 |
| `1.0.2` | 修复 Windows 上 CI 交付文件校验失败（改用 `check_deliverables.py`） |
| `1.0.1` | 在线抓取支持并发、断点续跑、失败重试与连接自适应；增加进度条与统计自检 |
| `1.0.0` | 初版：转写质量审查（固定 6 项）+ 同型号属性冲突检测 |

改口径后的回归建议：跑一遍 `input/` 核对预期结果 → `python scripts/sync_skill.py` → 提交。

## 许可证

MIT，见 [LICENSE](LICENSE)。可自由用于商业与个人项目，保留版权声明即可。

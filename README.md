# china-memory-sector · 中国存储芯片板块分析技能

一条命令生成**中文** Excel 估值底稿 + **中文** Word 研究报告。Claude Code、Codex、Cursor、Gemini 共用同一份技能。

## 安装位置（单一真源）
```
~/.agents/skills/china-memory-sector      ← 真实目录（git 管理）
~/.claude/skills/china-memory-sector      → 软链接
~/.codex/skills/china-memory-sector       → 软链接
~/.cursor/skills/china-memory-sector      → 软链接
~/.gemini/skills/china-memory-sector      → 软链接
```
只改真实目录，所有 Agent 立即生效。

## 快速使用
```bash
CMS=~/.agents/skills/china-memory-sector/bin/cms
$CMS fetch   ~/research/memory-2026-11      # 抓数据（首次自动建 Python 环境）
$CMS summary ~/research/memory-2026-11      # 打印关键数据
# Agent 联网补充 industry.json，并撰写中文 analysis.json（参考 examples/）
$CMS all     ~/research/memory-2026-11      # 生成 Excel + Word 并校验公式
```
对任意 Agent 说“做一份存储芯片板块分析”即可触发完整流程（见 SKILL.md）。

## 目录
- `SKILL.md` —— Agent 执行说明（含中文输出规则）
- `bin/cms` —— 统一命令入口
- `scripts/` —— 抓数、Excel、报告、校验脚本
- `references/` —— 公司池、分析框架、报告结构
- `examples/` —— 2026-11-09 样例 analysis.json / industry.json

## 修改样本
编辑 `references/universe.csv`（细分板块须为：存储设计 / 接口与主控 / 模组与存储 / 分销），
或临时指定：`UNIVERSE=/path/to/my.csv $CMS fetch <workdir>`。

## 多台电脑同步
仓库：https://github.com/LeYoungChen/china-memory-sector（私有）

**本机修改后上传：**
```bash
cd ~/.agents/skills/china-memory-sector && git add -A && git commit -m "更新说明" && git push
```

**其他电脑首次安装：**
```bash
git clone https://github.com/LeYoungChen/china-memory-sector.git ~/.agents/skills/china-memory-sector
for a in claude codex cursor gemini; do mkdir -p ~/.$a/skills && ln -sfn ../../.agents/skills/china-memory-sector ~/.$a/skills/china-memory-sector; done
~/.agents/skills/china-memory-sector/bin/cms setup
```

**其他电脑拉取更新：**
```bash
cd ~/.agents/skills/china-memory-sector && git pull
```

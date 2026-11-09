# china-memory-sector · 中国存储芯片板块分析技能

一条命令生成**中文** Excel 估值底稿 + **中文** Word 研究报告。Claude Code、Codex、Cursor、Gemini 共用同一份技能。

覆盖 12 家 A 股存储公司，并与 13 家海外同业对标（三星、SK 海力士、美光、铠侠、闪迪、南亚科、华邦、旺宏、Rambus、慧荣、群联、威刚、创见）。

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
$CMS fetch   ~/research/memory-2026-11      # 抓 A 股 + 海外数据（首次自动建 Python 环境）
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
海外同业编辑 `references/universe_global.csv`（Yahoo Finance 代码，如 `000660.KS`、`2344.TW`、`MU`；细分须为：存储原厂 / 利基存储 / 接口与主控 / 模组与存储），
或临时指定：`UNIVERSE_GLOBAL=/path/to/global.csv $CMS fetch <workdir>`。

## 多台电脑同步
仓库：https://github.com/LeYoungChen/china-memory-sector（公开）

**本机修改后上传：**
```bash
cd ~/.agents/skills/china-memory-sector && git add -A && git commit -m "更新说明" && git push
```

**其他电脑首次安装：**

前提：已安装 git 和 python3（3.10+）。仓库是公开的，无需登录 GitHub。安装（任一步失败会立即停止）：
```bash
set -e
git clone https://github.com/LeYoungChen/china-memory-sector.git ~/.agents/skills/china-memory-sector
for a in claude codex cursor gemini; do mkdir -p ~/.$a/skills && ln -sfn ../../.agents/skills/china-memory-sector ~/.$a/skills/china-memory-sector; done
~/.agents/skills/china-memory-sector/bin/cms setup
```

**其他电脑拉取更新：**
```bash
cd ~/.agents/skills/china-memory-sector && git pull
```
（Python 依赖有新增时，下次运行 `cms` 会自动补装。）

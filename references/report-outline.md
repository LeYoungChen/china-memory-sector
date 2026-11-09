# 报告结构与 analysis.json 格式

Word 报告由 `build_report.py` 生成：表格和图表来自数据，正文来自 Claude 在第 5 步撰写的 `analysis.json`。
**全部正文用简体中文撰写**；公司用中文简称；金额单位统一为亿元。

| 报告章节 | 自动生成内容 | analysis.json 字段 |
|---|---|---|
| 封面 | 数据日期、样本数量 | `title`、`subtitle` |
| 一、核心观点 | — | `key_takeaways`（3–6 条字符串） |
| 二、周期位置判断 | 周期图（板块营收、毛利率、净利率）、分细分毛利率图、行业价格表、单季汇总表 | `cycle_view`（段落列表） |
| 三、产业结构与价值链 | 细分板块表 | `industry_structure`（段落列表） |
| 四、竞争格局 | 公司对比表 | `competitive_landscape`（段落列表）、`segment_notes`（细分 → 文字） |
| 五、估值分析 | 可比估值表、PB 分位图、PB-ROE 散点图 | `valuation_view`（段落列表） |
| 六、核心争议 | — | `debates`（[{topic, bull, bear}]） |
| 七、重点关注公司 | — | `names_to_watch`（[{name, angle, risk}]） |
| 八、风险与催化剂 | — | `risks`、`catalysts`（字符串列表） |
| 附录 | 股价表现、数据说明 | `sources`、`data_notes`（字符串列表） |

`segment_notes` 的键必须与 universe.csv 中的细分名称一致：存储设计 / 接口与主控 / 模组与存储 / 分销。
报告篇幅控制在 8–15 页。

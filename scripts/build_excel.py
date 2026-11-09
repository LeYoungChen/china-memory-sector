"""根据 dataset.json（及可选 industry.json）生成中文存储板块 Excel 底稿。

约定：蓝字 = 数据源硬输入；黑字 = 公式。金额单位：亿元。
用法: python build_excel.py --data dataset.json [--industry industry.json] --out 文件.xlsx
"""
import argparse, json, os

from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

FONT = "微软雅黑"
NAVY = "1F3864"
INPUT = Font(name=FONT, size=10, color="0000FF")
CALC = Font(name=FONT, size=10, color="000000")
BOLD = Font(name=FONT, size=10, bold=True)
HDR = Font(name=FONT, size=10, bold=True, color="FFFFFF")
HFILL = PatternFill("solid", fgColor=NAVY)
SFILL = PatternFill("solid", fgColor="D9E1F2")
THIN = Border(bottom=Side(style="thin", color="BFBFBF"))
YI = 1e8

F_YI = '#,##0.0;(#,##0.0);"-"'
F_X = '0.0"x";(0.0"x");"-"'
F_PCT = '0.0%;(0.0%);"-"'
F_PX = '#,##0.00'
SEGMENTS = ["存储设计", "接口与主控", "模组与存储", "分销"]
SH_Q = "单季财务"


def yi(v):
    return None if v is None else v / YI


def header(ws, row, labels, widths=None):
    for i, lab in enumerate(labels, 1):
        c = ws.cell(row=row, column=i, value=lab)
        c.font, c.fill = HDR, HFILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[row].height = 32
    if widths:
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[L(i)].width = w


def put(ws, r, c, v, fmt=None, inp=False, bold=False):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = BOLD if bold else (INPUT if inp else CALC)
    if fmt:
        cell.number_format = fmt
    return cell


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(name=FONT, size=14, bold=True, color=NAVY)
    ws["A2"] = sub
    ws["A2"].font = Font(name=FONT, size=9, italic=True, color="595959")


def cover(wb, d):
    ws = wb.active
    ws.title = "封面"
    title(ws, "中国存储芯片板块 —— 数据与可比估值底稿", f"数据日期 {d['as_of']} · 抓取时间 {d['fetched_at']}")
    rows = [
        ("样本", f"{len(d['companies'])} 家 A 股公司，分 4 个细分板块（见“公司池”）"),
        ("单位", "金额：亿元；估值倍数：x；利润率/收益率：%"),
        ("颜色约定", "蓝字 = 数据源硬输入 · 黑字 = 公式"),
        ("行情与估值", "东方财富（AKShare stock_value_em / stock_individual_info_em）"),
        ("股价收益", "东方财富（AKShare stock_zh_a_hist，前复权）"),
        ("财务数据", "同花顺（AKShare stock_financial_abstract_ths）"),
        ("TTM 口径", "最新累计 + 上一完整年度 − 上年同期累计"),
        ("单季口径", "单季 = 本期累计 − 上期累计；毛利 = 营收 × 披露毛利率"),
        ("抓取失败", "，".join(f"{f['code']}:{f['part']}" for f in d["failures"]) or "无"),
        ("免责声明", "仅供研究参考，不构成投资建议。对外使用前请核对公司公告。"),
    ]
    for i, (k, v) in enumerate(rows, 4):
        put(ws, i, 1, k, bold=True)
        put(ws, i, 2, v)
    ws.column_dimensions["A"].width = 14
    ws.column_dimensions["B"].width = 100
    r = len(rows) + 6
    put(ws, r, 1, "工作表目录", bold=True)
    for i, (s, desc) in enumerate([
        ("公司池", "公司、细分板块、业务简介"),
        ("可比估值", "规模、增速、利润率、PE/扣非PE/PS/PB 及统计值"),
        ("估值历史", "当前 PB、PE 相对近 3 年区间及分位；月度 PB"),
        ("股价表现", "近 1 月 / 3 月 / 年初至今 / 1 年收益（前复权）"),
        ("年度财务", "近 5 年营收、净利润、毛利率"),
        ("单季财务", "各公司单季营收、毛利、净利润（近 16 个季度）"),
        ("周期指标", "板块单季汇总：毛利率、营收同比、存货天数"),
        ("行业价格", "DRAM/NAND 合约价等行业数据（手工输入，注明来源）"),
    ], r + 1):
        put(ws, i, 1, s)
        put(ws, i, 2, desc)


def universe(wb, d):
    ws = wb.create_sheet("公司池")
    title(ws, "公司池", "修改 references/universe.csv 可调整样本")
    header(ws, 4, ["代码", "公司", "英文名", "细分板块", "上市板块", "业务简介"], [10, 12, 15, 14, 12, 80])
    for i, c in enumerate(d["companies"], 5):
        for j, k in enumerate(["code", "name_cn", "name_en", "segment", "board", "description"], 1):
            put(ws, i, j, c[k])


def comps(wb, d):
    ws = wb.create_sheet("可比估值")
    title(ws, "可比公司估值", f"股价日期 {d['as_of']} · 金额单位：亿元 · TTM 口径见封面")
    cols = ["代码", "公司", "细分板块", "股价（元）", "总股本（亿股）", "总市值", "TTM 营收",
            "TTM 归母净利", "每股净资产（元）", "年初至今营收", "上年同期营收", "年初至今净利", "上年同期净利",
            "毛利率（YTD）", "ROE（YTD）", "营收同比（YTD）", "净利同比（YTD）", "PE（TTM）", "PS（TTM）", "PB",
            "TTM 净利率", "最新报告期", "TTM 扣非净利", "扣非 PE（TTM）"]
    header(ws, 4, cols, [9, 11, 12, 10, 10, 11, 11, 11, 10, 11, 11, 11, 11, 10, 9, 10, 10, 9, 9, 8, 10, 11, 11, 10])
    first = r = 5
    for c in d["companies"]:
        v, t = c.get("valuation") or {}, c.get("ttm") or {}
        put(ws, r, 1, c["code"]); put(ws, r, 2, c["name_cn"]); put(ws, r, 3, c["segment"])
        put(ws, r, 4, v.get("close"), F_PX, inp=True)
        put(ws, r, 5, yi(v.get("shares")), '#,##0.00', inp=True)
        put(ws, r, 6, f"=D{r}*E{r}", F_YI)
        put(ws, r, 7, yi(t.get("revenue")), F_YI, inp=True)
        put(ws, r, 8, yi(t.get("net_profit")), F_YI, inp=True)
        put(ws, r, 9, t.get("bvps_latest"), F_PX, inp=True)
        put(ws, r, 10, yi(t.get("revenue_ytd")), F_YI, inp=True)
        put(ws, r, 11, yi(t.get("revenue_ytd_prior")), F_YI, inp=True)
        put(ws, r, 12, yi(t.get("net_profit_ytd")), F_YI, inp=True)
        put(ws, r, 13, yi(t.get("net_profit_ytd_prior")), F_YI, inp=True)
        gm, roe = t.get("gross_margin_ytd_pct"), t.get("roe_ytd_pct")
        put(ws, r, 14, None if gm is None else gm / 100, F_PCT, inp=True)
        put(ws, r, 15, None if roe is None else roe / 100, F_PCT, inp=True)
        put(ws, r, 16, f'=IF(N(K{r})>0,J{r}/K{r}-1,"n.m.")', F_PCT)
        put(ws, r, 17, f'=IF(N(M{r})>0,L{r}/M{r}-1,"n.m.")', F_PCT)
        put(ws, r, 18, f'=IF(N(H{r})>0,F{r}/H{r},"n.m.")', F_X)
        put(ws, r, 19, f'=IF(N(G{r})>0,F{r}/G{r},"n.m.")', F_X)
        put(ws, r, 20, f'=IF(N(I{r})>0,D{r}/I{r},"n.m.")', F_X)
        put(ws, r, 21, f'=IF(N(G{r})>0,H{r}/G{r},"n.m.")', F_PCT)
        put(ws, r, 22, t.get("latest_period"))
        put(ws, r, 23, yi(t.get("net_profit_adj")), F_YI, inp=True)
        put(ws, r, 24, f'=IF(N(W{r})>0,F{r}/W{r},"n.m.")', F_X)
        for col in range(1, 25):
            ws.cell(row=r, column=col).border = THIN
        r += 1
    last = r - 1
    r += 1
    stat_cols = {6: F_YI, 14: F_PCT, 15: F_PCT, 16: F_PCT, 17: F_PCT, 18: F_X, 19: F_X, 20: F_X, 21: F_PCT, 24: F_X}
    for lab, fn in [("最大值", "MAX({})"), ("75 分位", "QUARTILE({},3)"), ("中位数", "MEDIAN({})"),
                    ("25 分位", "QUARTILE({},1)"), ("最小值", "MIN({})")]:
        put(ws, r, 2, lab, bold=True)
        for col, fmt in stat_cols.items():
            put(ws, r, col, f'=IFERROR({fn.format(f"{L(col)}{first}:{L(col)}{last}")},"-")', fmt)
        for col in range(1, 25):
            ws.cell(row=r, column=col).fill = SFILL
        r += 1
    r += 1
    put(ws, r, 2, "细分板块均值", bold=True)
    r += 1
    for seg in SEGMENTS:
        put(ws, r, 2, seg)
        put(ws, r, 3, "家数"); put(ws, r, 4, f'=COUNTIF($C${first}:$C${last},"{seg}")')
        for col, fmt in stat_cols.items():
            rng = f"{L(col)}{first}:{L(col)}{last}"
            if col == 6:
                put(ws, r, col, f'=SUMIF($C${first}:$C${last},"{seg}",{rng})', F_YI)
            else:
                put(ws, r, col, f'=IFERROR(AVERAGEIF($C${first}:$C${last},"{seg}",{rng}),"-")', fmt)
        r += 1
    put(ws, r + 1, 2, "注：细分板块市值为合计，其余为数值项简单平均。周期顶部利润下 PE 会低估常态估值，"
        "请结合“估值历史”中的 PB 分位及“年度财务”中的 5 年平均利润判断。")
    ws.freeze_panes = "D5"


def val_history(wb, d):
    ws = wb.create_sheet("估值历史")
    title(ws, "估值相对近 3 年历史", "东方财富每日 PB、PE(TTM)；分位 = 历史上不高于当前值的交易日占比")
    cols = ["代码", "公司", "细分板块", "当前 PB", "3 年最低 PB", "3 年 PB 中位数", "3 年最高 PB",
            "PB 分位", "PB 相对中位数", "当前 PE(TTM)", "3 年 PE 中位数", "PE 分位", "PE 相对中位数"]
    header(ws, 4, cols, [9, 11, 12, 10, 10, 11, 10, 10, 11, 11, 11, 10, 11])
    r = 5
    for c in d["companies"]:
        v = c.get("valuation") or {}
        pb, pe = v.get("pb_3y") or {}, v.get("pe_3y") or {}
        put(ws, r, 1, c["code"]); put(ws, r, 2, c["name_cn"]); put(ws, r, 3, c["segment"])
        put(ws, r, 4, pb.get("current"), F_X, inp=True)
        put(ws, r, 5, pb.get("min"), F_X, inp=True)
        put(ws, r, 6, pb.get("median"), F_X, inp=True)
        put(ws, r, 7, pb.get("max"), F_X, inp=True)
        put(ws, r, 8, None if pb.get("percentile_3y") is None else pb["percentile_3y"] / 100, F_PCT, inp=True)
        put(ws, r, 9, f'=IF(N(F{r})>0,D{r}/F{r}-1,"n.m.")', F_PCT)
        put(ws, r, 10, v.get("pe_ttm_em"), F_X, inp=True)
        put(ws, r, 11, pe.get("median"), F_X, inp=True)
        put(ws, r, 12, None if pe.get("percentile_3y") is None else pe["percentile_3y"] / 100, F_PCT, inp=True)
        put(ws, r, 13, f'=IF(AND(N(J{r})>0,N(K{r})>0),J{r}/K{r}-1,"n.m.")', F_PCT)
        r += 1
    last = r - 1
    put(ws, r, 2, "中位数", bold=True)
    for col, fmt in {4: F_X, 6: F_X, 8: F_PCT, 9: F_PCT, 10: F_X, 11: F_X, 12: F_PCT, 13: F_PCT}.items():
        put(ws, r, col, f'=IFERROR(MEDIAN({L(col)}5:{L(col)}{last}),"-")', fmt)
        ws.cell(row=r, column=col).fill = SFILL
    put(ws, r + 1, 2, "PE 分位仅统计 PE 为正的交易日。PB 使用东方财富口径净资产，可能与“可比估值”PB 略有差异。")

    ch = BarChart()
    ch.type = "bar"
    ch.title = "当前 PB 在近 3 年的分位"
    ch.y_axis.number_format = "0%"
    ch.y_axis.scaling.min, ch.y_axis.scaling.max = 0, 1
    ch.add_data(Reference(ws, min_col=8, min_row=4, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=2, min_row=5, max_row=last))
    ch.height, ch.width, ch.legend = 9, 16, None
    ws.add_chart(ch, "O4")

    months = sorted({m for c in d["companies"] for m in ((c.get("valuation") or {}).get("pb_monthly") or {})})
    r0 = last + 4
    put(ws, r0 - 1, 1, "月度 PB（月末）", bold=True)
    for i, lab in enumerate(["代码", "公司"] + months, 1):
        cell = ws.cell(row=r0, column=i, value=lab)
        cell.font, cell.fill = HDR, HFILL
    for i, c in enumerate(d["companies"], r0 + 1):
        put(ws, i, 1, c["code"]); put(ws, i, 2, c["name_cn"])
        pm = (c.get("valuation") or {}).get("pb_monthly") or {}
        for j, m in enumerate(months, 3):
            put(ws, i, j, pm.get(m), '0.00', inp=True)
    ws.freeze_panes = "D5"


def performance(wb, d):
    ws = wb.create_sheet("股价表现")
    title(ws, "股价表现", "前复权收盘价；收益 = 最新 / 基期 − 1")
    cols = ["代码", "公司", "细分板块", "最新收盘", "1 月前", "3 月前", "上年末", "1 年前",
            "近 1 月", "近 3 月", "年初至今", "近 1 年"]
    header(ws, 4, cols, [9, 11, 12, 10, 10, 10, 10, 10, 9, 9, 9, 9])
    r = 5
    for c in d["companies"]:
        p = c.get("performance") or {}
        put(ws, r, 1, c["code"]); put(ws, r, 2, c["name_cn"]); put(ws, r, 3, c["segment"])
        for j, k in enumerate(["last", "1M", "3M", "YTD", "1Y"], 4):
            put(ws, r, j, (p.get(k) or {}).get("close_qfq"), F_PX, inp=True)
        for j, base in zip(range(9, 13), "EFGH"):
            put(ws, r, j, f'=IF(N({base}{r})>0,D{r}/{base}{r}-1,"n.m.")', F_PCT)
        r += 1
    last = r - 1
    for lab, fn in [("中位数", "MEDIAN"), ("平均值", "AVERAGE")]:
        put(ws, r, 2, lab, bold=True)
        for j in range(9, 13):
            put(ws, r, j, f'=IFERROR({fn}({L(j)}5:{L(j)}{last}),"-")', F_PCT)
            ws.cell(row=r, column=j).fill = SFILL
        r += 1
    dates = next((c["performance"] for c in d["companies"] if c.get("performance")), {})
    lab = {"last": "最新", "1M": "1月前", "3M": "3月前", "YTD": "上年末", "1Y": "1年前"}
    put(ws, r + 1, 2, "基期日期：" + "，".join(f"{lab[k]}={v.get('date')}" for k, v in dates.items()))


def annual(wb, d):
    ws = wb.create_sheet("年度财务")
    title(ws, "年度财务", "单位：亿元 · 会计年度截至 12 月 31 日 · 同花顺")
    years = sorted({a["period"] for c in d["companies"] if c.get("financials")
                    for a in c["financials"]["annual"] if a["period"].isdigit()})[-5:]
    n = len(d["companies"])
    blocks = [("营业总收入（亿元）", "revenue", F_YI, True), ("营收同比", None, F_PCT, False),
              ("归母净利润（亿元）", "net_profit", F_YI, True), ("净利润同比", None, F_PCT, False),
              ("销售毛利率", "gross_margin_pct", F_PCT, True), ("销售净利率", None, F_PCT, False)]
    ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 9, 12
    r, starts = 4, {}
    for name, key, fmt, is_input in blocks:
        put(ws, r, 1, name, bold=True)
        for i, lab in enumerate(["代码", "公司"] + [f"{y}年" for y in years] +
                                (["5 年平均"] if key == "net_profit" else []), 1):
            cell = ws.cell(row=r + 1, column=i, value=lab)
            cell.font, cell.fill = HDR, HFILL
        starts[name] = r + 2
        for i, c in enumerate(d["companies"]):
            rr = r + 2 + i
            put(ws, rr, 1, c["code"]); put(ws, rr, 2, c["name_cn"])
            ann = {a["period"]: a for a in (c.get("financials") or {}).get("annual", [])}
            for j, y in enumerate(years, 3):
                col = L(j)
                if is_input:
                    val = (ann.get(y) or {}).get(key)
                    if val is not None:
                        val = val / 100 if key.endswith("_pct") else val / YI
                    put(ws, rr, j, val, fmt, inp=True)
                elif name in ("营收同比", "净利润同比"):
                    src = "营业总收入（亿元）" if name == "营收同比" else "归母净利润（亿元）"
                    s = starts[src] + i
                    if j > 3:
                        p = L(j - 1)
                        put(ws, rr, j, f'=IF(N({p}{s})>0,{col}{s}/{p}{s}-1,"n.m.")', fmt)
                else:
                    s1, s2 = starts["归母净利润（亿元）"] + i, starts["营业总收入（亿元）"] + i
                    put(ws, rr, j, f'=IF(N({col}{s2})>0,{col}{s1}/{col}{s2},"n.m.")', fmt)
            if key == "net_profit":
                j = 3 + len(years)
                put(ws, rr, j, f'=IFERROR(AVERAGE({L(3)}{rr}:{L(j - 1)}{rr}),"-")', F_YI)
        r += n + 4
    for j in range(3, 4 + len(years)):
        ws.column_dimensions[L(j)].width = 11


def quarterly(wb, d):
    ws = wb.create_sheet(SH_Q)
    title(ws, "单季财务", "单位：亿元 · 由累计报表倒推（见封面）")
    cols = ["代码", "公司", "细分板块", "季度", "上年同季", "营收", "毛利", "归母净利",
            "毛利率", "净利率", "营收同比", "存货周转天数（YTD）"]
    header(ws, 4, cols, [9, 11, 12, 9, 10, 10, 10, 10, 9, 9, 9, 12])
    allq = sorted({q["quarter"] for c in d["companies"] for q in c.get("quarters", [])})
    keep = set(allq[-16:])  # 展示 12 个季度 + 4 个季度作同比基数
    r = 5
    for c in d["companies"]:
        for q in c.get("quarters", []):
            if q["quarter"] not in keep:
                continue
            put(ws, r, 1, c["code"]); put(ws, r, 2, c["name_cn"]); put(ws, r, 3, c["segment"])
            put(ws, r, 4, q["quarter"])
            put(ws, r, 5, f"{int(q['quarter'][:4]) - 1}{q['quarter'][4:]}")
            put(ws, r, 6, yi(q["revenue"]), F_YI, inp=True)
            put(ws, r, 7, yi(q["gross_profit"]), F_YI, inp=True)
            put(ws, r, 8, yi(q["net_profit"]), F_YI, inp=True)
            put(ws, r, 9, f'=IF(N(F{r})>0,G{r}/F{r},"n.m.")', F_PCT)
            put(ws, r, 10, f'=IF(N(F{r})>0,H{r}/F{r},"n.m.")', F_PCT)
            put(ws, r, 12, q["inventory_days_ytd"], '0', inp=True)
            r += 1
    last = r - 1
    for i in range(5, r):
        put(ws, i, 11, f'=IFERROR(F{i}/SUMIFS($F$5:$F${last},$A$5:$A${last},A{i},$D$5:$D${last},E{i})-1,"n.m.")', F_PCT)
    ws.freeze_panes = "E5"
    ws.auto_filter.ref = f"A4:L{last}"
    return sorted(keep), last


def cycle(wb, quarters, qlast):
    ws = wb.create_sheet("周期指标")
    title(ws, "板块周期指标（样本合计）", "基于“单季财务”的 SUMIFS 汇总。注意：新上市公司在有数据后才进入样本。")
    cols = ["季度", "营收", "毛利", "归母净利", "板块毛利率", "板块净利率", "营收同比",
            "平均存货天数（YTD）", "家数"] + [f"毛利率：{s}" for s in SEGMENTS]
    header(ws, 4, cols, [10, 11, 11, 11, 10, 10, 10, 13, 7, 13, 13, 13, 11])
    col = lambda c: f"'{SH_Q}'!${c}$5:${c}${qlast}"
    for i, q in enumerate(quarters, 5):
        put(ws, i, 1, q)
        put(ws, i, 2, f'=SUMIFS({col("F")},{col("D")},$A{i})', F_YI)
        put(ws, i, 3, f'=SUMIFS({col("G")},{col("D")},$A{i})', F_YI)
        put(ws, i, 4, f'=SUMIFS({col("H")},{col("D")},$A{i})', F_YI)
        put(ws, i, 5, f'=IF(B{i}>0,C{i}/B{i},"n.m.")', F_PCT)
        put(ws, i, 6, f'=IF(B{i}>0,D{i}/B{i},"n.m.")', F_PCT)
        put(ws, i, 7, f'=IF(N(B{i - 4})>0,B{i}/B{i - 4}-1,"n.m.")' if i >= 9 else "n.m.", F_PCT)
        put(ws, i, 8, f'=IFERROR(AVERAGEIFS({col("L")},{col("D")},$A{i}),"n.m.")', '0')
        put(ws, i, 9, f'=COUNTIFS({col("D")},$A{i})')
        for j, seg in enumerate(SEGMENTS, 10):
            put(ws, i, j, f'=IFERROR(SUMIFS({col("G")},{col("D")},$A{i},{col("C")},"{seg}")/'
                          f'SUMIFS({col("F")},{col("D")},$A{i},{col("C")},"{seg}"),"n.m.")', F_PCT)
    last = 4 + len(quarters)
    ch = LineChart()
    ch.title = "分细分板块单季毛利率"
    ch.y_axis.number_format = "0%"
    ch.add_data(Reference(ws, min_col=5, min_row=4, max_row=last), titles_from_data=True)
    for j in range(10, 14):
        ch.add_data(Reference(ws, min_col=j, min_row=4, max_row=last), titles_from_data=True)
    ch.set_categories(Reference(ws, min_col=1, min_row=5, max_row=last))
    ch.height, ch.width = 9, 20
    ws.add_chart(ch, f"A{last + 3}")
    ch2 = BarChart()
    ch2.title = "板块单季营收（亿元）"
    ch2.add_data(Reference(ws, min_col=2, min_row=4, max_row=last), titles_from_data=True)
    ch2.set_categories(Reference(ws, min_col=1, min_row=5, max_row=last))
    ch2.height, ch2.width, ch2.legend = 9, 16, None
    ws.add_chart(ch2, f"H{last + 3}")


def industry(wb, ind):
    ws = wb.create_sheet("行业价格")
    title(ws, "存储合约价及行业数据", "手工输入，来源：TrendForce 集邦 / DRAMeXchange / CFM 闪存市场等。每行须注明来源和日期。")
    header(ws, 4, ["产品", "指标", "期间", "数值", "来源", "日期", "链接"], [26, 14, 10, 12, 22, 12, 60])
    rows = (ind or {}).get("prices", [])
    r = 5
    for p in rows:
        for j, k in enumerate(["product", "metric", "period", "value", "source", "date", "url"], 1):
            put(ws, r, j, p.get(k), inp=True)
        r += 1
    for _ in range(max(0, 12 - len(rows))):
        for j in range(1, 8):
            ws.cell(row=r, column=j).border = THIN
        r += 1
    notes = (ind or {}).get("notes", [])
    if notes:
        put(ws, r + 1, 1, "背景说明", bold=True)
        for i, n in enumerate(notes, r + 2):
            put(ws, i, 1, n)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--industry")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    d = json.load(open(a.data, encoding="utf-8"))
    ind = json.load(open(a.industry, encoding="utf-8")) if a.industry and os.path.exists(a.industry) else None
    wb = Workbook()
    cover(wb, d)
    universe(wb, d)
    comps(wb, d)
    val_history(wb, d)
    performance(wb, d)
    annual(wb, d)
    qs, qlast = quarterly(wb, d)
    cycle(wb, qs, qlast)
    industry(wb, ind)
    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.save(a.out)
    print("已生成", a.out)


if __name__ == "__main__":
    main()

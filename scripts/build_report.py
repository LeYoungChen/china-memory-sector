"""生成中文 Word 研究报告：表格/图表来自 dataset.json，正文来自 analysis.json。
用法: python build_report.py --data dataset.json --analysis analysis.json [--industry industry.json] --out 报告.docx
"""
import argparse, json, os, sys, tempfile

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.path.insert(0, os.path.dirname(__file__))
import metrics as M

NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x59, 0x59, 0x59)
PALETTE = ["#1F3864", "#2E75B6", "#9DC3E6", "#C55A11", "#7F7F7F"]
SEG_COLOR = dict(zip(M.SEGMENTS, PALETTE))
DOC_FONT_CN = "微软雅黑"

CJK_CANDIDATES = ["PingFang SC", "Hiragino Sans GB", "Heiti SC", "STHeiti", "Arial Unicode MS",
                  "Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Zen Hei"]
_avail = {f.name for f in font_manager.fontManager.ttflist}
CJK = [f for f in CJK_CANDIDATES if f in _avail]
if not CJK:
    print("[警告] 未找到中文字体，图表中文可能显示为方框", file=sys.stderr)
plt.rcParams.update({"font.family": "sans-serif", "font.sans-serif": CJK + ["DejaVu Sans"],
                     "axes.unicode_minus": False, "font.size": 9, "axes.spines.top": False,
                     "axes.spines.right": False, "axes.edgecolor": "#7F7F7F"})

CN_NUM = "一二三四五六七八九十"


def fx(x, kind="x"):
    if x is None:
        return "n.m."
    return {"x": f"{x:.1f}x", "p": f"{x * 100:+.0f}%", "m": f"{x * 100:.1f}%", "n": f"{x:,.1f}"}[kind]


def set_cn_font(run, name=DOC_FONT_CN):
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.append(fonts)
    fonts.set(qn("w:eastAsia"), name)


def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear"); sh.set(qn("w:color"), "auto"); sh.set(qn("w:fill"), hex_)
    tcPr.append(sh)


def note(doc, text):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.italic, r.font.size, r.font.color.rgb = True, Pt(7.5), GREY
    set_cn_font(r)


def table(doc, head, rows, widths=None, foot=None):
    t = doc.add_table(rows=1, cols=len(head))
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.style = "Table Grid"
    for i, h in enumerate(head):
        c = t.rows[0].cells[i]
        c.text = ""
        run = c.paragraphs[0].add_run(h)
        run.bold, run.font.size, run.font.color.rgb = True, Pt(8), RGBColor(0xFF, 0xFF, 0xFF)
        set_cn_font(run)
        shade(c, "1F3864")
    for k, r in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            run = cells[i].paragraphs[0].add_run(str(v))
            run.font.size = Pt(8)
            set_cn_font(run)
            if i > 1:
                cells[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if k % 2:
                shade(cells[i], "F2F2F2")
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    if foot:
        note(doc, foot)
    return t


def para(doc, text, style=None):
    p = doc.add_paragraph(style=style)
    run = p.add_run(text)
    set_cn_font(run)
    return p


def paras(doc, items):
    for s in items or []:
        para(doc, s)


def bullets(doc, items):
    for s in items or []:
        para(doc, s, "List Bullet")


def heading(doc, text, level):
    h = doc.add_heading(level=level)
    run = h.add_run(text)
    run.font.color.rgb = NAVY
    set_cn_font(run)
    return h


def fig(doc, path, caption, width=15.5):
    doc.add_picture(path, width=Cm(width))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    r = p.add_run(caption)
    r.italic, r.font.size = True, Pt(8)
    set_cn_font(r)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER


# ---------- 图表 ----------
def save(f, path):
    f.tight_layout(); f.savefig(path, dpi=200); plt.close(f)


def chart_cycle(q, path):
    f, ax = plt.subplots(figsize=(8, 3.4))
    xs = [r["quarter"] for r in q]
    ax.bar(xs, [r["rev"] for r in q], color="#9DC3E6", label="板块单季营收（亿元）")
    ax.set_ylabel("亿元")
    ax2 = ax.twinx()
    ax2.plot(xs, [r["gm"] * 100 if r["gm"] else None for r in q], color="#1F3864", marker="o", label="板块毛利率（%）")
    ax2.plot(xs, [r["nm"] * 100 if r["nm"] else None for r in q], color="#C55A11", marker="s", label="板块净利率（%）")
    ax2.set_ylabel("%"); ax2.spines["right"].set_visible(True)
    h1, l1 = ax.get_legend_handles_labels(); h2, l2 = ax2.get_legend_handles_labels()
    ax.legend(h1 + h2, l1 + l2, loc="upper left", frameon=False, fontsize=8)
    ax.tick_params(axis="x", rotation=45)
    save(f, path)


def chart_seg_gm(q, path):
    f, ax = plt.subplots(figsize=(8, 3.2))
    xs = [r["quarter"] for r in q]
    for s in M.SEGMENTS:
        ax.plot(xs, [r["seg_gm"].get(s) * 100 if r["seg_gm"].get(s) is not None else None for r in q],
                marker="o", ms=3, color=SEG_COLOR[s], label=s)
    ax.set_ylabel("毛利率（%）"); ax.legend(frameon=False, fontsize=8, ncol=4)
    ax.tick_params(axis="x", rotation=45)
    save(f, path)


def chart_pb_pct(rows, path):
    rs = sorted([r for r in rows if r["pb_pct"] is not None], key=lambda r: r["pb_pct"])
    f, ax = plt.subplots(figsize=(8, 3.6))
    ax.barh([r["name"] for r in rs], [r["pb_pct"] * 100 for r in rs], color=[SEG_COLOR[r["segment"]] for r in rs])
    for i, r in enumerate(rs):
        ax.text(r["pb_pct"] * 100 + 1, i, f"{r['pb_em']:.1f}x", va="center", fontsize=7)
    ax.set_xlim(0, 110); ax.set_xlabel("当前 PB 在近 3 年日度数据中的分位（%）")
    ax.axvline(50, color="#7F7F7F", ls="--", lw=0.8)
    save(f, path)


def chart_pb_roe(rows, path):
    f, ax = plt.subplots(figsize=(8, 3.8))
    for r in rows:
        if r["pb"] is None or r["roe"] is None:
            continue
        ax.scatter(r["roe"] * 100, r["pb"], s=max(20, (r["mcap"] or 0) / 15), alpha=0.75,
                   color=SEG_COLOR[r["segment"]], label=r["segment"])
        ax.annotate(r["name"], (r["roe"] * 100, r["pb"]), fontsize=7, xytext=(4, 3), textcoords="offset points")
    h, l = ax.get_legend_handles_labels()
    uniq = dict(zip(l, h))
    ax.legend(uniq.values(), uniq.keys(), frameon=False, fontsize=8)
    ax.set_xlabel("ROE（最新年初至今，%，未年化）"); ax.set_ylabel("PB（x）")
    save(f, path)


def chart_perf(p, path):
    rs = sorted([r for r in p if r["YTD"] is not None], key=lambda r: r["YTD"])
    f, ax = plt.subplots(figsize=(8, 3.4))
    ax.barh([r["name"] for r in rs], [r["YTD"] * 100 for r in rs], color=[SEG_COLOR[r["segment"]] for r in rs])
    ax.set_xlabel("年初至今股价涨跌幅（%，前复权）")
    save(f, path)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    ap.add_argument("--analysis", required=True)
    ap.add_argument("--industry")
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    d = json.load(open(a.data, encoding="utf-8"))
    an = json.load(open(a.analysis, encoding="utf-8"))
    ind = json.load(open(a.industry, encoding="utf-8")) if a.industry and os.path.exists(a.industry) else {}
    rows, q, p = M.comps(d), M.quarters_agg(d), M.perf(d)
    tmp = tempfile.mkdtemp()
    ch = {k: os.path.join(tmp, k + ".png") for k in ("cycle", "seg", "pbpct", "pbroe", "perf")}
    chart_cycle(q, ch["cycle"]); chart_seg_gm(q, ch["seg"]); chart_pb_pct(rows, ch["pbpct"])
    chart_pb_roe(rows, ch["pbroe"]); chart_perf(p, ch["perf"])

    doc = Document()
    st = doc.styles["Normal"]
    st.font.name, st.font.size = "Arial", Pt(10)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), DOC_FONT_CN)
    for sec in doc.sections:
        sec.left_margin = sec.right_margin = Cm(2)
        sec.top_margin = sec.bottom_margin = Cm(1.8)

    t = doc.add_heading(level=0)
    tr = t.add_run(an.get("title", "中国存储芯片板块"))
    tr.font.color.rgb = NAVY
    set_cn_font(tr)
    sub = doc.add_paragraph()
    r = sub.add_run(f"{an.get('subtitle', '')}\n数据日期 {d['as_of']} · 样本 {len(d['companies'])} 家 A 股公司 · "
                    "数据来源：东方财富、同花顺（AKShare）")
    r.font.size, r.font.color.rgb = Pt(9), GREY
    set_cn_font(r)

    sec = iter(CN_NUM)
    heading(doc, f"{next(sec)}、核心观点", 1)
    bullets(doc, an.get("key_takeaways"))

    heading(doc, f"{next(sec)}、周期位置判断", 1)
    paras(doc, an.get("cycle_view"))
    fig(doc, ch["cycle"], "图 1：样本合计单季营收与利润率")
    fig(doc, ch["seg"], "图 2：分细分板块单季毛利率")
    if ind.get("prices"):
        table(doc, ["产品", "指标", "期间", "数值", "来源", "日期"],
              [[x.get("product"), x.get("metric"), x.get("period"), x.get("value"), x.get("source"), x.get("date")]
               for x in ind["prices"]], foot="行业数据为第三方机构估计，详见附录来源。")
    table(doc, ["季度", "营收（亿元）", "同比", "毛利率", "净利率", "平均存货天数（YTD）", "家数"],
          [[x["quarter"], fx(x["rev"], "n"), fx(x["rev_yoy"], "p"), fx(x["gm"], "m"), fx(x["nm"], "m"),
            f"{x['inv_avg']:.0f}" if x["inv_avg"] else "n.m.", x["n"]] for x in q[-8:]],
          foot="单季数据由累计报表倒推；新上市公司进入样本会影响可比性。")

    heading(doc, f"{next(sec)}、产业结构与价值链", 1)
    paras(doc, an.get("industry_structure"))
    seg_rows = [[s, "、".join(r["name"] for r in rows if r["segment"] == s),
                 fx(sum(r["mcap"] or 0 for r in rows if r["segment"] == s), "n")] for s in M.SEGMENTS]
    table(doc, ["细分板块", "公司", "总市值（亿元）"], seg_rows, widths=[3, 11, 3])

    heading(doc, f"{next(sec)}、竞争格局", 1)
    paras(doc, an.get("competitive_landscape"))
    table(doc, ["公司", "细分板块", "TTM 营收", "营收同比（YTD）", "毛利率（YTD）", "TTM 归母净利", "总市值"],
          [[r["name"], r["segment"], fx(r["ttm_rev"], "n"), fx(r["rev_yoy"], "p"), fx(r["gm"], "m"),
            fx(r["ttm_np"], "n"), fx(r["mcap"], "n")] for r in sorted(rows, key=lambda r: -(r["mcap"] or 0))],
          foot=f"金额单位：亿元。TTM 与 YTD 截至最新报告期（{rows[0]['latest_period']}）。来源：同花顺、东方财富。")
    for seg, txt in (an.get("segment_notes") or {}).items():
        pp = doc.add_paragraph()
        b = pp.add_run(seg + "："); b.bold = True; set_cn_font(b)
        set_cn_font(pp.add_run(txt))

    heading(doc, f"{next(sec)}、估值分析", 1)
    paras(doc, an.get("valuation_view"))
    comp_rows = [[r["name"], fx(r["mcap"], "n"), fx(r["pe"]), fx(r["pe_adj"]), fx(r["ps"]), fx(r["pb"]),
                  fx(r["pb_pct"], "m"), fx(r["roe"], "m"), fx(r["np_yoy"], "p")]
                 for r in sorted(rows, key=lambda r: -(r["mcap"] or 0))]
    med = lambda k, kind="x": fx(M.stats(rows, k).get("median"), kind)
    comp_rows.append(["中位数", "", med("pe"), med("pe_adj"), med("ps"), med("pb"), med("pb_pct", "m"),
                      med("roe", "m"), med("np_yoy", "p")])
    table(doc, ["公司", "总市值（亿元）", "PE（TTM）", "扣非 PE", "PS（TTM）", "PB", "PB 3 年分位", "ROE（YTD）",
                "净利同比（YTD）"], comp_rows,
          foot="扣非 PE 使用扣除非经常性损益后的净利润。PB = 股价 / 最新每股净资产；分位基于东方财富日度 PB。ROE 未年化。")
    fig(doc, ch["pbpct"], "图 3：当前 PB 在近 3 年区间中的分位（标签为东方财富口径当前 PB，可能与估值表中“股价/最新每股净资产”口径不同）")
    fig(doc, ch["pbroe"], "图 4：PB 与 ROE（气泡大小 = 总市值）")

    heading(doc, f"{next(sec)}、核心争议", 1)
    if an.get("debates"):
        table(doc, ["争议点", "多方观点", "空方观点"],
              [[x["topic"], x["bull"], x["bear"]] for x in an["debates"]], widths=[3.5, 6.5, 6.5])

    heading(doc, f"{next(sec)}、重点关注公司", 1)
    if an.get("names_to_watch"):
        table(doc, ["公司", "关注逻辑", "主要风险"],
              [[x["name"], x["angle"], x["risk"]] for x in an["names_to_watch"]], widths=[2.5, 8.5, 5.5])

    heading(doc, f"{next(sec)}、风险与催化剂", 1)
    heading(doc, "风险", 2)
    bullets(doc, an.get("risks"))
    heading(doc, "催化剂", 2)
    bullets(doc, an.get("catalysts"))

    heading(doc, "附录一：股价表现", 1)
    fig(doc, ch["perf"], "图 5：年初至今股价涨跌幅")
    table(doc, ["公司", "近 1 月", "近 3 月", "年初至今", "近 1 年"],
          [[r["name"]] + [fx(r[k], "p") for k in ("1M", "3M", "YTD", "1Y")] for r in p])
    heading(doc, "附录二：数据来源与说明", 1)
    bullets(doc, (an.get("sources") or []) + (an.get("data_notes") or []))
    if d["failures"]:
        para(doc, "数据抓取失败：" + "；".join(f"{f['code']} {f['part']}" for f in d["failures"]))
    note(doc, "本报告仅供研究参考，不构成投资建议。数据为时点快照，对外使用前请与公司公告核对。")
    doc.save(a.out)
    print("已生成", a.out)


if __name__ == "__main__":
    main()

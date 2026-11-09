"""打印撰写 analysis.json 所需的关键数据（中文）。
用法: python summarize.py --data dataset.json
"""
import argparse, json, os, sys

sys.path.insert(0, os.path.dirname(__file__))
import metrics as M


def f(x, kind="x"):
    if x is None:
        return "n.m."
    return {"x": f"{x:.1f}x", "p": f"{x * 100:+.0f}%", "m": f"{x * 100:.1f}%", "n": f"{x:,.1f}"}[kind]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--data", required=True)
    d = json.load(open(ap.parse_args().data, encoding="utf-8"))
    print(f"数据日期 {d['as_of']}  抓取失败: {d['failures'] or '无'}\n")

    print("== 周期指标（样本合计，单季，亿元）")
    for q in M.quarters_agg(d):
        seg = "  ".join(f"{s}={f(v, 'm')}" for s, v in q["seg_gm"].items())
        print(f"{q['quarter']}: 营收 {f(q['rev'], 'n'):>7}  同比 {f(q['rev_yoy'], 'p'):>6}  毛利率 {f(q['gm'], 'm'):>6}  "
              f"净利率 {f(q['nm'], 'm'):>6}  存货天数 {q['inv_avg'] and round(q['inv_avg'])}  家数={q['n']} | {seg}")

    rows = M.comps(d)
    by_code = {c["code"]: c for c in d["companies"]}
    print("\n== 可比估值（市值、利润单位：亿元；增速为年初至今同比）")
    print("公司      细分        市值    营收同比  净利同比  毛利率   PE   扣非PE   PS    PB   PB分位  TTM净利  近5年均净利  最新报告期")
    for r in sorted(rows, key=lambda r: -(r["mcap"] or 0)):
        avg = M.annual_avg_np(by_code[r["code"]])
        print(f"{r['name']:6}{r['segment']:6}{f(r['mcap'], 'n'):>9}{f(r['rev_yoy'], 'p'):>8}{f(r['np_yoy'], 'p'):>8}"
              f"{f(r['gm'], 'm'):>7}{f(r['pe']):>7}{f(r['pe_adj']):>8}{f(r['ps']):>7}{f(r['pb']):>6}{f(r['pb_pct'], 'm'):>8}"
              f"{f(r['ttm_np'], 'n'):>9}{f(avg, 'n'):>10}  {r['latest_period']}")
    names = {"pe": "PE", "pe_adj": "扣非PE", "ps": "PS", "pb": "PB", "pb_pct": "PB分位", "gm": "毛利率", "rev_yoy": "营收同比"}
    for k, lab in names.items():
        s = M.stats(rows, k)
        print(f"  {lab:6} 中位数 {s.get('median'):.3f}  最小 {s.get('min'):.3f}  最大 {s.get('max'):.3f}")
    print(f"  样本总市值 {sum(r['mcap'] or 0 for r in rows):,.0f} 亿元")

    print("\n== 每股净资产变化（核查增发/其他综合收益）")
    for c in d["companies"]:
        ps = {p["period"]: p for p in c["financials"]["periods"]} if c.get("financials") else {}
        fy = max((k for k in ps if k.endswith("12-31")), default=None)
        if fy and c.get("ttm"):
            print(f"  {c['name_cn']}: {fy} {ps[fy]['bvps']} → {c['ttm']['latest_period']} {c['ttm']['bvps_latest']}")

    print("\n== 股价表现（前复权）")
    for r in M.perf(d):
        print(f"{r['name']:6} " + "  ".join(f"{k} {f(r[k], 'p')}" for k in ("1M", "3M", "YTD", "1Y")))

    gpath = os.path.join(os.path.dirname(os.path.abspath(ap.parse_args().data)), "global.json")
    if not os.path.exists(gpath):
        print("\n（未找到 global.json，跳过海外对标）")
        return
    g = json.load(open(gpath, encoding="utf-8"))
    grows = M.global_comps(g)
    print(f"\n== 海外可比（市值：亿元人民币；汇率 {g['fx_to_cny']}；失败 {g['failures'] or '无'}）")
    print("公司        细分      市场     市值     PE(口径)      远期PE   PB   EV/EBITDA  最新单季毛利率  上季   最新季营收同比  对标A股")
    for r in sorted(grows, key=lambda r: -(r["mcap"] or 0)):
        print(f"{r['name']:8}{r['segment']:6}{r['market']:6}{f(r['mcap'], 'n'):>10}  {f(r['pe']):>6}({r['pe_src']})"
              f"{f(r['fpe']):>7}{f(r['pb']):>6}{f(r['ev_ebitda']):>8}   {r['gm_lastq_label']} {f(r['gm_lastq'], 'm'):>6}"
              f"{f(r['gm_prevq'], 'm'):>7}{f(r['rev_yoy_lastq'], 'p'):>9}   {r['a_share_peers']}")
        for w in r["warnings"]:
            print("     ⚠", w)
    print("\n== 中外对比（中位数）")
    for p in M.peer_compare(rows, grows):
        print(f"A股 {p['a_seg']}({p['a_n']}) vs 海外 {p['g_seg']}({p['g_n']}): PE {f(p['a_pe'])} vs {f(p['g_pe'])} "
              f"(海外远期 {f(p['g_fpe'])}) | PB {f(p['a_pb'])} vs {f(p['g_pb'])} | 毛利率 {f(p['a_gm'], 'm')} vs {f(p['g_gm'], 'm')}")
    print("\n== 海外单季毛利率")
    by = {}
    for q in M.global_quarters(g):
        by.setdefault(q["name"], []).append(f"{q['quarter']}:{f(q['gm'], 'm')}")
    for k, v in by.items():
        print(f"  {k:8} " + "  ".join(v))
    print("\n== 海外股价表现（复权）")
    for r in M.global_perf(g):
        print(f"{r['name']:8} " + "  ".join(f"{k} {f(r[k], 'p')}" for k in ("1M", "3M", "YTD", "1Y")))


if __name__ == "__main__":
    main()

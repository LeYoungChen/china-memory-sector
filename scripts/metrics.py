"""共享指标计算（与 Excel 公式口径一致），供 summarize.py 和 build_report.py 使用。金额单位：亿元。"""
import statistics as st

YI = 1e8
SEGMENTS = ["存储设计", "接口与主控", "模组与存储", "分销"]


def _div(a, b):
    return a / b if a is not None and b is not None and b > 0 else None


def _yi(v):
    return None if v is None else v / YI


def _growth(cur, prior):
    r = _div(cur, prior)
    return None if r is None else r - 1


def comps(d):
    rows = []
    for c in d["companies"]:
        v, t = c.get("valuation") or {}, c.get("ttm") or {}
        price, shares = v.get("close"), v.get("shares")
        mcap = price * shares / YI if price and shares else None
        rev, npf, npa = _yi(t.get("revenue")), _yi(t.get("net_profit")), _yi(t.get("net_profit_adj"))
        pb3 = v.get("pb_3y") or {}
        rows.append({
            "code": c["code"], "name": c["name_cn"], "name_en": c["name_en"], "segment": c["segment"],
            "price": price, "mcap": mcap, "ttm_rev": rev, "ttm_np": npf, "ttm_np_adj": npa,
            "rev_yoy": _growth(t.get("revenue_ytd"), t.get("revenue_ytd_prior")),
            "np_yoy": _growth(t.get("net_profit_ytd"), t.get("net_profit_ytd_prior")),
            "gm": None if t.get("gross_margin_ytd_pct") is None else t["gross_margin_ytd_pct"] / 100,
            "roe": None if t.get("roe_ytd_pct") is None else t["roe_ytd_pct"] / 100,
            "pe": _div(mcap, npf), "pe_adj": _div(mcap, npa), "ps": _div(mcap, rev),
            "pb": _div(price, t.get("bvps_latest")), "net_margin": _div(npf, rev),
            "pb_pct": None if pb3.get("percentile_3y") is None else pb3["percentile_3y"] / 100,
            "pb_med3y": pb3.get("median"), "pb_em": v.get("pb_em"),
            "latest_period": t.get("latest_period"),
        })
    return rows


def stats(rows, key):
    xs = sorted(r[key] for r in rows if r.get(key) is not None)
    return {"min": xs[0], "median": st.median(xs), "max": xs[-1], "n": len(xs)} if xs else {}


def quarters_agg(d, n=12):
    allq = sorted({q["quarter"] for c in d["companies"] for q in c.get("quarters", [])})
    view = allq[-n:]
    out = []
    for q in allq[-(n + 4):]:
        rec = {"quarter": q, "rev": 0.0, "gp": 0.0, "np": 0.0, "inv": [], "n": 0,
               "seg": {s: [0.0, 0.0] for s in SEGMENTS}}
        for c in d["companies"]:
            for x in c.get("quarters", []):
                if x["quarter"] == q and x["revenue"] is not None:
                    rec["rev"] += x["revenue"] / YI
                    rec["gp"] += (x["gross_profit"] or 0) / YI
                    rec["np"] += (x["net_profit"] or 0) / YI
                    rec["n"] += 1
                    if x["inventory_days_ytd"] is not None:
                        rec["inv"].append(x["inventory_days_ytd"])
                    s = rec["seg"].setdefault(c["segment"], [0.0, 0.0])
                    s[0] += x["revenue"] / YI
                    s[1] += (x["gross_profit"] or 0) / YI
        rec["gm"] = rec["gp"] / rec["rev"] if rec["rev"] else None
        rec["nm"] = rec["np"] / rec["rev"] if rec["rev"] else None
        rec["inv_avg"] = st.mean(rec["inv"]) if rec["inv"] else None
        rec["seg_gm"] = {s: (v[1] / v[0] if v[0] else None) for s, v in rec["seg"].items()}
        out.append(rec)
    by = {r["quarter"]: r for r in out}
    for r in out:
        prev = by.get(f"{int(r['quarter'][:4]) - 1}{r['quarter'][4:]}")
        r["rev_yoy"] = r["rev"] / prev["rev"] - 1 if prev and prev["rev"] else None
    return [r for r in out if r["quarter"] in view]


def annual_avg_np(c, years=5):
    ann = [a for a in (c.get("financials") or {}).get("annual", []) if a["period"].isdigit()]
    vals = [a["net_profit"] for a in ann[-years:] if a["net_profit"] is not None]
    return sum(vals) / len(vals) / YI if vals else None


GLOBAL_SEGMENTS = ["存储原厂", "利基存储", "接口与主控", "模组与存储"]
# A 股细分 → 最可比的海外细分
PEER_MAP = {"存储设计": "利基存储", "接口与主控": "接口与主控", "模组与存储": "模组与存储", "分销": "存储原厂"}


def global_comps(g):
    """海外可比：市值折人民币亿元；PE 优先用连续 4 季 TTM 计算，否则回退 Yahoo trailingPE。"""
    fx = g["fx_to_cny"]
    rows = []
    for c in g["companies"]:
        i = c.get("info") or {}
        cur, fcur = i.get("currency"), i.get("financialCurrency") or i.get("currency")
        fx_t, fx_f = fx.get(cur), fx.get(fcur)
        mcap_local = i.get("marketCap")
        mcap = mcap_local * fx_t / YI if mcap_local and fx_t else None
        t = c.get("ttm_q")
        np_ttm = t["net_profit"] * fx_f / YI if t and fx_f else None
        rev_ttm = t["revenue"] * fx_f / YI if t and fx_f else None
        pe, pe_src = (_div(mcap, np_ttm), "季度合计") if t else (i.get("trailingPE"), "Yahoo")
        qs = c.get("quarters") or []
        lastq = qs[-1] if qs else {}
        by = {q["quarter"]: q for q in qs}
        prev = by.get(f"{int(lastq['quarter'][:4]) - 1}{lastq['quarter'][4:]}") if lastq else None
        gmq = [q for q in qs if q.get("gross_profit") is not None]
        rows.append({
            "ticker": c["ticker"], "name": c["name_cn"], "name_en": c["name_en"], "segment": c["segment"],
            "market": c["market"], "currency": cur, "fx": fx_t, "mcap": mcap, "ttm_rev": rev_ttm, "ttm_np": np_ttm,
            "ttm_through": t["through"] if t else None, "pe": pe, "pe_src": pe_src,
            "fpe": i.get("forwardPE"), "pb": i.get("priceToBook"), "ev_ebitda": i.get("enterpriseToEbitda"),
            "gm_lastq": gmq[-1]["gross_profit"] / gmq[-1]["revenue"] if gmq else None,
            "gm_lastq_label": gmq[-1]["quarter"] if gmq else None,
            "gm_prevq": gmq[-2]["gross_profit"] / gmq[-2]["revenue"] if len(gmq) > 1 else None,
            "rev_yoy_lastq": _growth(lastq.get("revenue"), prev["revenue"]) if prev else None,
            "last_quarter": lastq.get("quarter"), "a_share_peers": c.get("a_share_peers") or "",
            "warnings": [f"{q['quarter']}：{q['warning']}" for q in qs if q.get("warning")],
        })
    return rows


def global_quarters(g):
    out = []
    for c in g["companies"]:
        for q in c.get("quarters") or []:
            out.append({"name": c["name_cn"], "segment": c["segment"], "quarter": q["quarter"],
                        "period_end": q["period_end"],
                        "gm": q["gross_profit"] / q["revenue"] if q.get("gross_profit") is not None else None,
                        "nm": q["net_profit"] / q["revenue"] if q.get("net_profit") is not None else None})
    return out


def global_perf(g):
    rows = []
    for c in g["companies"]:
        p = c.get("performance") or {}
        last = (p.get("last") or {}).get("close_adj")
        r = {"name": c["name_cn"], "segment": c["segment"]}
        for k in ("1M", "3M", "YTD", "1Y"):
            base = (p.get(k) or {}).get("close_adj")
            r[k] = last / base - 1 if last and base else None
        rows.append(r)
    return rows


def peer_compare(a_rows, g_rows):
    """A 股细分 vs 海外可比细分的中位数对比。"""
    out = []
    for a_seg, g_seg in PEER_MAP.items():
        a = [r for r in a_rows if r["segment"] == a_seg]
        b = [r for r in g_rows if r["segment"] == g_seg]
        out.append({"a_seg": a_seg, "g_seg": g_seg, "a_n": len(a), "g_n": len(b),
                    "a_pe": stats(a, "pe").get("median"), "g_pe": stats(b, "pe").get("median"),
                    "a_pb": stats(a, "pb").get("median"), "g_pb": stats(b, "pb").get("median"),
                    "a_gm": stats(a, "gm").get("median"), "g_gm": stats(b, "gm_lastq").get("median"),
                    "g_fpe": stats(b, "fpe").get("median")})
    return out


def perf(d):
    rows = []
    for c in d["companies"]:
        p = c.get("performance") or {}
        last = (p.get("last") or {}).get("close_qfq")
        r = {"name": c["name_cn"], "segment": c["segment"]}
        for k in ("1M", "3M", "YTD", "1Y"):
            base = (p.get(k) or {}).get("close_qfq")
            r[k] = last / base - 1 if last and base else None
        rows.append(r)
    return rows

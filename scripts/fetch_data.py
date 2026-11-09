"""Fetch prices, valuation history and financials for the China memory universe.

Sources (via AKShare):
  - stock_individual_info_em   East Money: price, shares, market cap
  - stock_value_em             East Money: daily PE(TTM), PB, PS history
  - stock_zh_a_hist            East Money: daily forward-adjusted (qfq) prices
  - stock_financial_abstract_ths  THS (同花顺): annual & cumulative quarterly financials

Usage: python fetch_data.py --universe universe.csv --out data/ [--asof 2026-11-09]
"""
import argparse, csv, datetime as dt, json, math, os, re, sys, time, traceback

import akshare as ak
import pandas as pd

AMOUNT_RE = re.compile(r"^\s*(-?[\d.]+)\s*(万亿|亿|万)?\s*$")


def num(v):
    """Parse THS values like '110.66亿', '2740.06万', '66.22%', '-', False -> float (RMB / %)."""
    if v is None or v is False:
        return None
    if isinstance(v, (int, float)):
        return None if (isinstance(v, float) and math.isnan(v)) else float(v)
    s = str(v).strip().replace(",", "")
    if s in ("", "-", "--", "False", "nan"):
        return None
    if s.endswith("%"):
        try:
            return float(s[:-1])
        except ValueError:
            return None
    m = AMOUNT_RE.match(s)
    if not m:
        return None
    x = float(m.group(1))
    unit = m.group(2)
    return x * {"万亿": 1e12, "亿": 1e8, "万": 1e4, None: 1}[unit]


def retry(fn, tries=3, wait=2):
    for i in range(tries):
        try:
            return fn()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(wait * (i + 1))


FIN_FIELDS = {
    "营业总收入": "revenue",
    "净利润": "net_profit",
    "扣非净利润": "net_profit_adj",
    "销售毛利率": "gross_margin_pct",
    "销售净利率": "net_margin_pct",
    "净资产收益率": "roe_pct",
    "每股净资产": "bvps",
    "存货周转天数": "inventory_days",
    "资产负债率": "debt_to_assets_pct",
}


def financials(code):
    out = {}
    for key, ind in (("annual", "按年度"), ("periods", "按报告期")):
        df = retry(lambda: ak.stock_financial_abstract_ths(symbol=code, indicator=ind))
        rows = []
        for _, r in df.iterrows():
            row = {"period": str(r["报告期"])}
            for cn, en in FIN_FIELDS.items():
                row[en] = num(r.get(cn))
            rows.append(row)
        out[key] = rows
    return out


def derive_quarters(periods):
    """Cumulative YTD report periods -> single-quarter revenue / gross profit / net profit."""
    by = {p["period"]: p for p in periods if re.match(r"\d{4}-\d{2}-\d{2}", p["period"])}
    q_end = {"03-31": 1, "06-30": 2, "09-30": 3, "12-31": 4}
    prev_end = {2: "03-31", 3: "06-30", 4: "09-30"}
    quarters = []
    for per in sorted(by):
        y, md = per[:4], per[5:]
        if md not in q_end:
            continue
        q = q_end[md]
        cur = by[per]
        rev, npf, gm = cur["revenue"], cur["net_profit"], cur["gross_margin_pct"]
        gp = rev * gm / 100 if rev is not None and gm is not None else None
        if q > 1:
            prv = by.get(f"{y}-{prev_end[q]}")
            if not prv:
                continue
            prev_gp = (prv["revenue"] * prv["gross_margin_pct"] / 100
                       if prv["revenue"] is not None and prv["gross_margin_pct"] is not None else None)
            rev = rev - prv["revenue"] if rev is not None and prv["revenue"] is not None else None
            npf = npf - prv["net_profit"] if npf is not None and prv["net_profit"] is not None else None
            gp = gp - prev_gp if gp is not None and prev_gp is not None else None
        quarters.append({
            "quarter": f"{y}Q{q}", "period_end": per,
            "revenue": rev, "gross_profit": gp, "net_profit": npf,
            "inventory_days_ytd": cur["inventory_days"],
        })
    return quarters


def ttm(periods, annual):
    """TTM = latest YTD + last FY - same YTD prior year."""
    dated = sorted(p["period"] for p in periods if re.match(r"\d{4}-\d{2}-\d{2}", p["period"]))
    if not dated:
        return {}
    latest = dated[-1]
    by = {p["period"]: p for p in periods}
    cur = by[latest]
    y, md = int(latest[:4]), latest[5:]
    res = {"latest_period": latest}
    for f in ("revenue", "net_profit", "net_profit_adj"):
        if md == "12-31":
            res[f] = cur[f]
            continue
        fy = by.get(f"{y-1}-12-31")
        prior = by.get(f"{y-1}-{md}")
        if fy and prior and None not in (cur[f], fy[f], prior[f]):
            res[f] = cur[f] + fy[f] - prior[f]
        else:
            res[f] = None
    res["bvps_latest"] = cur["bvps"]
    res["gross_margin_ytd_pct"] = cur["gross_margin_pct"]
    res["roe_ytd_pct"] = cur["roe_pct"]
    res["inventory_days_ytd"] = cur["inventory_days"]
    # prior-year same-period YTD for growth
    prior = by.get(f"{y-1}-{md}")
    res["revenue_ytd"] = cur["revenue"]
    res["net_profit_ytd"] = cur["net_profit"]
    res["revenue_ytd_prior"] = prior["revenue"] if prior else None
    res["net_profit_ytd_prior"] = prior["net_profit"] if prior else None
    return res


def valuation(code, asof):
    df = retry(lambda: ak.stock_value_em(symbol=code))
    df["数据日期"] = pd.to_datetime(df["数据日期"])
    df = df[df["数据日期"] <= pd.Timestamp(asof)].sort_values("数据日期")
    last = df.iloc[-1]
    start3 = pd.Timestamp(asof) - pd.DateOffset(years=3)
    h = df[df["数据日期"] >= start3]

    def stats(col, positive_only=False):
        s = h[col].dropna()
        if positive_only:
            s = s[s > 0]
        if s.empty:
            return None
        cur = float(last[col]) if pd.notna(last[col]) else None
        pct = float((s <= cur).mean() * 100) if cur is not None else None
        return {"current": cur, "min": float(s.min()), "median": float(s.median()),
                "max": float(s.max()), "percentile_3y": pct, "n_days": int(len(s))}

    monthly = (h.set_index("数据日期")[["市净率", "PE(TTM)"]]
               .resample("ME").last().dropna(how="all"))
    return {
        "date": last["数据日期"].strftime("%Y-%m-%d"),
        "close": float(last["当日收盘价"]),
        "mkt_cap": float(last["总市值"]),
        "shares": float(last["总股本"]),
        "pe_ttm_em": float(last["PE(TTM)"]) if pd.notna(last["PE(TTM)"]) else None,
        "pb_em": float(last["市净率"]) if pd.notna(last["市净率"]) else None,
        "ps_em": float(last["市销率"]) if pd.notna(last["市销率"]) else None,
        "pb_3y": stats("市净率"),
        "pe_3y": stats("PE(TTM)", positive_only=True),
        "pb_monthly": {d.strftime("%Y-%m"): (None if pd.isna(v) else round(float(v), 3))
                       for d, v in monthly["市净率"].items()},
    }


def performance(code, asof):
    a = pd.Timestamp(asof)
    df = retry(lambda: ak.stock_zh_a_hist(symbol=code, period="daily",
                                          start_date=(a - pd.DateOffset(days=400)).strftime("%Y%m%d"),
                                          end_date=a.strftime("%Y%m%d"), adjust="qfq"))
    df["日期"] = pd.to_datetime(df["日期"])
    df = df.sort_values("日期").set_index("日期")["收盘"]

    def close_on_or_before(d):
        s = df[df.index <= d]
        return (s.index[-1].strftime("%Y-%m-%d"), float(s.iloc[-1])) if len(s) else (None, None)

    points = {
        "last": close_on_or_before(a),
        "1M": close_on_or_before(a - pd.DateOffset(months=1)),
        "3M": close_on_or_before(a - pd.DateOffset(months=3)),
        "YTD": close_on_or_before(pd.Timestamp(f"{a.year - 1}-12-31")),
        "1Y": close_on_or_before(a - pd.DateOffset(years=1)),
    }
    return {k: {"date": d, "close_qfq": c} for k, (d, c) in points.items()}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--universe", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--asof", default=dt.date.today().isoformat())
    args = ap.parse_args()
    os.makedirs(args.out, exist_ok=True)

    with open(args.universe, encoding="utf-8") as f:
        universe = list(csv.DictReader(f))

    companies, failures = [], []
    for u in universe:
        code = u["code"].zfill(6)
        rec = dict(u, code=code)
        print(f"[{code}] {u['name_cn']} ...", flush=True)
        for part, fn in (("valuation", lambda: valuation(code, args.asof)),
                         ("financials", lambda: financials(code)),
                         ("performance", lambda: performance(code, args.asof))):
            try:
                rec[part] = fn()
            except Exception as e:
                rec[part] = None
                failures.append({"code": code, "part": part, "error": repr(e)[:300]})
                print(f"   ! {part} 抓取失败: {e!r}"[:200], flush=True)
        if rec.get("financials"):
            rec["quarters"] = derive_quarters(rec["financials"]["periods"])
            rec["ttm"] = ttm(rec["financials"]["periods"], rec["financials"]["annual"])
        companies.append(rec)

    dataset = {
        "as_of": args.asof,
        "fetched_at": dt.datetime.now().isoformat(timespec="seconds"),
        "units": "金额为元（原始值）；*_pct 字段为百分数",
        "sources": {
            "prices_valuation": "东方财富（AKShare stock_value_em / stock_individual_info_em）",
            "returns": "东方财富（AKShare stock_zh_a_hist，前复权）",
            "financials": "同花顺（AKShare stock_financial_abstract_ths）",
        },
        "companies": companies,
        "failures": failures,
    }
    path = os.path.join(args.out, "dataset.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(dataset, f, ensure_ascii=False, indent=1)
    print(f"\n已写入 {path}：{len(companies)} 家公司，{len(failures)} 项失败")
    for fl in failures:
        print("  失败", fl)


if __name__ == "__main__":
    main()

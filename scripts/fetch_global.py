"""抓取海外存储公司数据（Yahoo Finance，经 yfinance），并统一折算为人民币。

输出 <out>/global.json：估值（市值、PE、一致预期远期 PE、PB、EV/EBITDA）、TTM、近 6 个季度单季损益、股价表现。
用法: python fetch_global.py --universe universe_global.csv --out data/ [--asof 2026-11-09]
"""
import argparse, csv, datetime as dt, json, math, os, time

import pandas as pd
import yfinance as yf

FX_TICKER = {"USD": "CNY=X", "KRW": "KRWCNY=X", "JPY": "JPYCNY=X", "TWD": "TWDCNY=X",
             "HKD": "HKDCNY=X", "EUR": "EURCNY=X"}


def clean(v):
    if v is None:
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if math.isnan(f) or math.isinf(f) else f


def retry(fn, tries=3, wait=3):
    for i in range(tries):
        try:
            return fn()
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(wait * (i + 1))


def fx_rates(currencies):
    out = {"CNY": 1.0}
    for c in currencies:
        if c in out:
            continue
        t = FX_TICKER.get(c)
        if not t:
            out[c] = None
            continue
        px = retry(lambda: yf.Ticker(t).history(period="5d")["Close"].dropna())
        out[c] = float(px.iloc[-1]) if len(px) else None
    return out


def cal_quarter(ts):
    ts = pd.Timestamp(ts)
    return f"{ts.year}Q{(ts.month - 1) // 3 + 1}"


def quarters(tk):
    q = retry(lambda: tk.quarterly_income_stmt)
    rows = []
    for col in sorted(q.columns):
        g = lambda k: clean(q.at[k, col]) if k in q.index else None
        rev, gp, ni = g("Total Revenue"), g("Gross Profit"), g("Net Income")
        if rev is None or rev <= 0:
            continue
        warn = None
        if gp is not None and not (-1 <= gp / rev <= 1):
            warn = f"毛利 {gp:,.0f} 超出营收 {rev:,.0f} 的合理范围，已置空（疑似数据源错误）"
            gp = None
        rows.append({"period_end": pd.Timestamp(col).strftime("%Y-%m-%d"), "quarter": cal_quarter(col),
                     "revenue": rev, "gross_profit": gp, "net_profit": ni, "warning": warn})
    return rows[-6:]


def ttm_from_quarters(qs):
    """连续 4 个季度（相邻期末间隔 80–100 天）才计算 TTM。"""
    if len(qs) < 4:
        return None
    last4 = qs[-4:]
    ends = [pd.Timestamp(x["period_end"]) for x in last4]
    if any(not (80 <= (b - a).days <= 100) for a, b in zip(ends, ends[1:])):
        return None
    if any(x["net_profit"] is None for x in last4):
        return None
    return {"revenue": sum(x["revenue"] for x in last4),
            "net_profit": sum(x["net_profit"] for x in last4),
            "gross_profit": sum(x["gross_profit"] for x in last4) if all(x["gross_profit"] is not None for x in last4) else None,
            "through": last4[-1]["period_end"]}


def performance(tk, asof):
    a = pd.Timestamp(asof)
    h = retry(lambda: tk.history(start=(a - pd.DateOffset(days=400)).strftime("%Y-%m-%d"),
                                 end=(a + pd.DateOffset(days=1)).strftime("%Y-%m-%d"), auto_adjust=True))
    s = h["Close"].dropna()
    s.index = s.index.tz_localize(None)

    def on_or_before(d):
        x = s[s.index <= d]
        return {"date": x.index[-1].strftime("%Y-%m-%d"), "close_adj": float(x.iloc[-1])} if len(x) else {}

    return {"last": on_or_before(a), "1M": on_or_before(a - pd.DateOffset(months=1)),
            "3M": on_or_before(a - pd.DateOffset(months=3)), "YTD": on_or_before(pd.Timestamp(f"{a.year - 1}-12-31")),
            "1Y": on_or_before(a - pd.DateOffset(years=1))}


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
        t = u["ticker"]
        print(f"[{t}] {u['name_cn']} ...", flush=True)
        rec = dict(u)
        tk = yf.Ticker(t)
        try:
            i = retry(lambda: tk.info)
            rec["info"] = {k: (clean(i.get(k)) if k not in ("currency", "financialCurrency") else i.get(k)) for k in (
                "currency", "financialCurrency", "regularMarketPrice", "marketCap", "enterpriseValue",
                "trailingPE", "forwardPE", "priceToBook", "bookValue", "enterpriseToEbitda",
                "totalRevenue", "netIncomeToCommon", "grossMargins", "returnOnEquity", "revenueGrowth",
                "sharesOutstanding")}
        except Exception as e:
            rec["info"] = None
            failures.append({"ticker": t, "part": "info", "error": repr(e)[:300]})
        for part, fn in (("quarters", lambda: quarters(tk)), ("performance", lambda: performance(tk, args.asof))):
            try:
                rec[part] = fn()
            except Exception as e:
                rec[part] = None
                failures.append({"ticker": t, "part": part, "error": repr(e)[:300]})
                print(f"   ! {part} 抓取失败: {e!r}"[:200], flush=True)
        rec["ttm_q"] = ttm_from_quarters(rec.get("quarters") or [])
        companies.append(rec)

    curs = {c["info"][k] for c in companies if c.get("info") for k in ("currency", "financialCurrency") if c["info"].get(k)}
    fx = fx_rates(sorted(curs))
    out = {
        "as_of": args.asof, "fetched_at": dt.datetime.now().isoformat(timespec="seconds"),
        "fx_to_cny": fx,
        "source": "Yahoo Finance（yfinance）：估值、一致预期远期 PE、季度损益、复权股价；汇率为 Yahoo 外汇报价",
        "companies": companies, "failures": failures,
    }
    path = os.path.join(args.out, "global.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=1)
    print(f"\n已写入 {path}：{len(companies)} 家海外公司，{len(failures)} 项失败；汇率 {fx}")
    for fl in failures:
        print("  失败", fl)


if __name__ == "__main__":
    main()

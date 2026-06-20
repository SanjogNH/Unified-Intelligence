"""
process.py  —  Nat Habit Unified Intelligence Dashboard
Reads three sheets from Google Sheets (exported as xlsx or fetched via API),
computes all metrics and decision flags, writes processed.json.

Input sheets (expected column names):
  Sheet 1 - Sales Master:
    Platform, MTD Updated Till (Date), SKU, Short Name, Category,
    Planned Quantity, Planned MRP Revenue, Planned SP Revenue,
    MTD Actual Quantity, MTD Actual MRP Revenue, MTD Actual SP Revenue,
    Last Month Units, Last Month SP Revenue,
    Last 3month Units, Last 3month SP Revenue

  Sheet 2 - City Master:
    Platform, City, SKU, Short Name, Category,
    MTD Actual Quantity, Last Month Actual Quantity, Last 3 Months Actual Quantity

  Sheet 3 - Ads Data:
    Platform, Time, SKU,
    Gross Clicks, Gross Units, Gross Sales,
    Ad Spend, Ad Impressions, Ad Clicks, Ad Units, Ad Sales
    (Time is either a date value for current MTD, or the string 'LFM' / 'L3M')

Output:
  processed.json  — three top-level keys: skus, city, meta
"""

import json
import sys
import math
from datetime import datetime, date
from pathlib import Path

import pandas as pd
import numpy as np


# ─── Thresholds (adjust here, never scattered through code) ───────────────────
T = {
    "SPEND_BLEED_MIN_SPEND":    10_000,   # ₹ MTD spend threshold
    "SPEND_BLEED_MAX_ROAS":     2.0,
    "ROAS_DEGRADED_DROP_PCT":   30,       # % drop vs LFM
    "TACOS_HIGH":               30,       # %
    "SCALE_MIN_ROAS":           5.0,
    "SCALE_MAX_ACH":            80,       # % achievement
    "AD_DEPENDENT_ORG_PCT":     20,       # % organic units
    "DEAD_SKU_MIN_LM":          0,        # LM rev > this to flag
    "VELOCITY_DROP_TREND":      60,       # % MoM
    "VELOCITY_DROP_MIN_LM":     5_000,    # ₹ LM rev
    "BREAKOUT_TREND":           150,      # % MoM
    "BREAKOUT_MIN_MTD":         1_000,    # ₹ MTD rev
    "CITY_CONC_SHARE":          80,       # % top city
    "CITY_CONC_MIN_QTY":        50,
}

# Per-platform CVR/ROAS benchmarks derived from data analysis
PLATFORM_BENCHMARKS = {
    "Amazon":     {"median_roas": 2.35, "median_tacos": 18.9, "median_cvr": 21.1},
    "Blinkit":    {"median_roas": 1.55, "median_tacos": 28.4, "median_cvr": 47.0},
    "Zepto":      {"median_roas": 1.75, "median_tacos": 32.2, "median_cvr": 27.8},
    "Flipkart":   {"median_roas": 5.21, "median_tacos":  5.7, "median_cvr":  9.7},
    "Instamart":  {"median_roas": 1.17, "median_tacos": 38.1, "median_cvr": 15.8},
    "Nykaa":      {"median_roas": 1.93, "median_tacos":  9.2, "median_cvr":  5.6},
    "First Cry":  {"median_roas": 0.87, "median_tacos":  0.0, "median_cvr":  2.1},
    "Myntra":     {"median_roas": 1.61, "median_tacos": 11.0, "median_cvr":  8.9},
    "Big Basket": {"median_roas": 0.93, "median_tacos": 47.8, "median_cvr": 213.1},
}


# ─── Helpers ──────────────────────────────────────────────────────────────────

def safe(v):
    """Convert numpy scalars / NaN / inf to Python native for JSON."""
    if v is None:
        return None
    if isinstance(v, float) and (math.isnan(v) or math.isinf(v)):
        return None
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        return float(v)
    if isinstance(v, (pd.Timestamp, datetime, date)):
        try:
            return v.strftime("%Y-%m-%d")
        except (ValueError, OSError):
            return None
    return v


def div(a, b, scale=1.0):
    """Safe division, returns None if b is zero/None."""
    try:
        if b and b != 0:
            return a / b * scale
        return None
    except Exception:
        return None


def days_in_month(d):
    if pd.isna(d):
        return None
    next_month = pd.Timestamp(d.year + (d.month // 12), (d.month % 12) + 1, 1)
    return (next_month - pd.Timestamp(d.year, d.month, 1)).days


def pro_ratio(mtd_date):
    """day-of-month / total-days-in-month."""
    if pd.isna(mtd_date):
        return None
    dim = days_in_month(mtd_date)
    if not dim:
        return None
    return mtd_date.day / dim


# ─── Load data ────────────────────────────────────────────────────────────────

def load_data(sales_path: str, city_path: str = None, ads_path: str = None):
    """
    Flexible loader — three layouts supported:

    Layout A — single xlsx with 3 sheets (sales_path only):
      sheet 0 = sales master, sheet 1 = city master, sheet 2 = ads

    Layout B — sales xlsx has 2 sheets + separate ads file (Google Sheets default):
      sales_path sheet 0 = sales master, sheet 1 = city master
      ads_path = ads data xlsx

    Layout C — three fully separate files:
      sales_path = sales master, city_path = city master, ads_path = ads data
    """
    xl_sales = pd.ExcelFile(sales_path)
    n_sheets  = len(xl_sales.sheet_names)

    if city_path is None and ads_path is None:
        # Layout A: all 3 sheets in one file
        df_sales = pd.read_excel(xl_sales, sheet_name=xl_sales.sheet_names[0])
        df_city  = pd.read_excel(xl_sales, sheet_name=xl_sales.sheet_names[1])
        df_ads   = pd.read_excel(xl_sales, sheet_name=xl_sales.sheet_names[2])

    elif city_path is None and ads_path is not None:
        # Layout B: sales xlsx has sales + city, ads is separate
        df_sales = pd.read_excel(xl_sales, sheet_name=xl_sales.sheet_names[0])
        df_city  = pd.read_excel(xl_sales, sheet_name=xl_sales.sheet_names[1])
        df_ads   = pd.read_excel(ads_path)

    else:
        # Layout C: three separate files
        df_sales = pd.read_excel(sales_path)
        df_city  = pd.read_excel(city_path)
        df_ads   = pd.read_excel(ads_path)

    return df_sales, df_city, df_ads


# ─── Process Sales ────────────────────────────────────────────────────────────

def process_sales(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # Normalise column names (strip whitespace)
    df.columns = [c.strip() for c in df.columns]

    # Parse MTD date
    df["_mtd_date"] = pd.to_datetime(df["MTD Updated Till (Date)"], errors="coerce")
    df["_pro_ratio"] = df["_mtd_date"].apply(pro_ratio)

    # Numeric coerce
    num_cols = [
        "Planned Quantity", "Planned MRP Revenue", "Planned SP Revenue",
        "MTD Actual Quantity", "MTD Actual MRP Revenue", "MTD Actual SP Revenue",
        "Last Month Units", "Last Month SP Revenue",
        "Last 3month Units", "Last 3month SP Revenue",
    ]
    for c in num_cols:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0)

    # Pro-rata plan
    df["prorata_rev"] = df["Planned SP Revenue"] * df["_pro_ratio"]
    df["prorata_qty"] = df["Planned Quantity"]   * df["_pro_ratio"]

    # Achievement
    df["rev_ach_pct"] = df.apply(lambda r: div(r["MTD Actual SP Revenue"], r["prorata_rev"], 100), axis=1)
    df["qty_ach_pct"] = df.apply(lambda r: div(r["MTD Actual Quantity"],   r["prorata_qty"], 100), axis=1)

    # Deltas vs pro-rata
    df["rev_delta"] = df["MTD Actual SP Revenue"] - df["prorata_rev"].fillna(0)
    df["qty_delta"] = df["MTD Actual Quantity"]   - df["prorata_qty"].fillna(0)

    # Discount %
    df["disc_pct"] = df.apply(
        lambda r: div(r["MTD Actual MRP Revenue"] - r["MTD Actual SP Revenue"],
                      r["MTD Actual MRP Revenue"], 100), axis=1
    )

    # MoM trend (actual vs pro-rated last month)
    df["lm_prorated_rev"] = df["Last Month SP Revenue"] * df["_pro_ratio"]
    df["lm_prorated_qty"] = df["Last Month Units"]      * df["_pro_ratio"]
    df["mom_rev_trend"]   = df.apply(lambda r: div(r["MTD Actual SP Revenue"], r["lm_prorated_rev"], 100), axis=1)
    df["mom_qty_trend"]   = df.apply(lambda r: div(r["MTD Actual Quantity"],   r["lm_prorated_qty"], 100), axis=1)

    # L3M monthly avg
    df["l3m_avg_rev"] = df["Last 3month SP Revenue"] / 3
    df["l3m_avg_qty"] = df["Last 3month Units"]      / 3

    return df


# ─── Process Ads ──────────────────────────────────────────────────────────────

def process_ads(df: pd.DataFrame):
    """
    Returns three DataFrames: mtd, lfm, l3m.
    Time column: datetime values = MTD; 'LFM' = last full month; 'L3M' = last 3 months.
    """
    df = df.copy()
    df.columns = [c.strip() for c in df.columns]

    num_cols = ["Gross Clicks", "Gross Units", "Gross Sales",
                "Ad Spend", "Ad Impressions", "Ad Clicks", "Ad Units", "Ad Sales"]
    for c in num_cols:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0)

    def is_date(v):
        return not isinstance(v, str)

    mask_mtd = df["Time"].apply(is_date)
    mask_lfm = df["Time"] == "LFM"
    mask_l3m = df["Time"] == "L3M"

    def enrich(sub):
        s = sub.copy()
        s["roas"]      = s.apply(lambda r: div(r["Ad Sales"],         r["Ad Spend"]),          axis=1)
        s["acos"]      = s.apply(lambda r: div(r["Ad Spend"],         r["Ad Sales"],    100),   axis=1)
        s["ctr"]       = s.apply(lambda r: div(r["Ad Clicks"],        r["Ad Impressions"], 100), axis=1)
        s["ad_cvr"]    = s.apply(lambda r: div(r["Ad Units"],         r["Ad Clicks"],   100),   axis=1)
        s["org_units"] = s["Gross Units"] - s["Ad Units"]
        s["org_pct"]   = s.apply(lambda r: div(r["org_units"],        r["Gross Units"], 100),   axis=1)
        s["tacos"]     = s.apply(lambda r: div(r["Ad Spend"],         r["Gross Sales"], 100),   axis=1)
        s["traffic"]   = s["Gross Units"]   # raw traffic volume
        return s

    return enrich(df[mask_mtd]), enrich(df[mask_lfm]), enrich(df[mask_l3m])


# ─── Decision Flags ───────────────────────────────────────────────────────────

def compute_flags(row: dict) -> dict:
    """
    Given a merged row dict (sales + ads MTD + ads LFM),
    return a flags dict: {flag_name: bool}.
    All threshold logic lives here.
    """
    flags = {}

    spend   = row.get("ad_spend")      or 0
    roas    = row.get("roas")
    lfm_roas= row.get("lfm_roas")
    tacos   = row.get("tacos")
    org_pct = row.get("org_pct")
    rev_ach = row.get("rev_ach_pct")
    mtd_rev = row.get("mtd_actual_rev") or 0
    lm_rev  = row.get("lm_rev")         or 0
    mom     = row.get("mom_rev_trend")

    # ── Ads flags ──
    flags["spend_bleed"] = bool(
        spend >= T["SPEND_BLEED_MIN_SPEND"]
        and roas is not None
        and roas < T["SPEND_BLEED_MAX_ROAS"]
    )

    if roas is not None and lfm_roas is not None and lfm_roas > 0:
        roas_drop = (lfm_roas - roas) / lfm_roas * 100
        flags["roas_degraded"] = roas_drop > T["ROAS_DEGRADED_DROP_PCT"]
    else:
        flags["roas_degraded"] = False

    flags["tacos_high"] = bool(tacos is not None and tacos > T["TACOS_HIGH"])

    flags["scale_now"] = bool(
        roas is not None
        and roas >= T["SCALE_MIN_ROAS"]
        and rev_ach is not None
        and rev_ach < T["SCALE_MAX_ACH"]
    )

    flags["ad_dependent"] = bool(
        org_pct is not None
        and org_pct < T["AD_DEPENDENT_ORG_PCT"]
    )

    # ── Sales flags ──
    flags["dead_sku"] = bool(mtd_rev == 0 and lm_rev > T["DEAD_SKU_MIN_LM"])

    flags["velocity_drop"] = bool(
        mom is not None
        and mom < T["VELOCITY_DROP_TREND"]
        and lm_rev > T["VELOCITY_DROP_MIN_LM"]
    )

    flags["breakout"] = bool(
        mom is not None
        and mom > T["BREAKOUT_TREND"]
        and mtd_rev > T["BREAKOUT_MIN_MTD"]
    )

    return flags


def flag_priority(flags: dict) -> str:
    """
    Returns the single highest-priority flag label for inline display.
    Order: critical → warning → opportunity → none
    """
    priority_order = [
        "spend_bleed", "dead_sku", "roas_degraded",
        "tacos_high", "velocity_drop", "ad_dependent",
        "scale_now", "breakout",
    ]
    for f in priority_order:
        if flags.get(f):
            labels = {
                "spend_bleed":   "BLEED",
                "dead_sku":      "DEAD",
                "roas_degraded": "ROAS↓",
                "tacos_high":    "TACoS↑",
                "velocity_drop": "DROP",
                "ad_dependent":  "AD-DEP",
                "scale_now":     "SCALE↑",
                "breakout":      "BREAKOUT",
            }
            return labels[f]
    return None


FLAG_SEVERITY = {
    "spend_bleed":   "critical",
    "dead_sku":      "critical",
    "roas_degraded": "warning",
    "tacos_high":    "warning",
    "velocity_drop": "warning",
    "ad_dependent":  "warning",
    "scale_now":     "opportunity",
    "breakout":      "opportunity",
}

FLAG_ACTION = {
    "spend_bleed":   "Pause or cap daily budget — spending more than ad revenue",
    "dead_sku":      "Zero sales this month — check availability, listing, or OOS",
    "roas_degraded": "ROAS dropped >30% vs last month — review bids and keywords",
    "tacos_high":    "Ad spend >30% of gross sales — ads consuming margin",
    "velocity_drop": "Sales running <60% of last month's pace — investigate cause",
    "ad_dependent":  "<20% organic units — fragile revenue, improve listing quality",
    "scale_now":     "High ROAS but below plan — increase budget by 30–50%",
    "breakout":      "Trending >150% vs last month — protect inventory, double down",
}


# ─── City Concentration ───────────────────────────────────────────────────────

def city_concentration(df_city: pd.DataFrame) -> dict:
    """
    Returns a dict keyed by (Platform, SKU) with top_city, top_city_share, city_count.
    """
    df = df_city.copy()
    df.columns = [c.strip() for c in df.columns]
    df["MTD Actual Quantity"] = pd.to_numeric(df.get("MTD Actual Quantity", 0), errors="coerce").fillna(0)

    result = {}
    grouped = df.groupby(["Platform", "SKU"])
    for (plat, sku), grp in grouped:
        total = grp["MTD Actual Quantity"].sum()
        if total == 0:
            continue
        top = grp.loc[grp["MTD Actual Quantity"].idxmax()]
        top_share = top["MTD Actual Quantity"] / total * 100
        result[(plat, sku)] = {
            "top_city":       str(top["City"]),
            "top_city_share": round(top_share, 1),
            "city_count":     int(grp["City"].nunique()),
            "total_qty":      int(total),
        }
    return result


# ─── Channel-level summaries (for Signals view) ───────────────────────────────

def channel_summaries(sku_rows: list) -> list:
    """
    Aggregate sku_rows by platform to produce channel-level summary
    including flag counts and top priority action sentence.
    """
    channels = {}
    for r in sku_rows:
        ch = r["platform"]
        if ch not in channels:
            channels[ch] = {
                "platform": ch,
                "prorata_rev": 0, "actual_rev": 0,
                "prorata_qty": 0, "actual_qty": 0,
                "lm_rev": 0, "l3m_avg_rev": 0,
                "flag_counts": {
                    "critical": 0, "warning": 0, "opportunity": 0
                },
                "top_flags": {},   # flag_name -> count
                "sku_count": 0,
                "active_sku_count": 0,
                "total_ad_spend": 0,
                "benchmark": PLATFORM_BENCHMARKS.get(ch, {}),
            }
        c = channels[ch]
        c["prorata_rev"]     += r.get("prorata_rev") or 0
        c["actual_rev"]      += r.get("mtd_actual_rev") or 0
        c["prorata_qty"]     += r.get("prorata_qty") or 0
        c["actual_qty"]      += r.get("mtd_actual_qty") or 0
        c["lm_rev"]          += r.get("lm_rev") or 0
        c["l3m_avg_rev"]     += r.get("l3m_avg_rev") or 0
        c["total_ad_spend"]  += r.get("ad_spend") or 0
        c["sku_count"]       += 1
        if (r.get("mtd_actual_rev") or 0) > 0:
            c["active_sku_count"] += 1

        for fname, fval in r.get("flags", {}).items():
            if fval:
                sev = FLAG_SEVERITY.get(fname, "warning")
                c["flag_counts"][sev] += 1
                c["top_flags"][fname] = c["top_flags"].get(fname, 0) + 1

    out = []
    for ch, c in channels.items():
        rev_ach = div(c["actual_rev"], c["prorata_rev"], 100)
        lm_pro  = c["lm_rev"] * (c["actual_rev"] / c["actual_rev"] if c["actual_rev"] else 1)

        # Top priority signal sentence
        top_flag = max(c["top_flags"].items(), key=lambda x: x[1], default=(None, 0))
        if top_flag[0]:
            signal = f"{top_flag[1]} SKUs: {FLAG_ACTION[top_flag[0]]}"
        else:
            signal = "No critical flags"

        out.append({
            **c,
            "rev_ach_pct":   safe(rev_ach),
            "top_signal":    signal,
        })

    out.sort(key=lambda x: x["actual_rev"], reverse=True)
    return out


# ─── Category × Platform heatmap ──────────────────────────────────────────────

def category_heatmap(sku_rows: list) -> list:
    """
    Returns [{category, platform, rev_ach_pct, actual_rev, prorata_rev}]
    Used by the Signals view heatmap.
    """
    agg = {}
    for r in sku_rows:
        key = (r["category"], r["platform"])
        if key not in agg:
            agg[key] = {"actual": 0, "prorata": 0}
        agg[key]["actual"]  += r.get("mtd_actual_rev") or 0
        agg[key]["prorata"] += r.get("prorata_rev")    or 0

    out = []
    for (cat, plat), v in agg.items():
        out.append({
            "category":    cat,
            "platform":    plat,
            "actual_rev":  round(v["actual"], 2),
            "prorata_rev": round(v["prorata"], 2),
            "rev_ach_pct": safe(div(v["actual"], v["prorata"], 100)),
        })
    return out


# ─── Top Actions (Signals view ranked list) ───────────────────────────────────

def top_actions(sku_rows: list) -> list:
    """
    Produces a ranked list of actionable decisions across all channels,
    sorted by revenue impact.
    """
    action_map = {}  # (flag, platform) -> {skus, spend_at_risk, potential_rev}

    for r in sku_rows:
        for fname, fval in r.get("flags", {}).items():
            if not fval:
                continue
            key = (fname, r["platform"])
            if key not in action_map:
                action_map[key] = {
                    "flag":        fname,
                    "platform":    r["platform"],
                    "sku_count":   0,
                    "impact_rev":  0,
                    "severity":    FLAG_SEVERITY.get(fname, "warning"),
                    "action":      FLAG_ACTION.get(fname, ""),
                    "skus":        [],
                }
            entry = action_map[key]
            entry["sku_count"] += 1
            entry["skus"].append(r["sku"])

            if fname == "spend_bleed":
                entry["impact_rev"] += r.get("ad_spend") or 0
            elif fname == "dead_sku":
                entry["impact_rev"] += r.get("lm_rev")    or 0
            elif fname == "scale_now":
                entry["impact_rev"] += abs(r.get("rev_delta") or 0)
            elif fname in ("velocity_drop", "roas_degraded"):
                entry["impact_rev"] += abs(r.get("rev_delta") or 0)
            elif fname == "breakout":
                entry["impact_rev"] += r.get("mtd_actual_rev") or 0
            else:
                entry["impact_rev"] += abs(r.get("rev_delta") or 0)

    actions = list(action_map.values())

    # Severity weight: critical first, then by impact
    sev_weight = {"critical": 3, "warning": 2, "opportunity": 1}
    actions.sort(key=lambda x: (sev_weight.get(x["severity"], 0), x["impact_rev"]), reverse=True)

    # Keep top 20, trim sku list to 10 per action
    for a in actions[:20]:
        a["impact_rev"] = round(a["impact_rev"], 2)
        a["skus"] = a["skus"][:10]

    return actions[:20]


# ─── City sheet output ────────────────────────────────────────────────────────

def process_city_sheet(df_city: pd.DataFrame) -> list:
    df = df_city.copy()
    df.columns = [c.strip() for c in df.columns]

    for c in ["MTD Actual Quantity", "Last Month Actual Quantity", "Last 3 Months Actual Quantity"]:
        df[c] = pd.to_numeric(df.get(c, 0), errors="coerce").fillna(0)

    rows = []
    for _, r in df.iterrows():
        rows.append({
            "platform":    str(r.get("Platform", "")),
            "city":        str(r.get("City", "")),
            "sku":         str(r.get("SKU", "")),
            "short_name":  str(r.get("Short Name", "")),
            "category":    str(r.get("Category", "")),
            "mtd_qty":     safe(r.get("MTD Actual Quantity",           0)),
            "lm_qty":      safe(r.get("Last Month Actual Quantity",    0)),
            "l3m_qty":     safe(r.get("Last 3 Months Actual Quantity", 0)),
        })
    return [r for r in rows if r["platform"] and r["city"] and r["sku"]]


# ─── Main pipeline ────────────────────────────────────────────────────────────

def run(sales_path: str, city_path: str = None, ads_path: str = None,
        output_path: str = "processed.json"):

    print(f"[process.py] Loading data...")
    df_sales, df_city, df_ads = load_data(sales_path, city_path, ads_path)

    print(f"  Sales rows: {len(df_sales)}, City rows: {len(df_city)}, Ads rows: {len(df_ads)}")

    # ── Sales ──
    df_sales = process_sales(df_sales)

    # ── Ads: split into MTD / LFM / L3M ──
    df_mtd, df_lfm, df_l3m = process_ads(df_ads)

    # Index ads by (Platform, SKU) for fast lookup
    mtd_idx = {(r["Platform"], r["SKU"]): r for _, r in df_mtd.iterrows()}
    lfm_idx = {(r["Platform"], r["SKU"]): r for _, r in df_lfm.iterrows()}

    # ── City concentration ──
    city_conc = city_concentration(df_city)

    # ── Build SKU rows ──
    print(f"[process.py] Building SKU rows and computing flags...")
    sku_rows = []

    for _, s in df_sales.iterrows():
        plat = str(s.get("Platform", "")).strip()
        sku  = str(s.get("SKU", "")).strip()
        if not plat or not sku:
            continue

        key = (plat, sku)
        ads_mtd = mtd_idx.get(key, {})
        ads_lfm = lfm_idx.get(key, {})

        mtd_date = s.get("_mtd_date")

        row = {
            # ── identifiers ──
            "platform":       plat,
            "sku":            sku,
            "short_name":     str(s.get("Short Name", "")),
            "category":       str(s.get("Category", "")),
            "mtd_date":       safe(mtd_date),
            "pro_ratio":      safe(s.get("_pro_ratio")),

            # ── plan ──
            "planned_rev":    safe(s.get("Planned SP Revenue")),
            "planned_qty":    safe(s.get("Planned Quantity")),
            "prorata_rev":    safe(s.get("prorata_rev")),
            "prorata_qty":    safe(s.get("prorata_qty")),

            # ── actuals ──
            "mtd_actual_rev": safe(s.get("MTD Actual SP Revenue")),
            "mtd_actual_qty": safe(s.get("MTD Actual Quantity")),
            "mtd_actual_mrp": safe(s.get("MTD Actual MRP Revenue")),

            # ── achievement ──
            "rev_ach_pct":    safe(s.get("rev_ach_pct")),
            "qty_ach_pct":    safe(s.get("qty_ach_pct")),
            "rev_delta":      safe(s.get("rev_delta")),
            "qty_delta":      safe(s.get("qty_delta")),
            "disc_pct":       safe(s.get("disc_pct")),

            # ── historical ──
            "lm_rev":         safe(s.get("Last Month SP Revenue")),
            "lm_qty":         safe(s.get("Last Month Units")),
            "l3m_rev":        safe(s.get("Last 3month SP Revenue")),
            "l3m_qty":        safe(s.get("Last 3month Units")),
            "l3m_avg_rev":    safe(s.get("l3m_avg_rev")),
            "l3m_avg_qty":    safe(s.get("l3m_avg_qty")),
            "lm_prorated_rev":safe(s.get("lm_prorated_rev")),

            # ── velocity ──
            "mom_rev_trend":  safe(s.get("mom_rev_trend")),
            "mom_qty_trend":  safe(s.get("mom_qty_trend")),

            # ── ads (MTD) ──
            "ad_spend":       safe(ads_mtd.get("Ad Spend")),
            "ad_sales":       safe(ads_mtd.get("Ad Sales")),
            "ad_units":       safe(ads_mtd.get("Ad Units")),
            "ad_clicks":      safe(ads_mtd.get("Ad Clicks")),
            "ad_impressions": safe(ads_mtd.get("Ad Impressions")),
            "gross_units":    safe(ads_mtd.get("Gross Units")),
            "gross_sales":    safe(ads_mtd.get("Gross Sales")),
            "roas":           safe(ads_mtd.get("roas")),
            "acos":           safe(ads_mtd.get("acos")),
            "ctr":            safe(ads_mtd.get("ctr")),
            "ad_cvr":         safe(ads_mtd.get("ad_cvr")),
            "org_units":      safe(ads_mtd.get("org_units")),
            "org_pct":        safe(ads_mtd.get("org_pct")),
            "tacos":          safe(ads_mtd.get("tacos")),

            # ── ads (LFM comparison) ──
            "lfm_roas":       safe(ads_lfm.get("roas")),
            "lfm_acos":       safe(ads_lfm.get("acos")),
            "lfm_org_pct":    safe(ads_lfm.get("org_pct")),
            "lfm_gross_units":safe(ads_lfm.get("Gross Units")),

            # ── ROAS trend ──
            "roas_change_pct": safe(
                div(ads_mtd.get("roas", 0) - (ads_lfm.get("roas") or 0),
                    ads_lfm.get("roas"), 100)
                if ads_lfm.get("roas") else None
            ),

            # ── city concentration ──
            "city_conc": {
                "top_city":       city_conc.get(key, {}).get("top_city"),
                "top_city_share": city_conc.get(key, {}).get("top_city_share"),
                "city_count":     city_conc.get(key, {}).get("city_count"),
                "total_qty":      city_conc.get(key, {}).get("total_qty"),
                "flag":           bool(
                    city_conc.get(key, {}).get("top_city_share", 0) >= T["CITY_CONC_SHARE"]
                    and city_conc.get(key, {}).get("total_qty", 0)  >= T["CITY_CONC_MIN_QTY"]
                ),
            },
        }

        # ── Decision flags ──
        row["flags"] = compute_flags(row)
        row["top_flag"] = flag_priority(row["flags"])
        row["flag_actions"] = {
            fname: FLAG_ACTION[fname]
            for fname, fval in row["flags"].items() if fval
        }

        sku_rows.append(row)

    # ── Derived summaries ──
    print(f"[process.py] Building channel summaries and heatmap...")
    channels  = channel_summaries(sku_rows)
    heatmap   = category_heatmap(sku_rows)
    actions   = top_actions(sku_rows)
    city_rows = process_city_sheet(df_city)

    # ── Meta ──
    flag_totals = {
        "critical":    sum(1 for r in sku_rows for f, v in r["flags"].items() if v and FLAG_SEVERITY.get(f) == "critical"),
        "warning":     sum(1 for r in sku_rows for f, v in r["flags"].items() if v and FLAG_SEVERITY.get(f) == "warning"),
        "opportunity": sum(1 for r in sku_rows for f, v in r["flags"].items() if v and FLAG_SEVERITY.get(f) == "opportunity"),
    }

    meta = {
        "generated_at":     datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sku_count":        len(sku_rows),
        "platform_count":   len(set(r["platform"] for r in sku_rows)),
        "flag_totals":      flag_totals,
        "thresholds":       T,
        "platform_benchmarks": PLATFORM_BENCHMARKS,
    }

    # ── Write output ──
    out = {
        "meta":     meta,
        "skus":     sku_rows,
        "channels": channels,
        "heatmap":  heatmap,
        "actions":  actions,
        "city":     city_rows,
    }

    output_file = Path(output_path)
    output_file.parent.mkdir(parents=True, exist_ok=True)

    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, separators=(",", ":"))

    size_kb = output_file.stat().st_size / 1024
    print(f"[process.py] Done. Output: {output_path}  ({size_kb:.0f} KB)")
    print(f"  SKUs: {len(sku_rows)}, Channels: {len(channels)}, City rows: {len(city_rows)}")
    print(f"  Flags — critical: {flag_totals['critical']}, warning: {flag_totals['warning']}, opportunity: {flag_totals['opportunity']}")

    return out


# ─── CLI ──────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    """
    Usage:
      # Single xlsx with 3 sheets (sheet1=sales, sheet2=city, sheet3=ads):
      python process.py data/combined.xlsx

      # Three separate files:
      python process.py data/sales.xlsx data/city.xlsx data/ads.xlsx

      # Specify output path:
      python process.py data/sales.xlsx data/city.xlsx data/ads.xlsx output/processed.json
    """
    args = sys.argv[1:]
    if not args:
        print("Usage: python process.py <sales.xlsx> [city.xlsx ads.xlsx] [output.json]")
        sys.exit(1)

    # Detect if last arg is a .json output path
    if args[-1].endswith(".json"):
        out_path = args[-1]
        args = args[:-1]
    else:
        out_path = "processed.json"

    if len(args) == 1:
        # Layout A: single file with 3 sheets
        run(args[0], output_path=out_path)
    elif len(args) == 2:
        # Layout B: sales+city in one file, ads separate
        run(args[0], ads_path=args[1], output_path=out_path)
    elif len(args) == 3:
        # Layout C: three separate files
        run(args[0], args[1], args[2], output_path=out_path)
    else:
        print("Usage: process.py <sales_city.xlsx> <ads.xlsx> [output.json]")
        print("   or: process.py <sales.xlsx> <city.xlsx> <ads.xlsx> [output.json]")
        sys.exit(1)

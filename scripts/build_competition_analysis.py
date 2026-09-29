# -*- coding: utf-8 -*-
"""
多竞店竞争分析：价格带 / 热销 / 缺口 / 防守 / 品类结构 / 商品力对比 + 汇总报告

输入：
  - output/竞店标准库/多竞店汇总库.xlsx   19 店标准化汇总
  - output/竞品价格矩阵.xlsx             价格矩阵（价格矩阵/竞店独有 sheet）
  - output/自家_标准版本_对齐增强.xlsx    自家端（EAN条码/月销量/售价/采购价）

产出（output/ 下）：
  - 价格带分析.xlsx            每类目价格分位（P25/P50/P75）+ 主流价格带
  - 热销品_多店Top100.xlsx     跨店聚合热销（月售合计/在售店数/价格区间）
  - 缺口清单.xlsx              竞店多店热销 且 自家无匹配 -> 引入候选
  - 防守清单.xlsx              自家有销量 但 竞店均未卖 -> 维持优势
  - 竞店品类结构.xlsx          各店品类占比 + 优势品类 + 自家品类缺口
  - 商品力对比_建议.xlsx        自家SKU x 竞店：标题/价格/月售对比与建议
  - 竞争分析报告.xlsx          多 sheet 汇总（总览/矩阵/价格带/缺口/品类/商品力）

用法：
  python scripts/build_competition_analysis.py
"""
import os
import re
import sys

import pandas as pd
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from import_adapter import CN_STD
OUT = os.path.join(BASE, "output")
SUMMARY_FILE = os.path.join(OUT, "竞店标准库", "多竞店汇总库.xlsx")
MATRIX_FILE = os.path.join(OUT, "竞品价格矩阵.xlsx")
OWN_FILE = os.path.join(OUT, "自家_标准版本_对齐增强.xlsx")

HOT_MIN_SALES = 30      # 热销阈值：单店月售
GAP_MIN_SALES = 60      # 缺口阈值：跨店月售合计
GAP_MIN_SHOPS = 2       # 缺口阈值：在售店数
DEFEND_MIN_SALES = 30   # 防守阈值：自家月销量


def clean_num(s):
    """销量文本转 float：识别 万 单位与 已售/月售 前缀（如 '已售4.2万+' -> 42000）。"""
    if s is None or (isinstance(s, float) and np.isnan(s)):
        return np.nan
    t = str(s).strip()
    m = re.search(r"([\d.]+)\s*(万)?", t)
    if not m:
        return np.nan
    v = float(m.group(1))
    if m.group(2):
        v *= 10000
    return v


def load_summary(shops_filter=None):
    df = pd.read_excel(SUMMARY_FILE, dtype={"竞品商品编号": str, "竞品单品编号": str})
    df = df.rename(columns={k: v for k, v in CN_STD.items() if k in df.columns})
    if shops_filter is not None:
        df = df[df["shop_name"].isin(shops_filter)].copy()
    raw_sales = df["sales"].astype(str)
    df["sales"] = raw_sales.apply(clean_num).fillna(0)
    # 「想买」是意向不是已售，不混入月售口径（另列记录）
    df["意向数"] = raw_sales.str.contains("想买")
    df.loc[df["意向数"], "sales"] = 0
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    df["cat"] = df["category3"].fillna("").astype(str).str.strip().replace("", "未分类")
    df["name"] = df["name"].fillna("").astype(str).str.strip()
    df["title_len"] = df["name"].str.len()
    return df


def core_words(name):
    """标题核心词：去括号/数字/单位/符号，保留前 10 字符（用于同品不同标题的粗聚合）。"""
    s = re.sub(r"[（(][^）)]*[）)]", "", str(name))
    s = re.sub(r"[\d.]+(mm|cm|m|个|颗|条|只|根|片|包|卷|张|套|米|w|v|a)?", "", s)
    s = re.sub(r"[^\u4e00-\u9fa5A-Za-z]", "", s)
    return s[:10]


def load_own():
    df = pd.read_excel(OWN_FILE, dtype={"商品编号": str, "单品编号": str, "商品条码": str,
                                        "真实条码(13位)": str, "条码类型": str})
    # 兼容新中文列名与旧源表列名（SPUID/SKUID）
    if "单品编号" in df.columns:
        df = df.rename(columns={"商品编号": "SPUID", "单品编号": "SKUID"})
    if "真实条码(13位)" in df.columns:
        df = df.rename(columns={"真实条码(13位)": "EAN条码"})
    df["EAN条码"] = df["EAN条码"].fillna("").astype(str).str.strip()
    df["月销量"] = df["月销量"].apply(clean_num).fillna(0)
    df["售价"] = pd.to_numeric(df["美团闪购渠道售价"], errors="coerce")
    df["采购价"] = pd.to_numeric(df["采购价"], errors="coerce")
    return df


# ===========================================================================
# 2.1 价格带分析
# ===========================================================================
def price_band(df):
    rows = []
    for cat, g in df.groupby("cat"):
        p = g["price"].dropna()
        if len(p) < 5:
            continue
        q25, q50, q75 = p.quantile([0.25, 0.5, 0.75])
        rows.append({
            "类目": cat,
            "商品数": len(g),
            "在售店数": g["shop_name"].nunique(),
            "价格中位数": round(q50, 2),
            "25%分位价": round(q25, 2),
            "75%分位价": round(q75, 2),
            "最低价": round(p.min(), 2),
            "最高价": round(p.max(), 2),
            "主流价格带": f"{q25:.0f}-{q75:.0f}元",
            "月售合计": int(g["sales"].sum()),
        })
    out = pd.DataFrame(rows).sort_values("商品数", ascending=False)
    out.to_excel(os.path.join(OUT, "价格带分析.xlsx"), index=False)
    print(f"价格带分析: {len(out)} 个类目 -> output/价格带分析.xlsx")
    return out


# ===========================================================================
# 2.2 热销多店 + 缺口 + 防守
# ===========================================================================
def hot_products(df, own):
    """跨店聚合热销：按 (类目, 商品名) 聚合月售，输出 Top100。
    附核心词列，供同品不同标题的粗聚合参考。"""
    own_eans = set(own["EAN条码"][own["EAN条码"] != ""])
    df = df.copy()
    df["核心词"] = df["name"].apply(core_words)
    key = df["cat"] + "|" + df["name"]
    g = df.groupby(key).agg(
        类目=("cat", "first"), 商品名=("name", "first"), 核心词=("核心词", "first"),
        月售合计=("sales", "sum"), 在售店数=("shop_name", "nunique"),
        最低价=("price", "min"), 最高价=("price", "max"), 中位价=("price", "median"),
        款式数=("sku_id", "nunique"),
    ).reset_index(drop=True)
    # 同品不同标题粗聚合：核心词一致的并算（参考维度）
    kw_g = df.groupby(["cat", "核心词"]).agg(核心词店数=("shop_name", "nunique"))
    kw_g = kw_g.reset_index().rename(columns={"cat": "类目"})
    g = g.merge(kw_g, on=["类目", "核心词"], how="left")
    g["月售合计"] = g["月售合计"].astype(int)
    g = g[g["月售合计"] > 0].sort_values("月售合计", ascending=False)
    hot = g.head(100).copy()
    hot.to_excel(os.path.join(OUT, "热销品_多店Top100.xlsx"), index=False)
    print(f"热销品_多店Top100: {len(hot)} 条 -> output/热销品_多店Top100.xlsx")
    return g, hot


def gap_list(g, own):
    """缺口：跨店热销且自家无 EAN 匹配。"""
    own_eans = set(own["EAN条码"][own["EAN条码"] != ""])
    mask = (g["月售合计"] >= GAP_MIN_SALES) & (g["在售店数"] >= GAP_MIN_SHOPS)
    cand = g[mask].copy()
    # 自家名称包含度粗筛：商品名关键字与自家商品名重合则标记（最终以人工确认池为准）
    own_names = set(own["商品名称"].fillna("").astype(str).str.strip())
    has_own = []
    for _, r in cand.iterrows():
        kw = re.sub(r"[^\u4e00-\u9fa5A-Za-z0-9]", "", r["商品名"])[:8]
        hit = any(kw and kw in n for n in own_names)
        has_own.append(hit)
    cand["自家疑似已有(名称粗筛)"] = has_own
    cand["建议"] = "引入候选：参考价=" + cand["中位价"].round(1).astype(str) + "元，需询货源"
    cand = cand[~cand["自家疑似已有(名称粗筛)"]].copy()
    cand.to_excel(os.path.join(OUT, "缺口清单.xlsx"), index=False)
    print(f"缺口清单: {len(cand)} 条（月售合计≥{GAP_MIN_SALES} 且在售店数≥{GAP_MIN_SHOPS}）-> output/缺口清单.xlsx")
    return cand


def defend_list(own, summary):
    """防守：自家有销量但竞店均未卖。"""
    shop_eans = set()
    for e in summary["barcode"].dropna().astype(str):
        if len(e) == 13:
            shop_eans.add(e)
    own_eans = own["EAN条码"]
    own["竞店有卖(EAN)"] = own_eans.isin(shop_eans)
    d = own[(own["月销量"] >= DEFEND_MIN_SALES) & ~own["竞店有卖(EAN)"]].copy()
    d["毛利"] = np.where(d["售价"].notna() & d["采购价"].notna(),
                         (d["售价"] - d["采购价"]).round(2), "")
    d = d.sort_values("月销量", ascending=False)
    out = d[["SPUID", "SKUID", "商品名称", "美团闪购类目三级", "月销量", "售价", "采购价", "毛利"]]
    out = out.rename(columns={"SPUID": "商品编号", "SKUID": "单品编号"})
    out.to_excel(os.path.join(OUT, "防守清单.xlsx"), index=False)
    print(f"防守清单: {len(out)} 条（自家月销量≥{DEFEND_MIN_SALES} 且竞店均无同标准条码）-> output/防守清单.xlsx")
    return out


# ===========================================================================
# 2.3 品类结构
# ===========================================================================
def category_structure(df, own):
    # 各店品类商品数/月售
    shop_cat = df.groupby(["shop_name", "cat"]).agg(
        商品数=("sku_id", "count"), 月售合计=("sales", "sum")).reset_index()
    shop_total = df.groupby("shop_name").agg(总商品数=("sku_id", "count"), 总月售=("sales", "sum")).reset_index()
    shop_cat = shop_cat.merge(shop_total, on="shop_name")
    shop_cat = shop_cat.rename(columns={"shop_name": "竞店名称", "cat": "类目"})
    shop_cat["商品数占比%"] = (shop_cat["商品数"] / shop_cat["总商品数"] * 100).round(1)
    shop_cat["月售占比%"] = (shop_cat["月售合计"] / shop_cat["总月售"].replace(0, np.nan) * 100).round(1)
    # 优势品类：该店某品类月售占比 >= 15%
    shop_cat["优势品类?"] = shop_cat["月售占比%"] >= 15
    # 自家品类统计
    own_cat = own.groupby("美团闪购类目三级").agg(自家商品数=("SKUID", "count")).reset_index()
    own_cat.columns = ["类目", "自家商品数"]
    # 竞店有、自家无的类目（自家缺口）
    cat_summary = df.groupby("cat").agg(竞店商品数=("sku_id", "count"), 竞店在售店数=("shop_name", "nunique")).reset_index()
    cat_summary = cat_summary.rename(columns={"cat": "类目"})
    cat_summary = cat_summary.merge(own_cat, on="类目", how="left")
    cat_summary["自家商品数"] = cat_summary["自家商品数"].fillna(0).astype(int)
    cat_summary["自家缺口?"] = cat_summary["自家商品数"] == 0
    cat_summary = cat_summary.sort_values("竞店商品数", ascending=False)

    with pd.ExcelWriter(os.path.join(OUT, "竞店品类结构.xlsx"), engine="openpyxl") as w:
        shop_cat.to_excel(w, sheet_name="各店品类占比", index=False)
        cat_summary.to_excel(w, sheet_name="类目总览与自家缺口", index=False)
    print(f"竞店品类结构 -> output/竞店品类结构.xlsx")
    return shop_cat, cat_summary


# ===========================================================================
# 2.4 商品力对比
# ===========================================================================
def product_power(matrix_df, summary):
    """自家SKU x 竞店：价格/月售对比 + 建议。"""
    if not os.path.exists(MATRIX_FILE):
        print("缺少 竞品价格矩阵.xlsx，先跑 build_price_matrix.py")
        return pd.DataFrame()
    m = pd.read_excel(MATRIX_FILE, sheet_name="价格矩阵", dtype={"自家单品编号": str})
    own = load_own()
    own_sales = own.groupby("SKUID")["月销量"].first()
    rows = []
    for _, r in m.iterrows():
        sku = str(r["自家单品编号"])
        own_sale = float(own_sales.get(sku, 0))
        for col in r.index:
            if col.endswith("_到手价") and pd.notna(r[col]) and str(r[col]) != "":
                shop = col.replace("_到手价", "")
                rows.append({
                    "自家单品编号": sku, "自家商品名": r.get("自家商品名", ""),
                    "自家月销量": own_sale,
                    "竞店": shop, "竞店到手价": float(r[col]),
                    "竞店月售": r.get(f"{shop}_月售", ""), "自家售价": r.get("自家售价", ""),
                    "价差率%": r.get("价差率%", ""), "结论": r.get("结论", ""),
                    "建议": "自家价更高，可评估调价或增加赠品/优惠" if str(r.get("结论", "")) == "最贵"
                            else ("竞店价更高，价格有优势" if str(r.get("结论", "")) == "最便宜" else ""),
                })
    out = pd.DataFrame(rows)
    if len(out):
        out.to_excel(os.path.join(OUT, "商品力对比_建议.xlsx"), index=False)
    print(f"商品力对比_建议: {len(out)} 条 -> output/商品力对比_建议.xlsx")
    return out


# ===========================================================================
# 2.5 竞争分析报告
# ===========================================================================
def final_report(df, price_band_df, hot, gap, defend, shop_cat, cat_summary, matrix_df):
    n_shop = df["shop_name"].nunique()
    n_sku = len(df)
    total_sales = int(df["sales"].sum())
    # 总览
    if os.path.exists(MATRIX_FILE):
        m = pd.read_excel(MATRIX_FILE, sheet_name="价格矩阵", dtype={"自家SKUID": str})
        n_match = len(m)
    else:
        n_match = 0
    overview = pd.DataFrame({
        "指标": ["竞店数", "竞店SKU总数", "竞店月售合计", "自家有匹配SKU数",
                 "热销Top100月售合计", "缺口候选数", "防守SKU数", "数据日期"],
        "数值": [n_shop, n_sku, total_sales, n_match,
                 int(hot["月售合计"].sum()) if len(hot) else 0,
                 len(gap), len(defend), "2026-08-07"],
    })

    with pd.ExcelWriter(os.path.join(OUT, "竞争分析报告.xlsx"), engine="openpyxl") as w:
        overview.to_excel(w, sheet_name="总览", index=False)
        m = pd.read_excel(MATRIX_FILE, sheet_name="价格矩阵") if os.path.exists(MATRIX_FILE) else pd.DataFrame()
        m.to_excel(w, sheet_name="价格矩阵", index=False)
        price_band_df.to_excel(w, sheet_name="价格带", index=False)
        hot.to_excel(w, sheet_name="热销Top100", index=False)
        gap.to_excel(w, sheet_name="缺口清单", index=False)
        defend.to_excel(w, sheet_name="防守清单", index=False)
        shop_cat.to_excel(w, sheet_name="品类结构", index=False)
        cat_summary.to_excel(w, sheet_name="类目缺口", index=False)
    print(f"竞争分析报告 -> output/竞争分析报告.xlsx")


def main():
    import argparse
    parser = argparse.ArgumentParser(description="多竞店竞争分析（选店）")
    parser.add_argument("--shops", help="参与分析的竞店名（逗号分隔，默认全部店）")
    args = parser.parse_args()
    shops_filter = None
    if args.shops:
        shops_filter = set(s.strip() for s in args.shops.split(",") if s.strip())
        print(f"选店分析: {len(shops_filter)} 家店 -> {sorted(shops_filter)}")

    print("加载数据...")
    df = load_summary(shops_filter)
    own = load_own()
    print(f"汇总库: {df.shape} ({df['shop_name'].nunique()} 店) | 自家: {own.shape}")

    print("\n[2.1] 价格带分析")
    pb = price_band(df)

    print("\n[2.2] 热销/缺口/防守")
    g, hot = hot_products(df, own)
    gap = gap_list(g, own)
    defend = defend_list(own, df)

    print("\n[2.3] 品类结构")
    shop_cat, cat_summary = category_structure(df, own)

    print("\n[2.4] 商品力对比")
    pp = product_power(None, df)

    print("\n[2.5] 竞争分析报告")
    matrix_df = pd.read_excel(MATRIX_FILE, sheet_name="价格矩阵") if os.path.exists(MATRIX_FILE) else pd.DataFrame()
    final_report(df, pb, hot, gap, defend, shop_cat, cat_summary, matrix_df)
    print("\n全部完成。")


if __name__ == "__main__":
    main()

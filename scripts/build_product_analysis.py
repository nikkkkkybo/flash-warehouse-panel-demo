# -*- coding: utf-8 -*-
"""
产品分析建议（品为单位，规则版，免费全量）

以「自家品组（SPU）」为品，聚合 竞品侧（对应到的竞品：数量/总月售/品牌/价格带/规格）
vs 自家侧（月销/售价/采购价/规格数），计算差距指标，产出结论标签与规则建议。

结论标签（判定用**有效销量 = 成交 + λ×收藏**，λ 见 sales_text.INTENT_WEIGHT）：
  机会品：竞品有效销量≥30 且 自家月销 < 有效销量一半（竞品需求强自家弱）
  防守品：自家月销≥30 且 竞品有效销量≥30（双方都强，需守住）
  领先品：自家月销≥30 且 竞品有效销量<30（自家主导）
  观察品：其余
  —— 「N人想买」的收藏量不剔除，只按 λ 折算：它照样是需求信号，只是权重低于成交。
建议（规则组合）：对标品牌（跨品牌）/ 补规格（规格缺口）/ 价格对标（竞品更便宜）/
  机会提示（竞品需求强）/ 防守提示

产出 output/产品分析建议.xlsx（品分析卡 + 按类目统计）
"""
import os
import sys
from collections import defaultdict

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUTPUT = os.path.join(BASE, "output")

from build_comparison import build_comparison_data  # noqa: E402
from spu_semantic import _clean, _num  # noqa: E402
import sales_text as _STx  # noqa: E402

# 阈值
OPP_SALES = 30      # 竞品有效销量 ≥ 30 才算需求强
DEF_SALES = 30      # 自家月销 ≥ 30 才算自家强
SALES_RATIO = 0.5   # 机会：自家 < 有效销量 × 0.5
PRICE_GAP = 0.9     # 价格对标：竞品中位价 < 自家售价 × 0.9


def build_product_analysis(shop="示例竞店A", detail=None):
    """返回 (品卡 df, 类目统计 df)。detail 可传已算好的对比明细（避免重复计算）。"""
    if detail is None:
        detail, _ = build_comparison_data(shop)
    if not len(detail):
        return pd.DataFrame(), pd.DataFrame()

    own = pd.read_excel(os.path.join(OUTPUT, "自家_归一化后.xlsx", ), dtype=str).fillna("")

    # ---- 品组信息（自家，向量化）----
    own = own.copy()
    # P4.3 关键修正：总库表同一 SKUID 有 ~5 个渠道行（售价/库存不同，商品事实相同）——
    # 销量是 SKU 的事实属性，按行求和会虚高 ~5 倍 → 品级聚合前必须按 SKUID 去重
    own = own.drop_duplicates(subset=["SKUID"], keep="first")
    own["__s"] = own["月销量"].apply(_num).fillna(0.0)
    own["__p"] = pd.to_numeric(own["美团闪购渠道售价"], errors="coerce")
    own["__c"] = pd.to_numeric(own["采购价"], errors="coerce")

    g = own.groupby("SPUID")
    g_sum = g["__s"].sum()
    g_cnt = g["SKUID"].count()
    g_spec = g["规格名称"].nunique() if "规格名称" in own.columns else g["SKUID"].count()
    g_pmin = g["__p"].min()
    g_pmax = g["__p"].max()
    g_cmin = g["__c"].min()
    g_cmax = g["__c"].max()
    rep_idx = g["__s"].idxmax()   # 代表行（组内月销最高）

    spu_info = {}
    for spuid in g_sum.index:
        spuid = _clean(spuid)
        if not spuid:
            continue
        r = own.loc[rep_idx[spuid]]
        pmin, pmax = g_pmin[spuid], g_pmax[spuid]
        cmin, cmax = g_cmin[spuid], g_cmax[spuid]
        spu_info[spuid] = {
            "品名": _clean(r.get("商品名称")) or _clean(r.get("清洗后商品名")),
            "类目": _clean(r.get("类目三级_归一")) or _clean(r.get("美团闪购类目三级")),
            "自家品牌": _clean(r.get("品牌（抽取）")) or _clean(r.get("品牌名称")),
            "自家月销": float(g_sum[spuid]),
            "SKU数": int(g_cnt[spuid]),
            "规格数": int(g_spec[spuid]),
            "售价区间": (float(pmin), float(pmax)) if pmin == pmin else (None, None),
            "采购价区间": (float(cmin), float(cmax)) if cmin == cmin else (None, None),
        }
    sku2spu = dict(zip(own["SKUID"].astype(str).str.strip(), own["SPUID"].astype(str).str.strip()))

    # ---- 对比行按品组聚合 ----
    agg = defaultdict(lambda: {
        "竞品数": 0, "竞品总月售": 0.0, "竞品最高月售": 0.0, "竞品收藏量": 0.0,
        "竞品品牌": set(), "竞品价格": [], "竞品规格": set(),
        "跨品牌数": 0, "跨品牌名": set(),
    })
    for _, r in detail.iterrows():
        spu = sku2spu.get(_clean(r.get("自家SKUID")))
        if not spu or spu not in spu_info:
            continue
        a = agg[spu]
        a["竞品数"] += 1
        raw_cs = r.get("竞品月售")
        cs = _num(raw_cs) or 0.0
        a["竞品总月售"] += cs
        # 收藏量（「N人想买」）单记一份：判定按 λ 折算，展示仍保留原始口径
        if _STx.is_intent(raw_cs):
            a["竞品收藏量"] += cs
        a["竞品最高月售"] = max(a["竞品最高月售"], cs)
        cb = _clean(r.get("竞品品牌"))
        if cb and cb != "（无品牌）":
            a["竞品品牌"].add(cb)
        cp = _num(r.get("竞品到手价"))
        if cp is not None:
            a["竞品价格"].append(cp)
        csp = _clean(r.get("竞品规格"))
        if csp:
            a["竞品规格"].add(csp[:20])
        ob = _clean(r.get("自家品牌"))
        if r.get("跨品牌") == "是":
            a["跨品牌数"] += 1
            if cb and cb != "（无品牌）":
                a["跨品牌名"].add(cb)

    # ---- 品卡行 ----
    rows = []
    for spu, a in agg.items():
        info = spu_info[spu]
        own_sales = info["自家月销"]
        comp_sales = a["竞品总月售"]
        comp_prices = sorted(a["竞品价格"])
        comp_med = comp_prices[len(comp_prices) // 2] if comp_prices else None
        price_lo = comp_prices[0] if comp_prices else None
        price_hi = comp_prices[-1] if comp_prices else None
        own_price_lo, own_price_hi = info["售价区间"]

        comp_want = a["竞品收藏量"]
        # 有效销量 = 成交 + λ×收藏：收藏照样是需求信号，只是权重低于成交
        comp_eff = (comp_sales - comp_want) + _STx.INTENT_WEIGHT * comp_want

        # 结论（判定用有效销量；展示的 竞品总月售 仍是原始口径，两列都在卡上可对照）
        if comp_eff >= OPP_SALES and own_sales < comp_eff * SALES_RATIO:
            label = "机会品"
        elif own_sales >= DEF_SALES and comp_eff >= OPP_SALES:
            label = "防守品"
        elif own_sales >= DEF_SALES:
            label = "领先品"
        else:
            label = "观察品"

        # 规则建议
        advices = []
        if label == "机会品":
            advices.append("机会：竞品需求强（有效销量%.0f=成交%.0f+收藏%.0f×%g），自家月销%.0f偏低"
                           "——对标竞品热销品牌/规格"
                           % (comp_eff, comp_sales - comp_want, comp_want,
                              _STx.INTENT_WEIGHT, own_sales))
            if comp_prices:
                advices.append("参考价格带 %.2f-%.2f 元" % (price_lo, price_hi))
        if a["跨品牌数"] >= 1:
            advices.append("品牌差距：竞品品牌 %s 在卖（自家 %s），评估品牌力/引进"
                           % ("、".join(sorted(a["跨品牌名"])[:3]) or "多品牌",
                              info["自家品牌"] or "无"))
        if info["规格数"] < len(a["竞品规格"]):
            advices.append("规格缺口：竞品规格 %d 个 > 自家 %d 个，考虑补规格"
                           % (len(a["竞品规格"]), info["规格数"]))
        if comp_med is not None and own_price_lo and comp_med < own_price_lo * PRICE_GAP:
            advices.append("价格对标：竞品中位价 %.2f < 自家最低售价 %.2f，考虑定价优化"
                           % (comp_med, own_price_lo))
        if label == "防守品":
            advices.append("防守：竞品同样强势（有效销量%.0f），保持价格/规格竞争力" % comp_eff)
        if not advices:
            advices.append("保持观察")

        rows.append({
            "品组ID": spu,
            "品名": info["品名"][:36],
            "类目": info["类目"],
            "自家品牌": info["自家品牌"] or "（无）",
            "自家月销": round(own_sales, 0),
            "自家规格数": info["规格数"],
            "自家售价区间": ("%.2f-%.2f" % (info["售价区间"][0], info["售价区间"][1]))
                              if info["售价区间"][0] is not None else "",
            "对应竞品数": a["竞品数"],
            "竞品总月售": round(comp_sales, 0),
            "竞品收藏量": round(comp_want, 0),
            "竞品有效月售": round(comp_eff, 1),
            "竞品最高月售": round(a["竞品最高月售"], 0),
            "在卖品牌数": len(a["竞品品牌"]),
            "竞品品牌": "、".join(sorted(a["竞品品牌"])[:5]),
            "竞品价格带": ("%.2f-%.2f" % (price_lo, price_hi)) if comp_prices else "",
            "竞品规格数": len(a["竞品规格"]),
            "结论": label,
            "建议": "；".join(advices),
        })
    pdf = pd.DataFrame(rows)
    if len(pdf):
        rank = {"机会品": 0, "防守品": 1, "领先品": 2, "观察品": 3}
        pdf["__rk"] = pdf["结论"].map(rank)
        pdf = pdf.sort_values(["__rk", "竞品有效月售"], ascending=[True, False])
        pdf = pdf.drop(columns=["__rk"])
        pdf.insert(0, "序号", range(1, len(pdf) + 1))

    # 类目统计
    cat_rows = []
    if len(pdf):
        for cat, g in pdf.groupby("类目"):
            cat_rows.append({
                "类目": cat,
                "品数": len(g),
                "机会品": int((g["结论"] == "机会品").sum()),
                "防守品": int((g["结论"] == "防守品").sum()),
                "竞品总月售": int(g["竞品总月售"].sum()),
                "竞品收藏量": int(g["竞品收藏量"].sum()),
                "竞品有效月售": round(float(g["竞品有效月售"].sum()), 1),
                "自家月销合计": int(g["自家月销"].sum()),
            })
    cdf = pd.DataFrame(cat_rows)
    if len(cdf):
        cdf = cdf.sort_values("竞品总月售", ascending=False)
        cdf.insert(0, "序号", range(1, len(cdf) + 1))
    return pdf, cdf


def build_product_detail(shop, spu, detail=None, own_df=None):
    """品详情页数据：该品的自家SKU列表 + 对应竞品列表 + 差距指标。
    detail/own_df 可传入已加载的数据（避免重复读表）。返回 dict 或 None。"""
    if detail is None:
        detail, _ = build_comparison_data(shop)
    if own_df is None:
        own_df = pd.read_excel(os.path.join(OUTPUT, "自家_归一化后.xlsx"),
                               dtype=str).fillna("")

    own_skus = own_df[own_df["SPUID"].astype(str).str.strip() == str(spu)]
    if not len(own_skus):
        return None
    rep = own_skus.iloc[0]

    # ---- 自家 SKU 列表 ----
    own_list, own_specs = [], set()
    own_sales_total = 0.0
    for _, r in own_skus.iterrows():
        spec = _clean(r.get("规格名称"))
        if spec:
            own_specs.add(spec)
        s = _num(r.get("月销量")) or 0.0
        own_sales_total += s
        own_list.append({
            "SKUID": _clean(r.get("SKUID")),
            "规格": spec,
            "售价": _clean(r.get("美团闪购渠道售价")),
            "采购价": _clean(r.get("采购价")),
            "月销": round(s, 0),
            "图片": _clean(r.get("商品图片url")) or _clean(r.get("图片url")),
        })
    own_list.sort(key=lambda x: -_num(x["月销"]))

    # ---- 竞品列表（对比明细里自家SKUID属于该品的行）----
    spu_sku_set = set(own_skus["SKUID"].astype(str).str.strip())
    comp_rows = detail[detail["自家SKUID"].astype(str).str.strip().isin(spu_sku_set)]
    comp_list, comp_specs, comp_brands, comp_prices = [], set(), set(), []
    comp_sales_total = 0.0
    comp_want_total = 0.0
    for _, r in comp_rows.iterrows():
        raw_cs = r.get("竞品月售")
        cs = _num(raw_cs) or 0.0
        comp_sales_total += cs
        is_iv = _STx.is_intent(raw_cs)
        if is_iv:
            comp_want_total += cs
        cb = _clean(r.get("竞品品牌"))
        if cb and cb != "（无品牌）":
            comp_brands.add(cb)
        cp = _num(r.get("竞品到手价"))
        if cp is not None:
            comp_prices.append(cp)
        csp = _clean(r.get("竞品规格"))
        if csp:
            comp_specs.add(csp)
        comp_list.append({
            "竞品名": _clean(r.get("竞品商品名")),
            "品牌": cb,
            "月售": round(cs, 0),
            "意向": is_iv,
            "到手价": _clean(r.get("竞品到手价")),
            "规格": csp,
            "竞品Sku_id": _clean(r.get("竞品Sku_id")),
            "跨品牌": r.get("跨品牌") == "是",
        })
    comp_list.sort(key=lambda x: -_num(x["月售"]))

    # ---- 差距指标（判定用有效销量：收藏按 λ 折算，不剔除）----
    comp_eff_total = (comp_sales_total - comp_want_total) + _STx.INTENT_WEIGHT * comp_want_total
    spec_gap = sorted(comp_specs - own_specs)   # 竞品有、自家没有的规格
    if comp_eff_total >= 30 and own_sales_total < comp_eff_total * 0.5:
        label = "机会品"
    elif own_sales_total >= 30 and comp_eff_total >= 30:
        label = "防守品"
    elif own_sales_total >= 30:
        label = "领先品"
    else:
        label = "观察品"
    sales_ratio = (comp_eff_total / own_sales_total) if own_sales_total > 0 else None

    # 自家售价区间 + 价格信号（竞品中位价 vs 自家最低售价）
    own_prices = []
    for x in own_list:
        p = _num(x["售价"])
        if p is not None:
            own_prices.append(p)
    own_price_range = (min(own_prices), max(own_prices)) if own_prices else (None, None)
    comp_med = sorted(comp_prices)[len(comp_prices) // 2] if comp_prices else None
    price_signal = ""
    if comp_med is not None and own_price_range[0]:
        if comp_med < own_price_range[0] * 0.9:
            price_signal = "竞品更便宜 %.0f%%" % ((own_price_range[0] - comp_med) / own_price_range[0] * 100)
        elif comp_med > own_price_range[1] * 1.1:
            price_signal = "自家更便宜 %.0f%%" % ((comp_med - own_price_range[1]) / comp_med * 100)

    return {
        "spu": spu,
        "品名": _clean(rep.get("商品名称")) or _clean(rep.get("清洗后商品名")),
        "类目": _clean(rep.get("类目三级_归一")) or _clean(rep.get("美团闪购类目三级")),
        "自家品牌": _clean(rep.get("品牌（抽取）")) or _clean(rep.get("品牌名称")) or "（无）",
        "结论": label,
        "自家SKU": own_list,
        "自家月销合计": round(own_sales_total, 0),
        "自家规格数": len(own_specs),
        "自家售价区间": own_price_range,
        "价格信号": price_signal,
        "竞品": comp_list,
        "竞品数": len(comp_list),
        "竞品总月售": round(comp_sales_total, 0),
        "竞品收藏量": round(comp_want_total, 0),
        "竞品有效月售": round(comp_eff_total, 1),
        "在卖品牌": sorted(comp_brands),
        "竞品价格带": (min(comp_prices), max(comp_prices)) if comp_prices else (None, None),
        "销量比": round(sales_ratio, 1) if sales_ratio else None,
        "规格缺口": spec_gap,
    }


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--shop", default="示例竞店A")
    args = ap.parse_args()
    pdf, cdf = build_product_analysis(args.shop)
    out = os.path.join(OUTPUT, "产品分析建议.xlsx")
    with pd.ExcelWriter(out) as w:
        pdf.to_excel(w, sheet_name="品分析卡", index=False)
        cdf.to_excel(w, sheet_name="类目统计", index=False)
    print("产出:", out)
    print("  品卡 %d 个 | 类目 %d" % (len(pdf), len(cdf)))
    if len(pdf):
        print("\n结论分布:", pdf["结论"].value_counts().to_dict())
        print("\n机会品 Top10:")
        opp = pdf[pdf["结论"] == "机会品"].head(10)
        print(opp[["品名", "自家月销", "竞品总月售", "在卖品牌数", "建议"]].to_string(index=False, max_colwidth=24))


if __name__ == "__main__":
    main()

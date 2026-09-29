# -*- coding: utf-8 -*-
"""
同品对比视图（平台初衷落地）：同一品节点的 竞品 vs 自家 对比

输入：试跑目录 人工池（品级语义候选+其他候选）+ 汇总库（自动确认）+ 竞品表 + 自家表
产出 output/同品对比.xlsx：
  - 明细：每对对应（竞品侧 品牌/月售/到手价/规格 vs 自家侧 品牌/月销/售价/采购价/规格
          + 价差率/月售比/是否跨品牌/结论）
  - 品聚合：按「类目×宽品名×细品名词」聚合（竞品数/竞品总月售/品牌数/自家月销/机会判断）

结论规则（聚焦"看好不看坏"）：
  跨品牌 + 竞品月售显著高于自家月销  -> 机会：竞品品牌卖得好，考虑引进/对标
  同品竞品月售高                      -> 参考：规格/价格差异分析
  竞品月售低                          -> 低价值参考
"""
import os
import sys

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUTPUT = os.path.join(BASE, "output")

from spu_semantic import _clean, _num  # noqa: E402


def _read(path, **kw):
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, dtype=str, **kw)
    except Exception:
        return pd.DataFrame()


def build_comparison_data(shop="示例竞店A"):
    """计算同品对比数据。返回 (detail_df, agg_df)，供 Excel 导出与 Web 面板共用。"""
    trial = os.path.join(OUTPUT, "试跑_" + shop)

    # 1) 对应关系：汇总库 SKU 层（自动确认）+ 人工池候选
    pairs = []  # {竞品Sku_id, 竞品商品名, 竞品品牌, 自家SKUID, 自家商品名, 匹配, 置信, 语义层级}
    lib = _read(os.path.join(OUTPUT, "对齐汇总库.xlsx"), sheet_name="SKU匹配层")
    if len(lib):
        for _, r in lib[lib["竞店名称"] == shop].iterrows():
            pairs.append({
                "竞品Sku_id": _clean(r.get("竞品Sku_id")),
                "竞品商品名": _clean(r.get("竞品商品名")),
                "竞品品牌": _clean(r.get("竞品品牌")),
                "自家SKUID": _clean(r.get("自家SKUID")),
                "自家商品名": _clean(r.get("自家商品名")),
                "匹配": "自动确认(" + _clean(r.get("匹配层")) + ")",
                "置信度": _clean(r.get("置信度")),
                "语义层级": "",
            })
    pool = _read(os.path.join(trial, "对齐_人工确认池.xlsx"))
    if len(pool):
        for _, r in pool.iterrows():
            pairs.append({
                "竞品Sku_id": _clean(r.get("竞品Sku_id")),
                "竞品商品名": _clean(r.get("竞品商品名")),
                "竞品品牌": _clean(r.get("竞品品牌")),
                "自家SKUID": _clean(r.get("自家SKUID")),
                "自家商品名": _clean(r.get("自家商品名")),
                "匹配": _clean(r.get("匹配层")),
                "置信度": _clean(r.get("置信度")),
                "语义层级": _clean(r.get("语义层级")),
            })
    pdf = pd.DataFrame(pairs).drop_duplicates("竞品Sku_id", keep="first")

    # 2) 竞品侧信息（月售/到手价/规格/类目/品牌）
    comp = _read(os.path.join(trial, "竞品_归一化后.xlsx"))
    comp_info = {}
    if len(comp):
        sku_col = "sku_id" if "sku_id" in comp.columns else "Sku_id"
        brand_col = "品牌（抽取）" if "品牌（抽取）" in comp.columns else "品牌名称"
        for _, r in comp.iterrows():
            sid = _clean(r.get(sku_col))
            if sid:
                comp_info[sid] = {
                    "竞品月售": _clean(r.get("sales") or r.get("月销量") or r.get("月售")),
                    "竞品到手价": _clean(r.get("price") or r.get("到手价") or r.get("折后价")),
                    "竞品规格": _clean(r.get("spec") or r.get("规格")),
                    "竞品类目": _clean(r.get("类目三级_归一") or r.get("category3")),
                    "竞品品牌": _clean(r.get(brand_col)),
                }
    pdf["竞品月售"] = pdf["竞品Sku_id"].map(lambda s: comp_info.get(s, {}).get("竞品月售", ""))
    pdf["竞品到手价"] = pdf["竞品Sku_id"].map(lambda s: comp_info.get(s, {}).get("竞品到手价", ""))
    pdf["竞品规格"] = pdf["竞品Sku_id"].map(lambda s: comp_info.get(s, {}).get("竞品规格", ""))
    pdf["竞品类目"] = pdf["竞品Sku_id"].map(lambda s: comp_info.get(s, {}).get("竞品类目", ""))
    # 竞品品牌补全：汇总库行无品牌列，从竞品归一化表补（池行已有则保留）
    pdf["__brand_comp"] = pdf["竞品Sku_id"].map(lambda s: comp_info.get(s, {}).get("竞品品牌", ""))
    pdf["竞品品牌"] = pdf.apply(
        lambda r: _clean(r.get("竞品品牌")) or r.get("__brand_comp", ""), axis=1)
    pdf = pdf.drop(columns=["__brand_comp"])

    # 3) 自家侧信息（月销/售价/采购价/规格/品牌）
    own = _read(os.path.join(OUTPUT, "自家_归一化后.xlsx"))
    own_info = {}
    if len(own):
        for _, r in own.iterrows():
            sid = _clean(r.get("SKUID"))
            if sid:
                own_info[sid] = {
                    "自家月销": _clean(r.get("月销量") or r.get("月售")),
                    "自家售价": _clean(r.get("美团闪购渠道售价") or r.get("售价")),
                    "自家采购价": _clean(r.get("采购价")),
                    "自家规格": _clean(r.get("规格名称") or r.get("规格")),
                    "自家品牌": _clean(r.get("品牌（抽取）")) or _clean(r.get("品牌名称")),
                }
    pdf["自家月销"] = pdf["自家SKUID"].map(lambda s: own_info.get(s, {}).get("自家月销", ""))
    pdf["自家售价"] = pdf["自家SKUID"].map(lambda s: own_info.get(s, {}).get("自家售价", ""))
    pdf["自家采购价"] = pdf["自家SKUID"].map(lambda s: own_info.get(s, {}).get("自家采购价", ""))
    pdf["自家规格"] = pdf["自家SKUID"].map(lambda s: own_info.get(s, {}).get("自家规格", ""))
    pdf["自家品牌"] = pdf["自家SKUID"].map(lambda s: own_info.get(s, {}).get("自家品牌", ""))

    # 4) 计算指标
    def _ratio(a, b):
        na, nb = _num(a), _num(b)
        if na is None or nb is None or nb == 0:
            return ""
        return round((na - nb) / nb, 3)

    pdf["价差率"] = [_ratio(a, b) for a, b in zip(pdf["竞品到手价"], pdf["自家售价"])]
    pdf["月售比"] = [_ratio(a, b) for a, b in zip(pdf["竞品月售"], pdf["自家月销"])]

    # 销量差距（重点维度）：竞品月售 vs 自家月销
    def _sales_gap(r):
        cs = _num(r.get("竞品月售"))
        os_ = _num(r.get("自家月销"))
        if cs is None:
            return ""
        if os_ is None or os_ <= 0:
            return "竞品在卖·自家无销量" if cs >= 6 else ""
        if cs >= os_ * 1.5:
            return "竞品更高 %.0fx" % (cs / os_)
        if cs > os_:
            return "竞品略高"
        return "自家持平/更高"

    pdf["销量差距"] = pdf.apply(_sales_gap, axis=1)

    pdf["跨品牌"] = pdf.apply(
        lambda r: ("是" if (_clean(r.get("竞品品牌")) and _clean(r.get("自家品牌"))
                           and _clean(r.get("竞品品牌")) != _clean(r.get("自家品牌"))) else ""),
        axis=1)
    # 空品牌显示（无品牌），便于阅读
    pdf["竞品品牌"] = pdf["竞品品牌"].apply(lambda v: _clean(v) or "（无品牌）")
    pdf["自家品牌"] = pdf["自家品牌"].apply(lambda v: _clean(v) or "（无品牌）")

    def _conclusion(r):
        cs, os_ = _num(r.get("竞品月售")), _num(r.get("自家月销"))
        if cs is None:
            return "无销量数据"
        cross = r.get("跨品牌") == "是"
        gap = str(r.get("销量差距") or "")
        if gap.startswith("竞品在卖·自家无销量"):
            return "重点：竞品在卖自家无销量（对照引进）" if cross else "重点：竞品在卖自家无销量"
        if gap.startswith("竞品更高"):
            return "重点：竞品销量高出自家" + ("（跨品牌）" if cross else "")
        if cross and os_ is not None and cs > os_:
            return "竞品更受欢迎（跨品牌）"
        if cs < 6:
            return "低价值参考"
        return "观察"

    pdf["结论"] = pdf.apply(_conclusion, axis=1)

    # 排序：跨品牌优先 → 销量差距等级 → 月售降序
    def _gap_rank(r):
        g = str(r.get("销量差距") or "")
        if g.startswith("竞品在卖·自家无销量"):
            return 3
        if g.startswith("竞品更高"):
            return 2
        if g == "竞品略高":
            return 1
        return 0

    pdf["__gr"] = pdf.apply(_gap_rank, axis=1)
    pdf["__cs"] = pdf["竞品月售"].apply(_num).fillna(-1)
    pdf = pdf.sort_values(["跨品牌", "__gr", "__cs"], ascending=[False, False, False])
    pdf = pdf.drop(columns=["__gr", "__cs"])
    pdf.insert(0, "序号", range(1, len(pdf) + 1))

    # 5) 品聚合（按 类目×宽品名×细品名词——从语义层级/竞品类目近似）
    agg_rows = []
    by_cat = pdf.groupby("竞品类目")
    for cat, g in by_cat:
        if not cat:
            continue
        n_comp = len(g)
        total_sales = sum(_num(v) or 0 for v in g["竞品月售"])
        max_sales = max((_num(v) or 0) for v in g["竞品月售"])
        brands = set(_clean(b) for b in g["竞品品牌"]
                     if _clean(b) and _clean(b) != "（无品牌）")
        own_total = sum(_num(v) or 0 for v in g["自家月销"])
        cross_n = int((g["跨品牌"] == "是").sum())
        if total_sales >= 30 and (own_total < total_sales * 0.5 or cross_n >= 2):
            judge = "机会品（竞品需求强/多品牌在卖）"
        elif total_sales >= 30:
            judge = "热销品（自家有对应）"
        else:
            judge = "观察"
        agg_rows.append({
            "竞品类目": cat,
            "对应竞品数": n_comp,
            "竞品总月售": round(total_sales, 0),
            "竞品最高月售": round(max_sales, 0),
            "在卖品牌数": len(brands),
            "品牌": "、".join(sorted(brands)[:5]),
            "自家对应月销合计": round(own_total, 0),
            "跨品牌对应数": cross_n,
            "判断": judge,
        })
    adf = pd.DataFrame(agg_rows)
    if len(adf):
        adf = adf.sort_values("竞品总月售", ascending=False)
        adf.insert(0, "序号", range(1, len(adf) + 1))
    return pdf, adf


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--shop", default="示例竞店A")
    args = ap.parse_args()
    pdf, adf = build_comparison_data(args.shop)
    print("对应关系:", len(pdf), "条")
    out = os.path.join(OUTPUT, "同品对比.xlsx")
    with pd.ExcelWriter(out) as w:
        pdf.to_excel(w, sheet_name="对比明细", index=False)
        adf.to_excel(w, sheet_name="品聚合(按类目)", index=False)
    print("产出:", out)
    print("  对比明细 %d 行 | 品聚合 %d 类目" % (len(pdf), len(adf)))
    print("\n跨品牌机会 Top10:")
    opp = pdf[pdf["跨品牌"] == "是"].head(10)
    if len(opp):
        print(opp[["竞品商品名", "竞品品牌", "竞品月售", "自家品牌", "自家月销", "结论"]].to_string(index=False))
    print("\n品聚合 Top10:")
    print(adf.head(10)[["竞品类目", "对应竞品数", "竞品总月售", "在卖品牌数", "自家对应月销合计", "判断"]].to_string(index=False))


if __name__ == "__main__":
    main()

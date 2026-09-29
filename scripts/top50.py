# -*- coding: utf-8 -*-
"""
竞品销量前50品（商品ID聚合，同规格归组）—— 人工匹配工作台数据源

要点（2026-08-14 确认）：
- 以竞品 商品ID(comp_id) 为"品"，同商品下各规格(SKU)归为一组；
- 实测同商品各规格「月售」完全相同（产品级销量重复）→ 品月售取 max，不重复计；
- 品名/品牌/价格取组内月售最高规格为代表；SKU列表保留全部规格供对照；
- 排序按品月售降序取前 N（默认50）。
"""
import os

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    """销量文本 → 数值：委托 `sales_text.num`（认 万/千/k 单位、'月售2'/'8人想买'/'1,200'）。"""
    import sales_text
    return sales_text.num(v)


def _is_intent(v):
    """这条销量是不是意向类（非真实成交）：委托 `sales_text.is_intent`。"""
    import sales_text
    return sales_text.is_intent(v)


def _sales_rank_key(item):
    """热销榜排序键：按 有效销量 降序（数字取负使正序排序时大的在前）。

    `意向` 只作为次级键（同分时成交在前）与展示标记——收藏不再被整档压到所有成交之后，
    而是在 有效销量 里按 λ 折算后参与同一次排序。
    """
    tier = 1 if item.get("意向") else 0
    return (-(item.get("有效销量") if item.get("有效销量") is not None
              else item.get("月售", 0)), tier)


def build_top50(shop, n=50):
    """竞品销量前 n 个品（商品ID聚合）。返回 list[dict]（含 SKU列表）。

    排序按**有效销量 = 成交 + λ × 收藏**（λ 见 sales_text.INTENT_WEIGHT）：
    「N人想买」照样是需求证据，只是按权重降级——不再一刀切压到所有成交之后。
    """
    import sales_text as _STx
    f = os.path.join(OUTPUT, f"试跑_{shop}", "竞品_归一化后.xlsx")
    if not os.path.exists(f):
        return []
    comp = pd.read_excel(f, dtype=str).fillna("")
    if not len(comp):
        return []
    comp["__s"] = comp["sales"].apply(_num)
    comp["__intent"] = comp["sales"].apply(_is_intent)
    # 意向行的数量（成交行记 0）：以前只留 True/False 标志，"多少人在想买"这个数丢了，
    # 「意向虚高」探测要的是数量，没有它就永远出不了结论。
    comp["__iv"] = comp["__s"].where(comp["__intent"], 0.0)
    comp["__eff"] = comp["__s"] * comp["__intent"].map(lambda x: _STx.INTENT_WEIGHT if x else 1.0)
    # 排序键：有效销量降序（越小越靠前 → 取负）
    comp["__rank"] = -comp["__eff"]
    img_col = "image_url" if "image_url" in comp.columns else ("图片url" if "图片url" in comp.columns else None)
    products = []
    for pid, g in comp.groupby("comp_id"):
        best = g.loc[g["__rank"].idxmin()]
        skus = [{
            "sku_id": _clean(r.get("sku_id")),
            "规格": _clean(r.get("spec")),
            "月售": round(float(r["__s"]), 0),
            "意向": bool(r["__intent"]),
            "意向数": round(float(r["__iv"]), 0),
            "有效销量": round(float(r["__eff"]), 1),
            "价格": _clean(r.get("price")) or _clean(r.get("discount_price")),
            "原价": _clean(r.get("original_price")) or _clean(r.get("原价")),
            "条码": _clean(r.get("barcode")),
            "图片": _clean(r.get(img_col)) if img_col else "",
        } for _, r in g.iterrows()]
        skus.sort(key=_sales_rank_key)
        products.append({
            "comp_id": str(pid),
            "品名": _clean(best.get("name")) or _clean(best.get("清洗后商品名")),
            "品牌": _clean(best.get("品牌（抽取）")) or _clean(best.get("品牌名称")),
            "类目": _clean(best.get("类目三级_归一")) or _clean(best.get("category3"))
                    or _clean(best.get("category2")) or _clean(best.get("category1")),
            "月售": round(float(best["__s"]), 0),
            "意向": bool(best["__intent"]),
            "意向数": round(float(g["__iv"].max()), 0),
            "有效销量": round(float(best["__eff"]), 1),
            "价格": _clean(best.get("price")) or _clean(best.get("discount_price")),
            "原价": _clean(best.get("original_price")) or _clean(best.get("原价")),
            "规格数": len(g),
            "图片": _clean(best.get(img_col)) if img_col else "",
            "SKU列表": skus,
        })
    products.sort(key=_sales_rank_key)
    return products[:n]


def health_check(shop, sample=20):
    """竞品销量数据健康度检查：返回 (总行数, 销量>0行数, 样例 raw 值)。
    用于导入后立刻反馈：sales 字段是否被正确归一化（如新店月售列映射错误→全部 0）。"""
    f = os.path.join(OUTPUT, f"试跑_{shop}", "竞品_归一化后.xlsx")
    if not os.path.exists(f):
        return {"ok": False, "msg": f"未找到归一化文件 {f}", "total": 0, "has_sales": 0, "samples": []}
    try:
        comp = pd.read_excel(f, dtype=str).fillna("")
    except Exception as e:
        return {"ok": False, "msg": "读取失败: " + str(e), "total": 0, "has_sales": 0, "samples": []}
    if not len(comp):
        return {"ok": False, "msg": "归一化文件为空", "total": 0, "has_sales": 0, "samples": []}
    sales_raw = comp["sales"].tolist() if "sales" in comp.columns else []
    parsed = [_num(x) for x in sales_raw]
    has_sales = sum(1 for x in parsed if x > 0)
    samples = [(s[:24], p) for s, p in list(zip(sales_raw, parsed))[:sample]]
    return {
        "ok": True,
        "total": len(comp),
        "has_sales": has_sales,
        "ratio": round(has_sales / len(comp), 3) if len(comp) else 0,
        "samples": samples,
        "msg": "" if has_sales > 0 else "全部 sales 为 0，可能映射或归一化异常",
    }

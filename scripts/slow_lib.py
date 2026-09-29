# -*- coding: utf-8 -*-
"""
⑤ 滞销榜：以自家门店为主体，占压资金排序 + 竞品交叉 + 处置建议
====================================================================
口径（设计确认）：
- 主体：自家某门店在售品（未选门店 = 总库全部在售品）；
- 排序：占压资金 = Σ(库存 × 采购价) 降序（比按月销升序更指向决策痛感）；
- 分层：零动销（月销=0）/ 慢动销（月销>0 但 < 品类中位数×0.2，用品类内相对水平，
  避免冤枉"本来就长尾"的五金小件）；
- 竞品交叉三档（"竞品没这个品"无法从数据可靠推断，并入未对齐提示）：
    ① 已对应+竞品热卖（竞品月售≥30）→ 选品对、页面/价格问题：优化图文 或 调价；
    ② 已对应+竞品也冷 → 选品错：下架/清仓退出；
    ③ 未对应 → 数据盲区：先去对齐（链接到详情/工作台）。
- 库存取实物库存优先，缺失回退渠道库存。
"""
import pandas as pd


def _num(v):
    try:
        f = float(str(v).replace(",", "").replace("¥", "").strip() or 0)
        return f if f == f else 0.0
    except (ValueError, TypeError):
        return 0.0


def build_slow_rank(own_df, spu_set=None, comp_lookup=None, lib_sku_df=None,
                    hot_threshold=30, top_n=300):
    """构建滞销榜。

    own_df      : 自家总库（dtype=str）
    spu_set     : 门店在售 SPUID 集合（None=不限门店）
    comp_lookup : {店: {竞品sku: row}}（竞品月售/价格）
    lib_sku_df  : 汇总库 SKU 层 DataFrame（自家SKUID → 竞品对应）
    返回 {"items": [...], "stats": {...}}
    """
    if own_df is None or not len(own_df):
        return {"items": [], "stats": {"total": 0}}

    df = own_df
    # 在售过滤（售卖状态含"在售"或为空都算，防脏数据全滤掉）
    if "售卖状态" in df.columns:
        m = df["售卖状态"].astype(str).str.contains("在售", na=False)
        if m.any():
            df = df[m]
    # 门店在售过滤
    if spu_set is not None:
        spid = df["SPUID"].astype(str).str.strip()
        df = df[spid.isin(spu_set)]
    if not len(df):
        return {"items": [], "stats": {"total": 0}}

    d = pd.DataFrame({
        "SPUID": df["SPUID"].astype(str).str.strip(),
        "SKUID": df["SKUID"].astype(str).str.strip(),
        "名称": df.get("商品名称", pd.Series([""] * len(df), index=df.index)),
        "规格": df.get("规格名称", pd.Series([""] * len(df), index=df.index)),
        "品类": df.get("类目三级_归一", pd.Series([""] * len(df), index=df.index)).fillna(""),
        "月销": df.get("月销量", pd.Series(["0"] * len(df), index=df.index)).apply(_num),
        "售价": df.get("美团闪购渠道售价", pd.Series(["0"] * len(df), index=df.index)).apply(_num),
        "采购价": df.get("采购价", pd.Series(["0"] * len(df), index=df.index)).apply(_num),
        "库存": df.get("实物库存", pd.Series(["0"] * len(df), index=df.index)).apply(_num),
    })
    # 库存：实物缺失(0)回退渠道库存
    if "渠道库存" in df.columns:
        ch = df["渠道库存"].apply(_num)
        d.loc[d["库存"] <= 0, "库存"] = ch[d["库存"] <= 0]
    if "美团闪购类目三级" in df.columns:
        d.loc[d["品类"] == "", "品类"] = df.loc[d["品类"] == "", "美团闪购类目三级"].fillna("")
    d["品类"] = d["品类"].replace("", "未分类")
    d["占压"] = (d["库存"] * d["采购价"]).round(1)

    # P4.3 修正：总库渠道重复行——月销按 SKUID 去重（销量是 SKU 事实属性，
    # 按行求和虚高 ~5x 会漏判滞销）；库存/占压保留渠道行求和（渠道库存是渠道维度，
    # 合计=全渠道占压资金，语义正确）
    d_dedup = d.drop_duplicates(subset=["SKUID"], keep="first")
    med = d_dedup.groupby("品类")["月销"].median()
    sales_by_spu = d_dedup.groupby("SPUID")["月销"].sum()

    # 竞品反查：自家SKUID → [(店, 竞品sku)]
    own2comp = {}
    if lib_sku_df is not None and len(lib_sku_df):
        m = lib_sku_df[lib_sku_df["状态"].astype(str) == "有效"]
        for os_, shop, cs in zip(m["自家SKUID"].astype(str).str.strip(),
                                  m["竞店名称"].astype(str),
                                  m["竞品Sku_id"].astype(str).str.strip()):
            if os_:
                own2comp.setdefault(os_, []).append((shop, cs))

    def _comp_cross(skuids):
        """竞品交叉：返回 (档位, 竞品月售max, 竞品最低价, 竞品店数)。"""
        best_sales, lowest_price, shops_n, found = 0.0, None, set(), 0
        for sk in skuids:
            for shop, cs in own2comp.get(sk, []):
                row = (comp_lookup or {}).get(shop, {}).get(cs)
                if row is None:
                    continue
                found += 1
                shops_n.add(shop)
                s = _num(row.get("sales") or row.get("月销量") or row.get("月售"))
                best_sales = max(best_sales, s)
                p = _num(row.get("price") or row.get("折后价"))
                if p > 0:
                    lowest_price = p if lowest_price is None else min(lowest_price, p)
        if not found:
            return "未对应", 0.0, None, 0
        if best_sales >= hot_threshold:
            return "竞品热卖", best_sales, lowest_price, len(shops_n)
        return "竞品也冷", best_sales, lowest_price, len(shops_n)

    items = []
    for spu, g in d.groupby("SPUID"):
        sales_sum = float(sales_by_spu.get(spu, 0.0))  # 去重口径（渠道重复行已剔）
        capital = float(g["占压"].sum())
        median_cat = float(med.get(g["品类"].iloc[0], 0.0))
        if sales_sum == 0:
            tier = "零动销"
        elif sales_sum < median_cat * 0.2:
            tier = "慢动销"
        else:
            continue  # 动销正常，不入榜
        cross, comp_sales, comp_price, comp_shops = _comp_cross(g["SKUID"].tolist())
        if cross == "竞品热卖":
            advice, advice_cls = "选品对：优化图片标题 或 跟价（竞品热卖说明有需求）", "adv-hot"
        elif cross == "竞品也冷":
            advice, advice_cls = "选品存疑：下架观察 或 清仓退出（竞品也卖不动）", "adv-cold"
        else:
            advice, advice_cls = "数据盲区：先到竞品库/工作台完成对应，再定处置", "adv-unknown"
        price = float(g.loc[g["售价"] > 0, "售价"].iloc[0]) if (g["售价"] > 0).any() else 0.0
        items.append({
            "SPUID": spu,
            "品名": str(g["名称"].iloc[0] or ""),
            "品类": str(g["品类"].iloc[0]),
            "规格数": len(g),
            "月销": round(sales_sum, 0),
            "售价": price,
            "库存": float(g["库存"].sum()),
            "占压资金": round(capital, 0),
            "层级": tier,
            "竞品档位": cross,
            "竞品月售": round(comp_sales, 0),
            "竞品最低价": comp_price,
            "竞品店数": comp_shops,
            "建议": advice,
            "建议样式": advice_cls,
            "价差提示": round(price - comp_price, 1)
            if (comp_price and price) else None,
        })
    items.sort(key=lambda x: (-x["占压资金"], x["月销"]))
    stats = {
        "total": len(d),
        "zero": sum(1 for i in items if i["层级"] == "零动销"),
        "slow": sum(1 for i in items if i["层级"] == "慢动销"),
        "capital": round(sum(i["占压资金"] for i in items), 0),
        "hot": sum(1 for i in items if i["竞品档位"] == "竞品热卖"),
        "cold": sum(1 for i in items if i["竞品档位"] == "竞品也冷"),
        "unknown": sum(1 for i in items if i["竞品档位"] == "未对应"),
    }
    return {"items": items[:top_n], "stats": stats}

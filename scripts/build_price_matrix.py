# -*- coding: utf-8 -*-
"""
竞品价格矩阵构建（竞争分析中枢表）

行 = 自家 SKU（与至少一家竞店有匹配）
列 = 每家竞店的 到手价 / 月售 / 是否有货
末列 = 自家售价 / 采购价 / 毛利 / 价差率 / 结论（最贵/居中/最便宜）

匹配来源分三层：
  1. 条码层（全 19 店覆盖）：自家 EAN 条码 与 各店 EAN 条码 精确匹配
  2. 深层（已试跑店覆盖）：output/试跑_<店>/ 的对齐 2/3/4 层结果
  3. 纠错回写：output/对齐修正记录.xlsx（用户修正过的匹配优先于自动匹配）

用法：
  python scripts/build_price_matrix.py

产出：output/竞品价格矩阵.xlsx（4 个 sheet：价格矩阵/竞店独有/条码对齐明细/纠错回写）
"""
import os
import sys
import glob

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from import_adapter import CN_STD
import sales_text as _ST

LIB_DIR = os.path.join(BASE, "output", "竞店标准库")
OWN_FILE = os.path.join(BASE, "output", "自家_标准版本_对齐增强.xlsx")
FIX_FILE = os.path.join(BASE, "output", "对齐修正记录.xlsx")
OUT_FILE = os.path.join(BASE, "output", "竞品价格矩阵.xlsx")
TRIAL_PREFIX = os.path.join(BASE, "output", "试跑_")

# 竞店独有热销判定：月售 >= 该阈值 且 自家无匹配
HOT_SALES = 30
# 参与深层对齐的试跑目录名 -> 店名
def trial_shops():
    return {d: None for d in sorted(os.listdir(BASE + os.sep + "output"))
            if d.startswith("试跑_")}


def load_own():
    """自家增强表：商品编号/单品编号/真实条码/售价/采购价/名称/类目。"""
    df = pd.read_excel(OWN_FILE, dtype={"商品编号": str, "单品编号": str, "商品条码": str,
                                        "真实条码(13位)": str, "条码类型": str})
    # 兼容新中文列名与旧源表列名（SPUID/SKUID）
    if "单品编号" in df.columns:
        df = df.rename(columns={"商品编号": "SPUID", "单品编号": "SKUID"})
    if "真实条码(13位)" in df.columns:
        df = df.rename(columns={"真实条码(13位)": "EAN条码"})
    for c in ["SPUID", "SKUID", "EAN条码"]:
        df[c] = df[c].astype(str).str.strip()
    return df


def load_shops(shops_filter=None):
    """读取所有竞店标准表，返回 {店名: df}（barcode 列去重索引用）。
    shops_filter: 选店集合（None=全部）。"""
    shops = {}
    for f in sorted(glob.glob(os.path.join(LIB_DIR, "竞品_*_标准化.xlsx"))):
        shop = os.path.basename(f).replace("竞品_", "").replace("_标准化.xlsx", "")
        if shops_filter is not None and shop not in shops_filter:
            continue
        df = pd.read_excel(f, dtype={"竞品商品编号": str, "竞品单品编号": str, "商品条码": str})
        df = df.rename(columns={k: v for k, v in CN_STD.items() if k in df.columns})
        for c in ["comp_id", "sku_id", "barcode"]:
            df[c] = df[c].astype(str).str.strip()
        df["barcode"] = df["barcode"].where(df["barcode"].str.len() == 13, "")
        shops[shop] = df
    return shops


def load_fixes():
    """纠错回写：对齐修正记录（若有）。返回 { (自家单品编号, 店名): 竞品单品编号 }。"""
    fixes = {}
    if os.path.exists(FIX_FILE):
        fx = pd.read_excel(FIX_FILE, dtype=str)
        # 兼容新旧列名
        own_col = "自家单品编号" if "自家单品编号" in fx.columns else ("自家SKUID" if "自家SKUID" in fx.columns else None)
        comp_col = "竞品单品编号" if "竞品单品编号" in fx.columns else ("竞品Sku_id" if "竞品Sku_id" in fx.columns else None)
        if own_col and comp_col:
            for _, r in fx.iterrows():
                if pd.notna(r.get(comp_col)):
                    fixes[(str(r[own_col]).strip(), str(r.get("店名", "")).strip())] = str(r[comp_col]).strip()
    return fixes


def load_trial_layers(shop):
    """读取某店的试跑对齐 2/3/4 层结果：{自家SKUID: [(竞品Sku_id, 匹配层)]}。"""
    d = os.path.join(BASE, "output", f"试跑_{shop}")
    if not os.path.isdir(d):
        return {}
    out = {}
    for layer, fn in [("第2层", "对齐_第2层_强组合.xlsx"), ("第3层", "对齐_第3层_弱组合_人工池.xlsx"),
                      ("第4层", "对齐_第4层_模糊_人工池.xlsx")]:
        p = os.path.join(d, fn)
        if not os.path.exists(p):
            continue
        df = pd.read_excel(p, dtype={"竞品Sku_id": str, "自家SKUID": str})
        for _, r in df.iterrows():
            key = str(r["自家SKUID"]).strip()
            cid = str(r["竞品Sku_id"]).strip()
            out.setdefault(key, []).append((cid, layer))
    return out


def build_matrix(shops_filter=None):
    own = load_own()
    shops = load_shops(shops_filter)
    fixes = load_fixes()

    # 自家 EAN 索引：EAN条码 -> [行号]
    own_ean_idx = {}
    for i, r in own.iterrows():
        ean = str(r["EAN条码"]).strip()
        if ean and ean != "nan":
            own_ean_idx.setdefault(ean, []).append(i)
    # 自家 SKUID -> 行号
    own_sku_idx = {r["SKUID"]: i for i, r in own.iterrows()}

    # 各店索引：EAN -> [sku_id]；sku_id -> 该行
    shop_data = {}
    for shop, df in shops.items():
        ean_map, sku_map = {}, {}
        for _, r in df.iterrows():
            b = r["barcode"]
            if b and b != "nan":
                ean_map.setdefault(b, []).append(r["sku_id"])
            sku_map[r["sku_id"]] = r
        trial = load_trial_layers(shop)
        shop_data[shop] = {"df": df, "ean_map": ean_map, "sku_map": sku_map, "trial": trial}

    # 构建矩阵行：{自家SKUID: {店名: {price, sales, has, layer}}}
    matrix = {}
    for ean, own_idx_list in own_ean_idx.items():
        if ean in ("", "nan"):
            continue
        for shop, sd in shop_data.items():
            if ean in sd["ean_map"]:
                for comp_sku in sd["ean_map"][ean]:
                    for oi in own_idx_list:
                        r = own.iloc[oi]
                        key = r["SKUID"]
                        cell = matrix.setdefault(key, {})
                        cell.setdefault(shop, {"layer": "第1层", "comp_sku": comp_sku})
    # 深层（试跑店）
    for shop, sd in shop_data.items():
        for own_sku, pairs in sd["trial"].items():
            if own_sku not in own_sku_idx:
                continue
            for comp_sku, layer in pairs:
                cell = matrix.setdefault(own_sku, {})
                if shop not in cell or cell[shop]["layer"] != "第1层":
                    cell[shop] = {"layer": layer, "comp_sku": comp_sku}
    # 纠错回写覆盖
    for (own_sku, shop), comp_sku in fixes.items():
        if own_sku in own_sku_idx and shop in shop_data:
            matrix.setdefault(own_sku, {})[shop] = {"layer": "人工修正", "comp_sku": comp_sku}

    # 组装输出行
    rows = []
    for own_sku, cells in matrix.items():
        if own_sku not in own_sku_idx:
            continue
        r = own.iloc[own_sku_idx[own_sku]]
        row = {"自家商品编号": r["SPUID"], "自家单品编号": own_sku,
               "自家商品名": r.get("商品名称", ""), "类目": r.get("美团闪购类目三级", "")}
        prices = []
        for shop, sd in shop_data.items():
            cell = cells.get(shop)
            if cell:
                comp_sku = cell["comp_sku"]
                cr = sd["sku_map"].get(comp_sku)
                price = cr["price"] if cr is not None and pd.notna(cr.get("price")) else ""
                sales = cr["sales"] if cr is not None and pd.notna(cr.get("sales")) else ""
                row[f"{shop}_到手价"] = price
                row[f"{shop}_月售"] = sales
                row[f"{shop}_有货"] = "是"
                row[f"{shop}_匹配层"] = cell["layer"]
                if pd.notna(price):
                    try:
                        prices.append(float(price))
                    except (TypeError, ValueError):
                        pass
            else:
                row[f"{shop}_到手价"] = ""
                row[f"{shop}_月售"] = ""
                row[f"{shop}_有货"] = "否"
                row[f"{shop}_匹配层"] = ""
        # 自家侧 + 结论
        row["自家售价"] = r.get("美团闪购渠道售价", "")
        row["采购价"] = r.get("采购价", "")
        try:
            own_price = float(row["自家售价"]) if pd.notna(row["自家售价"]) else None
        except (TypeError, ValueError):
            own_price = None
        if own_price and prices:
            median = sorted(prices)[len(prices) // 2]
            row["竞店中位数"] = round(median, 2)
            row["竞店最低价"] = round(min(prices), 2)
            row["竞店数"] = len(prices)
            row["价差率%"] = round((own_price - median) / median * 100, 1)
            row["结论"] = "最贵" if own_price > max(prices) else ("最便宜" if own_price < min(prices) else "居中")
        else:
            row["竞店中位数"] = ""
            row["竞店最低价"] = ""
            row["竞店数"] = len(prices)
            row["价差率%"] = ""
            row["结论"] = "仅竞店有价" if not own_price and prices else ""
        if pd.notna(r.get("采购价")) and own_price:
            try:
                cost = float(r["采购价"])
                row["毛利"] = round(own_price - cost, 2)
                row["毛利率%"] = round((own_price - cost) / own_price * 100, 1) if own_price else ""
            except (TypeError, ValueError):
                row["毛利"], row["毛利率%"] = "", ""
        else:
            row["毛利"], row["毛利率%"] = "", ""
        rows.append(row)
    matrix_df = pd.DataFrame(rows)
    return matrix_df, shop_data, own, own_ean_idx


def build_unmatched(shop_data, own, own_ean_idx):
    """竞店独有：竞店热销（月售>=阈值）且自家 EAN 无匹配（条码层）。"""
    own_eans = set(own_ean_idx.keys())
    rows = []
    for shop, sd in shop_data.items():
        for _, r in sd["df"].iterrows():
            # 热销只认真实成交。float("1.5k") 会抛异常被吞成 0，把真热销漏掉 → 统一走 sales_text
            raw = r.get("sales")
            sales = 0.0 if _ST.is_intent(raw) else _ST.num(raw)
            if sales < HOT_SALES:
                continue
            b = r["barcode"]
            if b and b != "nan" and b in own_eans:
                continue  # 已有条码级匹配
            rows.append({"店名": shop, "竞品单品编号": r["sku_id"], "竞品商品名": r["name"],
                         "月售": sales, "到手价": r["price"], "类目": r["category3"],
                         "条码": b if b else "无标准条码", "备注": "自家无标准条码匹配，待深层对齐或人工判断"})
    return pd.DataFrame(rows)


def build_barcode_detail(own, shop_data, own_ean_idx):
    """条码对齐明细：自家EAN x 店。"""
    rows = []
    for ean, oi_list in own_ean_idx.items():
        for shop, sd in shop_data.items():
            if ean in sd["ean_map"]:
                for oi in oi_list[:3]:
                    r = own.iloc[oi]
                    for comp_sku in sd["ean_map"][ean][:3]:
                        rows.append({"标准条码(13位)": ean, "店名": shop, "自家单品编号": r["SKUID"],
                                     "自家商品名": r.get("商品名称", ""), "竞品单品编号": comp_sku})
    return pd.DataFrame(rows)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="竞品价格矩阵构建（选店分析）")
    parser.add_argument("--shops", help="参与分析的竞店名（逗号分隔，默认全部店）")
    args = parser.parse_args()
    shops_filter = None
    if args.shops:
        shops_filter = set(s.strip() for s in args.shops.split(",") if s.strip())
        print(f"选店分析: {len(shops_filter)} 家店 -> {sorted(shops_filter)}")

    print("构建竞品价格矩阵...")
    matrix_df, shop_data, own, own_ean_idx = build_matrix(shops_filter)
    print(f"价格矩阵: {len(matrix_df)} 行（自家有匹配 SKU）")

    unmatched = build_unmatched(shop_data, own, own_ean_idx)
    print(f"竞店独有热销(月售>={HOT_SALES} 且自家无条码匹配): {len(unmatched)} 条")

    detail = build_barcode_detail(own, shop_data, own_ean_idx)
    print(f"条码对齐明细: {len(detail)} 条")

    # 纠错回写模板（不存在则创建空模板；已有则保留用户修正）
    if os.path.exists(FIX_FILE):
        fix_df = pd.read_excel(FIX_FILE)
    else:
        fix_df = pd.DataFrame(columns=["店名", "自家单品编号", "竞品单品编号", "修正说明", "修正时间"])
        fix_df.to_excel(FIX_FILE, index=False)
        print(f"已创建纠错回写模板: {FIX_FILE}（填写后重跑本脚本即生效）")

    with pd.ExcelWriter(OUT_FILE, engine="openpyxl") as writer:
        matrix_df.to_excel(writer, sheet_name="价格矩阵", index=False)
        unmatched.to_excel(writer, sheet_name="竞店独有", index=False)
        detail.to_excel(writer, sheet_name="条码对齐明细", index=False)
        fix_df.to_excel(writer, sheet_name="纠错回写", index=False)
    print(f"-> {OUT_FILE}")
    print(f"纠错回写文件: {FIX_FILE}（修正后重跑本脚本即生效）")


if __name__ == "__main__":
    main()

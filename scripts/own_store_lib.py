# -*- coding: utf-8 -*-
"""
自家门店库：多自家门店存储 + 在售集合（SPUID 与总库同源，直接 join 零模糊匹配）。

存储：output/自家门店/<店名>.xlsx（每店一份，同店重传=覆盖更新）
索引：output/自家门店/_stores.json

在售判定：售卖状态=售卖中 且 规格上下架状态=上架（其余不算在售）。
门店导出表无销量列——比对月销沿用总库口径（前置仓30天总销量），门店维度只影响
在售范围/上架判定（方案 §5 选项 A）。
"""
import json
import os
import re
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STORE_DIR = os.path.join(BASE, "output", "自家门店")
INDEX_FILE = os.path.join(STORE_DIR, "_stores.json")

# 门店表提取列（源表 -> 标准列）
EXTRACT_COLS = {
    "SPUID": "SPUID",
    "SKUID": "SKUID",
    "门店编码": "门店编码",
    "门店名称": "门店名称",
    "商品名称": "商品名称",
    "规格名称": "规格名称",
    "售卖状态": "售卖状态",
    "规格上下架状态": "规格上下架状态",
    "渠道库存": "渠道库存",
    "实物库存": "实物库存",
}


def is_store_export(columns):
    """识别是否自家门店导出表：表头含 门店编码 + SPUID（与竞品表/总库区分）。"""
    cols = set(str(c).strip() for c in columns)
    return "门店编码" in cols and "SPUID" in cols


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def load_index():
    if os.path.exists(INDEX_FILE):
        try:
            with open(INDEX_FILE, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_index(idx):
    os.makedirs(STORE_DIR, exist_ok=True)
    with open(INDEX_FILE, "w", encoding="utf-8") as f:
        json.dump(idx, f, ensure_ascii=False, indent=2)


def import_store(excel_path):
    """导入一份门店导出表 → 提取标准列 → 存为该店文件 + 更新索引。
    返回 dict：{ok, store, rows, on_sale_spus, on_sale_skus, msg}。"""
    df = pd.read_excel(excel_path, dtype=str).fillna("")
    if not is_store_export(df.columns):
        return {"ok": False, "msg": "不是门店导出表（缺 门店编码/SPUID 列）"}

    # 店名：门店名称列首个非空值；回退 门店编码
    names = [_clean(v) for v in df["门店名称"].tolist() if _clean(v)]
    store = names[0] if names else ""
    if not store:
        codes = [_clean(v) for v in df["门店编码"].tolist() if _clean(v)]
        store = codes[0] if codes else "未命名门店"
    # 文件名安全（去 Windows 非法字符）
    safe = re.sub(r'[\\/:*?"<>|]', "_", store)

    keep = [c for c in EXTRACT_COLS if c in df.columns]
    out = df[keep].copy()
    out["上报时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    os.makedirs(STORE_DIR, exist_ok=True)
    path = os.path.join(STORE_DIR, f"{safe}.xlsx")
    out.to_excel(path, index=False)

    # 在售集合统计
    on_sale_mask = (out["售卖状态"].astype(str).str.contains("售卖", na=False)) & \
                   (~out["规格上下架状态"].astype(str).str.contains("下架", na=False))
    os_df = out[on_sale_mask]
    n_spu = os_df["SPUID"].astype(str).str.strip().nunique() if len(os_df) else 0

    idx = load_index()
    idx[store] = {
        "file": f"{safe}.xlsx",
        "code": _clean(df["门店编码"].iloc[0]) if "门店编码" in df.columns else "",
        "rows": len(out),
        "on_sale_spus": int(n_spu),
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    save_index(idx)
    return {"ok": True, "store": store, "rows": len(out),
            "on_sale_spus": int(n_spu), "msg": f"门店[{store}]已导入（{len(out)} 行 / 在售 {n_spu} SPU）"}


# ---------------------------------------------------------------------------
# 在售集合缓存（mtime 失效）
# ---------------------------------------------------------------------------
_cache = {}  # {店名: {"spu_set": set, "sku_set": set, "mtime": float}}


def _on_sale_mask(df):
    return (df["售卖状态"].astype(str).str.contains("售卖", na=False)) & \
           (~df["规格上下架状态"].astype(str).str.contains("下架", na=False))


def get_on_sale(store):
    """返回该店在售集合 {"spu_set": set[str], "sku_set": set[str]}；店不存在返回 None。"""
    idx = load_index()
    info = idx.get(store)
    if not info:
        return None
    path = os.path.join(STORE_DIR, info.get("file", ""))
    if not os.path.exists(path):
        return None
    mt = os.path.getmtime(path)
    c = _cache.get(store)
    if c and c["mtime"] == mt:
        return c
    df = pd.read_excel(path, dtype=str).fillna("")
    os_df = df[_on_sale_mask(df)]
    data = {
        "spu_set": set(os_df["SPUID"].astype(str).str.strip()) if len(os_df) else set(),
        "sku_set": set(os_df["SKUID"].astype(str).str.strip()) if len(os_df) else set(),
        "mtime": mt,
    }
    _cache[store] = data
    return data


def list_stores():
    """[{name, code, rows, on_sale_spus, time}]（索引驱动）。"""
    idx = load_index()
    out = []
    for name, v in idx.items():
        out.append({"name": name, "code": v.get("code", ""),
                    "rows": v.get("rows", 0), "on_sale_spus": v.get("on_sale_spus", 0),
                    "time": v.get("time", "")})
    return sorted(out, key=lambda x: x["name"])


def invalidate():
    _cache.clear()

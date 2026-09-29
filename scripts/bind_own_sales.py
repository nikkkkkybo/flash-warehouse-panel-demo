# -*- coding: utf-8 -*-
"""
自家销量/库存回填：从「每日补货参考表」按 商品SKU（回退 商品UPC）回填到 自家_归一化后.xlsx：
  - 月销量      ← 30天总销量（原有行为，0 销视为未提供，保留原值）
  - 7天总销量   ← 7天总销量（新）
  - 14天总销量  ← 14天总销量（新）
  - 预测日均销量 ← 预测日均销量（新）
  - 可用库存    ← 可用库存（新）

口径说明：快照可能覆盖单个 门店/仓 或多个（如 示例片区B仓+示例片区A仓）——
多仓时按 仓 求和=合计口径；覆盖范围写入 output/own_sales_meta.json（UI 动态标注，防误读）。

数据源（优先项目内目录，回退旧外置目录）：
  1) <项目根目录>\每日补货参考\前置仓补货参考*.xlsx（每日快照，取最新；SNAP_DIR 环境变量优先）
  2) D:\补货参考\补货参考快照\（旧路径兼容）
用法：python scripts/bind_own_sales.py [--date 20260831] [--dry-run]
"""
import argparse
import glob
import json
import os
import re
import shutil
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")
BACKUP = OWN_FILE + ".bak_bindsales"
META_FILE = os.path.join(BASE, "output", "own_sales_meta.json")
# 快照目录：环境变量 > 项目内每日补货参考 > 旧外置快照目录
_CANDIDATE_DIRS = [
    os.environ.get("SNAP_DIR", ""),
    os.path.join(BASE, "每日补货参考"),
    r"D:\补货参考\补货参考快照",
]
SNAP_DIR = next((d for d in _CANDIDATE_DIRS if d and os.path.isdir(d)), _CANDIDATE_DIRS[1])

# 回填列：own表列名 ← 快照表列名
FILL_COLS = {
    "月销量": "30天总销量",
    "7天总销量": "7天总销量",
    "14天总销量": "14天总销量",
    "预测日均销量": "预测日均销量",
    "可用库存": "可用库存",
}
# 月销量沿用旧行为：0/空 视为未提供（保留原值）；其余列照抄原始值（含 0，真实数据）


def _latest_snapshot():
    files = glob.glob(os.path.join(SNAP_DIR, "前置仓补货参考*.xlsx"))
    if not files:
        raise RuntimeError("未找到补货参考快照: " + SNAP_DIR)
    return max(files, key=os.path.getmtime)


def norm_bc(v):
    s = str(v).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    d = re.sub(r"\D", "", s)
    return d if len(d) == 13 else (("0" + d) if len(d) == 12 else "")


def _cell(v):
    s = str(v).strip() if v is not None else ""
    return "" if s.lower() in ("nan", "none") else s


def _sales_num(v):
    """30天总销量 → int 字符串；0/空/非法 → ''（旧行为：0 视为未提供）。"""
    s = _cell(v)
    if s in ("", "0", "0.0"):
        return ""
    try:
        return str(int(float(s)))
    except (ValueError, TypeError):
        return ""


def read_snapshot(snap_path, cols=None):
    """读补货参考表：该表是二级表头（第一行是分区标题），真实列名在第二级。
    cols 给定时只保留这些列（省内存）。全项目读这份表只许走这里。"""
    df = pd.read_excel(snap_path, header=[0, 1], dtype=str)
    df.columns = [str(c[1]).strip() if str(c[1]).strip() else str(c[0]).strip()
                  for c in df.columns]
    df = df.loc[:, ~df.columns.duplicated()]
    if cols:
        keep = [c for c in cols if c in df.columns]
        df = df[keep]
    return df


def load_snap_maps(snap_path):
    """读补货参考表（二级表头），返回 (maps, maps_upc, snap_df)。
    多仓快照（如 示例片区B仓+示例片区A仓）按 仓 求和 → 合计口径（元数据如实标注覆盖仓）。
    键 = 商品SKU 优先 + 商品UPC(13位) 回退（两套 map 各自独立）。"""
    df = read_snapshot(snap_path)
    df["__key_sku"] = df["商品SKU"].map(_cell)
    df["__key_upc"] = df["商品UPC"].map(norm_bc)

    # 数值列按 键 求和（多仓合计）
    num_cols = list(FILL_COLS.values())
    agg = df.copy()
    for c in num_cols:
        agg[c] = pd.to_numeric(agg[c], errors="coerce").fillna(0)
    sum_sku = agg.groupby("__key_sku")[num_cols].sum()
    sum_upc = agg.groupby("__key_upc")[num_cols].sum()

    def _val(own_col, total):
        if own_col == "月销量":
            # 月销量沿用旧行为：合计 0 视为未提供
            return "" if total <= 0 else str(int(round(total)))
        return str(int(round(total))) if total else "0"

    maps = {col: {} for col in FILL_COLS}
    for key, row in sum_sku.iterrows():
        if not key or key.lower() == "nan":
            continue
        for own_col, snap_col in FILL_COLS.items():
            v = _val(own_col, float(row[snap_col]))
            if v != "":
                maps[own_col][key] = v
    maps_upc = {col: {} for col in FILL_COLS}
    for key, row in sum_upc.iterrows():
        if not key:
            continue
        for own_col, snap_col in FILL_COLS.items():
            v = _val(own_col, float(row[snap_col]))
            if v != "":
                maps_upc[own_col][key] = v
    return maps, maps_upc, df


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="快照日期 YYYYMMDD（默认取最新）")
    ap.add_argument("--file", help="直接指定快照文件（面板导入向导用；优先于 --date/取最新）")
    ap.add_argument("--dry-run", action="store_true", help="只统计不写回")
    args = ap.parse_args()

    if args.file:
        snap_path = args.file if os.path.isabs(args.file) else os.path.join(BASE, args.file)
        snap_path = os.path.normpath(snap_path)
        if not os.path.exists(snap_path):
            raise RuntimeError("快照文件不存在: " + snap_path)
    elif args.date:
        snaps = glob.glob(os.path.join(SNAP_DIR, f"前置仓补货参考*{args.date}*.xlsx"))
        if not snaps:
            raise RuntimeError(f"未找到 {args.date} 的快照")
        snap_path = max(snaps, key=os.path.getmtime)
    else:
        snap_path = _latest_snapshot()
    print("数据源:", os.path.basename(snap_path))

    maps, maps_upc, snap_df = load_snap_maps(snap_path)
    print("SKU 映射:", {k: len(v) for k, v in maps.items()})
    print("UPC 映射:", {k: len(v) for k, v in maps_upc.items()})

    own = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    print("自家表行数:", len(own))

    own["__bc"] = own["商品条码"].apply(norm_bc)
    stats = {col: {"sku": 0, "upc": 0, "blank": 0} for col in FILL_COLS}
    out = {}
    for col in FILL_COLS:
        sm, um = maps[col], maps_upc[col]
        vals = []
        for _, r in own.iterrows():
            sku = str(r.get("SKUID") or "").strip()
            if sku in sm:
                vals.append(sm[sku])
                stats[col]["sku"] += 1
            elif r["__bc"] in um:
                vals.append(um[r["__bc"]])
                stats[col]["upc"] += 1
            else:
                # 未匹配：一律清空（单快照单口径——不与历史快照混用，防口径混杂：
                # 旧快照可能是整体口径、新快照是分仓合计，混在一列数字就失去意义）
                vals.append("")
                stats[col]["blank"] += 1
        out[col] = vals
    own["__bc"] = None  # 占位，稍后统一 drop

    for col in FILL_COLS:
        if col in own.columns:
            own[col] = out[col]
        else:
            own.insert(len(own.columns), col, out[col])

    for col in FILL_COLS:
        filled = int((own[col].astype(str).str.strip().ne("")).sum())
        print(f"  {col}: 回填后有值 {filled} 行（SKUID {stats[col]['sku']} / UPC {stats[col]['upc']} / 未匹配清空 {stats[col]['blank']}）")

    if args.dry_run:
        print("[dry-run] 未写回")
        return

    # P4.3 口径元数据：快照覆盖的 门店/仓 维度（UI 动态标注"销量是谁的销量"）
    scope_stores, scope_codes = [], []
    try:
        codes = snap_df["门店/仓编码"].fillna("").astype(str).str.strip().unique().tolist()
        names = snap_df["门店/仓名称"].fillna("").astype(str).str.strip().unique().tolist()
        scope_codes = [c for c in codes if c and c.lower() != "nan"]
        scope_stores = [s for s in names if s and s.lower() != "nan"]
    except Exception:
        pass
    snap_date = ""
    m = re.search(r"(\d{8})", os.path.basename(snap_path))
    if m:
        snap_date = m.group(1)

    # 分仓留存：每仓每 SKU 一行（合计喂总库分析，分仓喂 店对店比较——
    # 仓的销量归仓，总库 SPUID 是关联；两套口径分开存，不混用）
    wh_file = os.path.join(BASE, "output", "自家销量_分仓.xlsx")
    try:
        wh_cols = ["门店/仓编码", "门店/仓名称", "商品SKU", "商品UPC",
                   "商品名称", "30天总销量", "14天总销量", "7天总销量", "预测日均销量", "可用库存"]
        wh = snap_df[wh_cols].copy()
        wh = wh[wh["商品SKU"].fillna("").astype(str).str.strip().ne("")]
        wh.to_excel(wh_file, index=False)
        print("分仓留存 ->", wh_file, f"（{len(wh)} 行 / {len(scope_stores)} 仓）")
    except Exception as e:
        print("[警告] 分仓留存失败:", e)

    meta = {
        "snapshot": os.path.basename(snap_path),
        "snapshot_date": snap_date,
        "stores": scope_stores,
        "store_codes": scope_codes,
        "rows": int(len(snap_df)),
        "bind_time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "note": ("单仓口径：" + scope_stores[0]) if len(scope_stores) == 1
                else ("多仓合计：" + "+".join(scope_stores)),
    }
    try:
        with open(META_FILE, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        print("口径元数据 ->", META_FILE, "| 仓:", scope_stores, "| 日期:", snap_date)
    except Exception as e:
        print("[警告] 元数据写入失败:", e)

    shutil.copy(OWN_FILE, BACKUP)
    print("已备份 ->", BACKUP)
    own = own.drop(columns=["__bc"], errors="ignore")
    own.to_excel(OWN_FILE, index=False)
    print("已写回", OWN_FILE, "(", datetime.now().strftime("%H:%M:%S"), ")")


if __name__ == "__main__":
    main()

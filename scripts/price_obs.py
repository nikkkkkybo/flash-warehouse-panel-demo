# -*- coding: utf-8 -*-
"""
价格观测历史（优化项C·Phase4）：竞店每次导入 = 一次全量价格观测，append-only 落 jsonl。
之前"重导入=覆盖旧价，价格历史正在丢失"；本模块让历史开始积累，price_diff.py 据此出变动报告。

文件：output/竞店标准库/价格观测历史.jsonl（一行一个 SKU 观测）
行结构：{"t":"2026-08-19 10:00","店":"...","id":"竞品单品编号","品名":"...",
         "原价":12.9,"折后价":9.9,"神价":8.8,"到手价":8.8,"月销量":300,"库存":50,"在售":true}

为什么 JSONL 不是 Excel：追加 O(1)（openpyxl 追加要整文件重写，10万行后越来越慢）；
报告渲染时再出 Excel（price_diff.py）。

用法：
  python price_obs.py --baseline        # 首启基线：现有多竞店汇总库全量落 t0 快照（每店一次）
  python price_obs.py --stats           # 观测统计（各店快照数/时间范围）
钩子：build_comp_library.add_one / build_batch 每次导入自动调用 snapshot_shop。
"""
import argparse
import json
import os
from collections import defaultdict
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIB_DIR = os.path.join(BASE, "output", "竞店标准库")
HIST_FILE = os.path.join(LIB_DIR, "价格观测历史.jsonl")
SUMMARY_FILE = os.path.join(LIB_DIR, "多竞店汇总库.xlsx")

# 观测字段：jsonl键 -> (中文列名, 英文列名)
_FIELDS = {
    "id": ("竞品单品编号", "sku_id"),
    "品名": ("商品名称", "name"),
    "原价": ("原价", "original_price"),
    "折后价": ("折后价", "discount_price"),
    "神价": ("神价", "god_price"),
    "到手价": ("到手价", "price"),
    "月销量": ("月销量", "sales"),
    "库存": ("库存", "stock"),
}


def _pick(df, names):
    """按 (中文, 英文) 候选列名取列；都不存在返回 None。"""
    for n in names:
        if n in df.columns:
            return df[n]
    return None


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def snapshot_shop(shop, df, when=None, source="导入"):
    """把一次导入的标准表全量行追加为该店观测快照。
    df 列名中文/英文均可。返回追加行数。"""
    if df is None or not len(df):
        return 0
    when = when or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cols = {k: _pick(df, names) for k, names in _FIELDS.items()}
    os.makedirs(LIB_DIR, exist_ok=True)
    n = 0
    with open(HIST_FILE, "a", encoding="utf-8") as f:
        for i in range(len(df)):
            rec = {"t": when, "店": shop, "在售": True, "来源": source}
            skip = True
            for k, col in cols.items():
                v = _clean(col.iloc[i]) if col is not None else ""
                rec[k] = v
                if k == "id" and v:
                    skip = False
            if skip:
                continue  # 无单品编号的行无法diff，跳过
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
            n += 1
    return n


def read_history(shop=None):
    """读 jsonl -> list[dict]（shop 为空=全部）。坏行跳过。"""
    out = []
    if not os.path.exists(HIST_FILE):
        return out
    with open(HIST_FILE, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except (ValueError, TypeError):
                continue
            if shop and str(rec.get("店") or "") != shop:
                continue
            out.append(rec)
    return out


def shop_snapshots(shop):
    """该店按时间分组的快照序列：[{t, rows: {id: rec}}, ...]（时间升序）。"""
    groups = defaultdict(dict)
    times = {}
    for rec in read_history(shop):
        t = str(rec.get("t") or "")
        sid = str(rec.get("id") or "")
        if not t or not sid:
            continue
        groups[t][sid] = rec
        times[t] = True
    return [{"t": t, "rows": groups[t]} for t in sorted(times.keys())]


def has_snapshot(shop):
    """该店是否已有任何观测（首启基线去重用）。"""
    for rec in read_history(shop):
        return True
    return False


def baseline_all(force=False):
    """首启基线：多竞店汇总库全量 -> 每店一份 t0 快照（已有观测的店跳过）。"""
    if not os.path.exists(SUMMARY_FILE):
        print(f"未找到 {SUMMARY_FILE}，先跑一次导入")
        return {}
    df = pd.read_excel(SUMMARY_FILE, dtype=str)
    shop_col = "竞店名称" if "竞店名称" in df.columns else "shop_name"
    result = {}
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for shop, g in df.groupby(shop_col):
        shop = str(shop).strip()
        if not shop:
            continue
        if not force and has_snapshot(shop):
            print(f"  [{shop}] 已有观测历史，跳过基线")
            continue
        n = snapshot_shop(shop, g, when=now, source="首启基线")
        result[shop] = n
        print(f"  [{shop}] 基线落库 {n} 行")
    return result


def stats():
    """观测统计：各店快照数/时间范围/jsonl大小。"""
    by_shop = defaultdict(list)
    for rec in read_history():
        by_shop[str(rec.get("店") or "?")].append(str(rec.get("t") or "?"))
    size_mb = os.path.getsize(HIST_FILE) / 1024 / 1024 if os.path.exists(HIST_FILE) else 0
    print(f"观测文件: {HIST_FILE}（{size_mb:.1f} MB）")
    if not by_shop:
        print("  （空，先跑 --baseline 建基线）")
        return
    for shop, ts in sorted(by_shop.items()):
        snaps = sorted(set(ts))
        print(f"  [{shop}] 观测 {len(ts)} 行 / {len(snaps)} 次快照"
              f"（{snaps[0]} ~ {snaps[-1]}）")


def price_of(rec):
    """观测记录的有效价：到手价 > 神价 > 折后价 > 原价（与入库兜底顺序一致）。"""
    for k in ("到手价", "神价", "折后价", "原价"):
        v = rec.get(k)
        try:
            f = float(str(v).strip())
            if f > 0:
                return f
        except (ValueError, TypeError):
            continue
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--baseline", action="store_true", help="首启基线（多竞店汇总库全量落 t0 快照）")
    ap.add_argument("--force", action="store_true", help="基线强制重落（已有历史也重落）")
    ap.add_argument("--stats", action="store_true", help="观测统计")
    args = ap.parse_args()
    if args.baseline:
        baseline_all(force=args.force)
    elif args.stats:
        stats()
    else:
        stats()


if __name__ == "__main__":
    main()

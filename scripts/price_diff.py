# -*- coding: utf-8 -*-
"""
竞店变动报告（优化项C·Phase4）：同店每次导入后，diff 最近两次观测快照，
出「降价/涨价/新品/下架/断货回补/undercut警报」六类变动 Excel 报告。

数据源：output/竞店标准库/价格观测历史.jsonl（price_obs.py 积累）+ 对齐汇总库（undercut）。
输出：output/竞店变动报告_<店>_<日期>.xlsx
规则（方案 4.3）：
  降价/涨价：到手价变动幅度 ≥5%
  新品：id 首次出现
  下架：连续 2 次导入消失才报（防抖，单次消失不报）
  断货/回补：库存 0 / 从 0 恢复
  undercut：竞品到手价 < 自家售价 且已有对齐关系（汇总库有效行，零误报噪音）

用法：python price_diff.py [--shop 店名]   # 不带 --shop = 全店各出一份
钩子：build_comp_library.add_one 导入后自动调用（同店更新时）。
"""
import argparse
import os
from datetime import datetime

import pandas as pd

from price_obs import shop_snapshots, price_of

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
LIB_FILE = os.path.join(BASE, "output", "对齐汇总库.xlsx")
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")

MIN_PRICE_CHANGE = 0.05     # 降价/涨价报告阈值（幅度≥5%）
UNDERCUT_EPS = 0.01         # undercut 判定：竞品价 < 自家价 - 0.01

REPORT_COLS = ["竞品Sku_id", "品名", "变动", "前值", "现值", "幅度", "月销量", "备注"]


def _num(v):
    try:
        f = float(str(v).strip())
        return f if f == f else None   # NaN 防御
    except (ValueError, TypeError):
        return None


def _own_price_map():
    """自家 SKUID -> 售价（undercut 用）。"""
    if not os.path.exists(OWN_FILE):
        return {}
    try:
        df = pd.read_excel(OWN_FILE, dtype=str)
    except Exception:
        return {}
    if "SKUID" not in df.columns:
        return {}
    price_col = next((c for c in ["美团闪购渠道售价", "售价", "价格"] if c in df.columns), None)
    out = {}
    for _, r in df.iterrows():
        sku = str(r.get("SKUID") or "").strip()
        if sku and price_col:
            p = _num(r.get(price_col))
            if p:
                out[sku] = p
    return out


def _matched_pairs(shop):
    """汇总库 SKU 层该店有效对齐：{竞品Sku_id: (自家SKUID, 自家商品名)}。"""
    if not os.path.exists(LIB_FILE):
        return {}
    try:
        df = pd.read_excel(LIB_FILE, sheet_name="SKU匹配层", dtype=str).fillna("")
    except (ValueError, FileNotFoundError):
        return {}
    if not len(df):
        return {}
    m = df[(df["竞店名称"].astype(str) == shop) & (df["状态"].astype(str) == "有效")]
    return {str(r.get("竞品Sku_id") or "").strip():
            (str(r.get("自家SKUID") or "").strip(), str(r.get("自家商品名") or "").strip())
            for _, r in m.iterrows()
            if str(r.get("竞品Sku_id") or "").strip()}


def diff_shop(shop, with_undercut=True):
    """diff 该店最近两次观测。返回 (rows, meta)；快照不足 2 次时 rows 为空、meta 说明基线。"""
    snaps = shop_snapshots(shop)
    meta = {"shop": shop, "snapshots": len(snaps)}
    if len(snaps) < 2:
        meta["note"] = "观测不足2次（基线已建立，下次导入出变动报告）" if len(snaps) == 1 \
            else "无观测历史（先跑 price_obs --baseline 或导入一次）"
        return [], meta
    prev, cur = snaps[-2], snaps[-1]
    meta["prev_t"], meta["cur_t"] = prev["t"], cur["t"]

    rows = []
    prev_ids, cur_ids = set(prev["rows"]), set(cur["rows"])

    # ---- 价格变动（降价/涨价，幅度≥5%）----
    for sid in sorted(prev_ids & cur_ids):
        p_old, p_new = price_of(prev["rows"][sid]), price_of(cur["rows"][sid])
        if p_old is None or p_new is None or p_old <= 0:
            continue
        if abs(p_new - p_old) / p_old < MIN_PRICE_CHANGE:
            continue
        kind = "降价" if p_new < p_old else "涨价"
        rows.append({
            "竞品Sku_id": sid, "品名": cur["rows"][sid].get("品名", ""),
            "变动": kind, "前值": p_old, "现值": p_new,
            "幅度": f"{(p_new - p_old) / p_old * 100:+.1f}%",
            "月销量": cur["rows"][sid].get("月销量", ""),
            "备注": f"{prev['t']} → {cur['t']}",
        })

    # ---- 新品（首次出现）----
    ever = set()
    for s in snaps[:-1]:
        ever |= set(s["rows"])
    for sid in sorted(cur_ids - prev_ids):
        if sid in ever:      # 曾出现→中间下架又回来：恢复上架，不算新品
            rows.append({
                "竞品Sku_id": sid, "品名": cur["rows"][sid].get("品名", ""),
                "变动": "恢复上架", "前值": "", "现值": price_of(cur["rows"][sid]) or "",
                "幅度": "", "月销量": cur["rows"][sid].get("月销量", ""),
                "备注": "此前消失过，本次回归",
            })
        else:
            rows.append({
                "竞品Sku_id": sid, "品名": cur["rows"][sid].get("品名", ""),
                "变动": "新品", "前值": "", "现值": price_of(cur["rows"][sid]) or "",
                "幅度": "", "月销量": cur["rows"][sid].get("月销量", ""),
                "备注": "首次出现",
            })

    # ---- 下架（连续 2 次导入消失才报，防抖）----
    if len(snaps) >= 3:
        for sid in sorted(set(snaps[-3]["rows"]) - prev_ids - cur_ids):
            rows.append({
                "竞品Sku_id": sid, "品名": snaps[-3]["rows"][sid].get("品名", ""),
                "变动": "下架", "前值": price_of(snaps[-3]["rows"][sid]) or "",
                "现值": "", "幅度": "",
                "月销量": snaps[-3]["rows"][sid].get("月销量", ""),
                "备注": f"连续2次未出现（{snaps[-2]['t']}、{cur['t']}）",
            })

    # ---- 断货/回补（库存 0 <-> 恢复）----
    for sid in sorted(prev_ids & cur_ids):
        s_old, s_new = prev["rows"][sid].get("库存", ""), cur["rows"][sid].get("库存", "")
        k_old, k_new = _num(s_old), _num(s_new)
        if k_old is None or k_new is None:
            continue
        if k_old > 0 and k_new == 0:
            rows.append({"竞品Sku_id": sid, "品名": cur["rows"][sid].get("品名", ""),
                         "变动": "断货", "前值": k_old, "现值": 0, "幅度": "",
                         "月销量": cur["rows"][sid].get("月销量", ""), "备注": "库存归零"})
        elif k_old == 0 and k_new > 0:
            rows.append({"竞品Sku_id": sid, "品名": cur["rows"][sid].get("品名", ""),
                         "变动": "回补", "前值": 0, "现值": k_new, "幅度": "",
                         "月销量": cur["rows"][sid].get("月销量", ""), "备注": "库存恢复"})

    # ---- undercut 警报（有对齐关系 且 竞品现价 < 自家价）----
    if with_undercut:
        pairs = _matched_pairs(shop)
        own_price = _own_price_map()
        for sid, (own_sku, own_name) in pairs.items():
            rec = cur["rows"].get(sid)
            if not rec:
                continue
            p_new = price_of(rec)
            p_own = own_price.get(own_sku)
            if p_new is None or p_own is None:
                continue
            if p_new < p_own - UNDERCUT_EPS:
                rows.append({
                    "竞品Sku_id": sid, "品名": rec.get("品名", ""),
                    "变动": "undercut警报", "前值": p_own, "现值": p_new,
                    "幅度": f"-¥{round(p_own - p_new, 2)}",
                    "月销量": rec.get("月销量", ""),
                    "备注": f"比自家「{own_name}」(¥{p_own}) 便宜，建议核对防守价",
                })

    # 排序：警报/降价/涨价优先，其次月销量降序
    order = {"undercut警报": 0, "降价": 1, "涨价": 2, "断货": 3,
             "新品": 4, "恢复上架": 5, "下架": 6, "回补": 7}
    rows.sort(key=lambda r: (order.get(r["变动"], 9),
                             -(_num(r.get("月销量")) or 0)))
    return rows, meta


def build_report(shop):
    """生成该店变动报告 Excel。返回 (报告路径, meta) 或 (None, meta)。"""
    rows, meta = diff_shop(shop)
    date = datetime.now().strftime("%Y%m%d")
    path = os.path.join(OUTPUT, f"竞店变动报告_{shop}_{date}.xlsx")
    if not rows:
        # 无变动也落一份摘要（覆盖旧报告，避免误读上次结果）
        pd.DataFrame([{"说明": meta.get("note") or
                       f"无变动（{meta.get('prev_t')} → {meta.get('cur_t')}）"}]) \
            .to_excel(path, index=False)
        meta["report"] = path
        meta["rows"] = 0
        return path, meta
    df = pd.DataFrame(rows)[REPORT_COLS]
    # 单文件多 sheet 按变动类型分拆 + 全量首sheet
    with pd.ExcelWriter(path, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="全部变动", index=False)
        for kind in ["undercut警报", "降价", "涨价", "断货", "新品", "恢复上架", "下架", "回补"]:
            sub = df[df["变动"] == kind]
            if len(sub):
                sub.to_excel(w, sheet_name=kind[:31], index=False)
    meta["report"] = path
    meta["rows"] = len(rows)
    return path, meta


def build_all(shops=None):
    """全店（或指定店）各出一份报告。返回 [meta]。"""
    if shops is None:
        # 有观测的店全集
        from price_obs import read_history
        shops = sorted({str(r.get("店")) for r in read_history() if r.get("店")})
    metas = []
    for shop in shops:
        if not shop:
            continue
        _, meta = build_report(shop)
        metas.append(meta)
    return metas


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--shop", help="指定店名（缺省=全部有观测的店）")
    args = ap.parse_args()
    metas = build_all([args.shop] if args.shop else None)
    if not metas:
        print("无观测历史：先跑 python price_obs.py --baseline 建基线，或导入一次竞店表")
        return
    for m in metas:
        if m.get("note"):
            print(f"[{m['shop']}] {m['note']}")
            continue
        dist = {}
        # 重新取分布
        rows, _ = diff_shop(m["shop"])
        for r in rows:
            dist[r["变动"]] = dist.get(r["变动"], 0) + 1
        dist_txt = " | ".join(f"{k} {v}" for k, v in sorted(dist.items())) or "无变动"
        print(f"[{m['shop']}] {m['prev_t']} → {m['cur_t']}：{m.get('rows', 0)} 条变动（{dist_txt}）")
        print(f"  报告 -> {m.get('report')}")


if __name__ == "__main__":
    main()

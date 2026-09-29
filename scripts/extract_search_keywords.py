# -*- coding: utf-8 -*-
"""
P3 词分级提炼工具：搜索点击日志 → 词级分类 → 反哺词库

读 output/搜索学习日志.jsonl（词→品组 人工标注），按四级分级：
  点名词（品级强词）：频次≥2 且 指向品组≤2 且 词在品名中 → 直接写入关键词库（来源=搜索反哺）
  规格词：词匹配品组内"部分"SKU（区分度判定）→ 规格词库（品内筛选用）
  疑似同义词：词不在品名中 → 同义词待确认清单
  泛词：指向品组≥5 → 停用词候选清单

分级逻辑全部本地统计，不调用大模型（0 token）。
用法：python scripts/extract_search_keywords.py [--min-freq 2] [--apply] [--shop 示例竞店A]
      --apply 才把点名词写入关键词库（默认 dry-run 预览）
"""
import json
import os
import sys
from collections import defaultdict
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUTPUT = os.path.join(BASE, "output")
LOG_FILE = os.path.join(OUTPUT, "搜索学习日志.jsonl")
KW_FILE = os.path.join(OUTPUT, "自家关键词库.xlsx")
OWN_FILE = os.path.join(OUTPUT, "自家_归一化后.xlsx")
GENERIC_LIMIT = 5   # 指向品组数 ≥ 5 视为泛词


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _read_log():
    if not os.path.exists(LOG_FILE):
        return []
    rows = []
    try:
        with open(LOG_FILE, encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line:
                    try:
                        rows.append(json.loads(line))
                    except Exception:
                        continue
    except Exception:
        pass
    return rows


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--min-freq", type=int, default=2, help="点名词最低频次")
    ap.add_argument("--apply", action="store_true", help="真正写入关键词库（默认预览）")
    args = ap.parse_args()

    rows = _read_log()
    if not rows:
        print("搜索学习日志为空（先真实搜索点选积累样本）")
        return
    print("日志样本数:", len(rows))

    # 词→品组 频次；词→品组集合
    word_spu = defaultdict(lambda: defaultdict(int))
    for r in rows:
        for w in r.get("words", []):
            w = _clean(w)
            spu = _clean(r.get("spu"))
            if w and spu:
                word_spu[w][spu] += 1

    # 自家表（品组名/类目/规格，供区分度与"是否在品名"判定）
    own = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    spu_name = {}
    spu_cat = {}
    spu_specs = defaultdict(list)
    for spuid, g in own.groupby("SPUID"):
        spuid = _clean(spuid)
        if not spuid:
            continue
        rep = g.iloc[0]
        spu_name[spuid] = _clean(rep.get("商品名称")) or _clean(rep.get("清洗后商品名"))
        spu_cat[spuid] = _clean(rep.get("类目三级_归一")) or _clean(rep.get("美团闪购类目三级"))
        spu_specs[spuid] = [_clean(x) for x in g["规格名称"] if _clean(x)]

    # 分级
    points, specs, syns, generics = [], [], [], []
    for w, spu_cnt in word_spu.items():
        spus = list(spu_cnt.keys())
        n_groups = len(spus)
        total_freq = sum(spu_cnt.values())
        if n_groups >= GENERIC_LIMIT:
            generics.append({"词": w, "指向品组数": n_groups, "总频次": total_freq,
                             "品组样例": "、".join(spu_name.get(s, "?")[:12] for s in spus[:3])})
            continue
        for spu, freq in spu_cnt.items():
            name = spu_name.get(spu, "")
            cat = spu_cat.get(spu, "")
            in_name = w in name
            # 规格词判定：匹配品组内部分 SKU（区分度）
            specs_g = spu_specs.get(spu, [])
            hit = sum(1 for s in specs_g if w in s)
            is_spec = 0 < hit < len(specs_g) if len(specs_g) > 1 else False
            item = {"词": w, "品组ID": spu, "品名": name[:40], "类目": cat,
                    "频次": freq, "指向品组数": n_groups, "在品名中": "是" if in_name else "否",
                    "规格词": "是" if is_spec else ""}
            # 分级：规格词优先（品内区分）→ 不在品名且非规格词=疑似同义词 →
            # 高频+点名+在品名=点名词 → 其余低频观察（不入库，不误判同义词）
            if is_spec:
                specs.append(item)
            elif not in_name:
                syns.append(item)
            elif freq >= args.min_freq and n_groups <= 2:
                points.append(item)

    print("\n=== 分级结果 ===")
    print("点名词（可入库）: %d" % len(points))
    for p in points[:10]:
        print("  %s → %s（%d次，%d品组）" % (p["词"], p["品名"][:20], p["频次"], p["指向品组数"]))
    print("规格词: %d" % len(specs))
    for p in specs[:10]:
        print("  %s → %s（品内区分）" % (p["词"], p["品名"][:20]))
    print("疑似同义词: %d" % len(syns))
    for p in syns[:10]:
        print("  %s → %s（不在品名）" % (p["词"], p["品名"][:20]))
    print("泛词（指向≥%d品组）: %d" % (GENERIC_LIMIT, len(generics)))
    for p in generics[:5]:
        print("  %s（%d品组）" % (p["词"], p["指向品组数"]))

    # 待确认清单 Excel
    out = os.path.join(OUTPUT, "搜索反哺_词分级_待确认.xlsx")
    with pd.ExcelWriter(out) as w:
        if points:
            pd.DataFrame(points).to_excel(w, sheet_name="点名词(可入库)", index=False)
        if specs:
            pd.DataFrame(specs).to_excel(w, sheet_name="规格词", index=False)
        if syns:
            pd.DataFrame(syns).to_excel(w, sheet_name="疑似同义词", index=False)
        if generics:
            pd.DataFrame(generics).to_excel(w, sheet_name="泛词(停用候选)", index=False)
    print("\n待确认清单:", out)

    # 写入关键词库（点名词，--apply）
    if points and args.apply:
        kw = pd.read_excel(KW_FILE, dtype=str).fillna("") if os.path.exists(KW_FILE) \
            else pd.DataFrame(columns=["关键词", "品组ID", "品名", "类目", "来源", "时间", "状态"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        added = 0
        for p in points:
            if len(kw) and ((kw["关键词"] == p["词"]) & (kw["品组ID"] == p["品组ID"])).any():
                continue
            kw = pd.concat([kw, pd.DataFrame([{
                "关键词": p["词"], "品组ID": p["品组ID"], "品名": p["品名"],
                "类目": p["类目"], "来源": "搜索反哺", "时间": now, "状态": "启用"}])],
                ignore_index=True)
            added += 1
        if added:
            kw.to_excel(KW_FILE, index=False)
        print("已写入关键词库 %d 条（来源=搜索反哺）" % added)


if __name__ == "__main__":
    main()

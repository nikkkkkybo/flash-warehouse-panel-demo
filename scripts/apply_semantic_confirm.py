# -*- coding: utf-8 -*-
"""
0.80 强对应批量确认工具（可选执行，不默认开启）

把「品级语义层」置信度=0.80（类目✓+宽品名✓+细品✓）的候选一键写入
两层汇总库 SPU 匹配层（品已对应，规格待补）+ progress（spu_only）。

用法：python scripts/apply_semantic_confirm.py [--shop 示例竞店A]
      --dry-run 只预览不写入（默认 dry-run；加 --apply 才真正写入）
"""
import os
import sys
import json
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUTPUT = os.path.join(BASE, "output")
PROGRESS_FILE = os.path.join(BASE, "scripts", "progress.json")

from spu_semantic import _clean  # noqa: E402
from summary_library import load_library, save_library, upsert_manual  # noqa: E402

CONF = "0.80"


def main():
    import argparse
    ap = argparse.ArgumentParser()
    ap.add_argument("--shop", default="示例竞店A")
    ap.add_argument("--apply", action="store_true", help="真正写入（默认 dry-run）")
    ap.add_argument("--export", action="store_true", help="导出强对应预览 Excel 供查看")
    args = ap.parse_args()

    sem_f = os.path.join(OUTPUT, f"试跑_{args.shop}", "对齐_第2.5层_品级语义_人工池.xlsx")
    if not os.path.exists(sem_f):
        print("未找到:", sem_f)
        return
    df = pd.read_excel(sem_f, dtype=str).fillna("")
    # 置信度数值归一比较（Excel 写 0.8 而非 0.80）
    df["__conf"] = pd.to_numeric(df["置信度"], errors="coerce")
    strong = df[(df["__conf"] - float(CONF)).abs() < 0.001].drop(columns=["__conf"])
    print("品级语义候选:", len(df), "| 0.80 强对应:", len(strong))
    if not len(strong):
        return

    # ---- 导出预览（含竞品月售/价格 与 自家月销/售价/品牌，便于人工审核）----
    if args.export:
        comp = pd.read_excel(os.path.join(OUTPUT, f"试跑_{args.shop}", "竞品_归一化后.xlsx"),
                             dtype=str).fillna("")
        comp_info = {}
        if len(comp):
            sku_col = "sku_id" if "sku_id" in comp.columns else "Sku_id"
            for _, r in comp.iterrows():
                sid = _clean(r.get(sku_col))
                if sid:
                    comp_info[sid] = r
        own = pd.read_excel(os.path.join(OUTPUT, "自家_归一化后.xlsx"), dtype=str).fillna("")
        own_info = {}
        if len(own):
            for _, r in own.iterrows():
                sid = _clean(r.get("SKUID"))
                if sid:
                    own_info[sid] = r

        exp = strong.copy()
        exp["竞品月售"] = exp["竞品Sku_id"].map(
            lambda s: _clean(comp_info.get(s, {}).get("月售")
                             or comp_info.get(s, {}).get("sales")
                             or comp_info.get(s, {}).get("月销量")) if s in comp_info else "")
        exp["竞品到手价"] = exp["竞品Sku_id"].map(
            lambda s: _clean(comp_info.get(s, {}).get("到手价")
                             or comp_info.get(s, {}).get("price")
                             or comp_info.get(s, {}).get("折后价")) if s in comp_info else "")
        exp["自家月销"] = exp["自家SKUID"].map(
            lambda s: _clean(own_info.get(s, {}).get("月销量")) if s in own_info else "")
        exp["自家售价"] = exp["自家SKUID"].map(
            lambda s: _clean(own_info.get(s, {}).get("美团闪购渠道售价")) if s in own_info else "")
        exp["自家品牌"] = exp["自家SKUID"].map(
            lambda s: (_clean(own_info.get(s, {}).get("品牌（抽取）"))
                       or _clean(own_info.get(s, {}).get("品牌名称"))) if s in own_info else "")
        exp["自家采购价"] = exp["自家SKUID"].map(
            lambda s: _clean(own_info.get(s, {}).get("采购价")) if s in own_info else "")
        # 跨品牌标记
        exp["跨品牌"] = exp.apply(
            lambda r: ("是" if (_clean(r.get("竞品品牌")) and _clean(r.get("自家品牌"))
                               and _clean(r.get("竞品品牌")) != _clean(r.get("自家品牌"))) else ""),
            axis=1)
        exp["竞品品牌"] = exp["竞品品牌"].apply(lambda v: _clean(v) or "（无品牌）")
        exp["自家品牌"] = exp["自家品牌"].apply(lambda v: _clean(v) or "（无品牌）")
        # 按月售降序（跨品牌在前）
        exp["__cs"] = pd.to_numeric(exp["竞品月售"], errors="coerce").fillna(-1)
        exp = exp.sort_values(["跨品牌", "__cs"], ascending=[False, False]).drop(columns=["__cs"])
        cols = ["竞品商品名", "竞品品牌", "竞品月售", "竞品到手价", "竞品类目",
                "自家商品名", "自家品牌", "自家月销", "自家售价", "自家采购价",
                "语义层级", "跨品牌", "竞品Sku_id", "自家SKUID"]
        exp = exp[[c for c in cols if c in exp.columns]]
        exp.insert(0, "序号", range(1, len(exp) + 1))
        out = os.path.join(OUTPUT, f"强对应_0.80_预览_{args.shop}_{datetime.now().strftime('%H%M%S')}.xlsx")
        exp.to_excel(out, index=False)
        print("已导出:", out)
        print("  其中跨品牌:", int((exp["跨品牌"] == "是").sum()), "条 | 有月售:", int((exp["竞品月售"].astype(str).str.strip().replace('nan','') != '').sum()), "条")
        return

    spu_df, sku_df = load_library()
    progress = {}
    if os.path.exists(PROGRESS_FILE):
        progress = json.load(open(PROGRESS_FILE, encoding="utf-8"))

    rows_written = 0
    for _, r in strong.iterrows():
        pid = _clean(r.get("竞品品ID"))
        sku = _clean(r.get("竞品Sku_id"))
        if not pid or not sku:
            continue
        # SPU 层行（品对应；规格待补）
        spu_row = {
            "竞品品ID": pid,
            "竞品商品名": _clean(r.get("竞品商品名")),
            "竞品品牌": _clean(r.get("竞品品牌")),
            "竞品型号": _clean(r.get("竞品型号")),
            "竞品类目": _clean(r.get("竞品类目")),
            "自家品组ID": _clean(r.get("自家SPU组ID")) or _clean(r.get("自家SPUID")),
            "自家品名": _clean(r.get("自家商品名")),
            "置信度": CONF,
        }
        if args.apply:
            spu_df = upsert_manual(spu_df, "竞品品ID", args.shop, spu_row)
            progress.setdefault("confirmed", {}).setdefault(args.shop, {})[sku] = {
                "sku_id": _clean(r.get("自家SKUID")),
                "spu_only": True,
                "source": "品级语义0.80",
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
        rows_written += 1

    if args.apply:
        save_library(spu_df, sku_df)
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(progress, f, ensure_ascii=False, indent=2)
        print(f"已写入 {rows_written} 条到 SPU 匹配层（{args.shop}，置信 0.80）")
    else:
        print(f"[dry-run] 将写入 {rows_written} 条（SPU 层品对应）。确认后加 --apply 执行。")
        # 预览前 10 条
        print("\n预览（前10条）：")
        for _, r in strong.head(10).iterrows():
            print("- %s → %s" % (str(r.get("竞品商品名"))[:34], str(r.get("自家商品名"))[:30]))


if __name__ == "__main__":
    main()

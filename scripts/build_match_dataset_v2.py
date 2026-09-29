# -*- coding: utf-8 -*-
"""微调 v2 训练数据生成器：品级方向（纠正 v1 类目词方向的塌缩教训）。

与 v1（build_match_dataset.py 产物喂类目词）的区别：
  - 正样本 = 竞品全名 ↔ 自家同品全名（品级对，不出现类目短词）
  - 增广 = 同品组内其他 SKU 的自家全名与同一竞品名配对（同品不同挂法，天然正对）
  - 留出 = 按 品组(SPUID) 切 80/20，同组绝不跨训练/验证，防泄漏
  - 难负 = 难负样本 sheet（同店同品类目 TF-IDF 近邻，品级难负）只做评估，不进损失
  - 文本在生成时就用 align._vec_clean 清洗好，训练/评估零依赖 align

产物（output/model_train/v2/）：
  mnr_v2.jsonl      {a, p}       训练对（MNR）
  val_pairs_v2.jsonl {q, gid}    黄金验证对（留出组）
  groups_v2.jsonl   [{gid, t}]   全部自家品组清洗名（召回评测库）
  hard_neg_v2.jsonl {a, b}       品级难负对（仅评估用）
跑法：python scripts/build_match_dataset_v2.py
"""
import io
import json
import os
import random
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
import pandas as pd
import align as AL
import sales_text as STx

OUT = os.path.join(BASE, "output")
V2 = os.path.join(OUT, "model_train", "v2")
EXTRA_SKU_CAP = 3
VAL_FRAC = 0.2
SEED = 42


def cl(name, brand=""):
    return AL._vec_clean(name, AL._clean(brand))


def main():
    os.makedirs(V2, exist_ok=True)
    random.seed(SEED)

    ds = pd.read_excel(os.path.join(OUT, "训练样本库.xlsx"), sheet_name="正样本", dtype=str)
    ds = ds[ds["自家名"].astype(str).str.strip().ne("") & ds["自家名"].notna()]
    own = pd.read_excel(os.path.join(OUT, "自家_归一化后.xlsx"), dtype=str)
    own = own.drop_duplicates(subset=["SKUID"])
    own["__s"] = own["月销量"].apply(lambda v: float(STx.eff(v) or 0))
    own["__sp"] = own["SPUID"].astype(str).str.strip()

    sku2gid = dict(zip(own["SKUID"].astype(str), own["__sp"]))
    groups, sku_names = {}, {}
    for gid, g in own.groupby("__sp"):
        g = g.sort_values("__s", ascending=False)
        b = g.iloc[0]
        nm = str(b.get("商品名称") or "").strip()
        if nm:
            groups[gid] = cl(nm, b.get("品牌（抽取）"))
        sku_names[gid] = [cl(str(x or ""), b.get("品牌（抽取）")) for x in g["商品名称"].head(1 + EXTRA_SKU_CAP)]

    pairs, seen = [], set()
    for _, r in ds.iterrows():
        oid = str(r.get("自家ID") or "").strip()
        gid = sku2gid.get(oid, "")
        cn, on = str(r["竞品名"]), str(r["自家名"])
        key = (cn, on)
        if not gid or gid not in groups or key in seen:
            continue
        seen.add(key)
        pairs.append({"a": cl(cn), "p": cl(on), "gid": gid})
    print(f"黄金对去重后 {len(pairs)} 条（原始 {len(ds)} 行）")

    gids = sorted({p["gid"] for p in pairs})
    random.shuffle(gids)
    n_val = max(5, int(len(gids) * VAL_FRAC))
    val_gids = set(gids[:n_val])
    train, val = [], []
    for p in pairs:
        (val if p["gid"] in val_gids else train).append(p)
    print(f"品组级切分：训练 {len(train)} 对 / 验证 {len(val)} 对（验证组 {n_val} 个）")

    aug = 0
    for p in list(train):
        for alt in sku_names.get(p["gid"], []):
            if alt and alt != p["p"]:
                train.append({"a": p["a"], "p": alt, "gid": p["gid"]})
                aug += 1
    print(f"同组 SKU 挂法增广 {aug} 条 → 训练对 {len(train)} 条")

    neg = pd.read_excel(os.path.join(OUT, "训练样本库.xlsx"), sheet_name="难负样本", dtype=str)
    neg = neg[neg["竞品名"].notna() & neg["自家名"].notna()]
    hard = [{"a": cl(str(r["竞品名"])), "b": cl(str(r["自家名"]))} for _, r in neg.iterrows()]
    hard = [h for h in hard if h["a"] and h["b"]]
    print(f"品级难负对 {len(hard)} 条")

    def dump(name, obj):
        with open(os.path.join(V2, name), "w", encoding="utf-8") as f:
            for x in obj:
                f.write(json.dumps(x, ensure_ascii=False) + "\n")
        print(f"  写 {name}")

    dump("mnr_v2.jsonl", [{"a": p["a"], "p": p["p"]} for p in train])
    dump("val_pairs_v2.jsonl", [{"q": p["a"], "gid": p["gid"]} for p in val])
    dump("groups_v2.jsonl", [{"gid": g, "t": t} for g, t in groups.items()])
    dump("hard_neg_v2.jsonl", hard)
    print("v2 数据就绪")


if __name__ == "__main__":
    main()

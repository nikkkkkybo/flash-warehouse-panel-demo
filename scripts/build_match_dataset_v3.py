# -*- coding: utf-8 -*-
"""训练数据 v3：竞品名 → 自家品组 对齐专用（替换 题词向量库/ 旧数据）。

设计原则（来自 2026-09-09 质检结论）：
  1. 标签只由"产品级等价证据"生成，绝不用"规格 token 相同"当同品判据；
  2. 正对三层来源：L1a 条码相同 / L1b 品牌+型号全等 / L3 人工确认；
  3. 负对（L5）：已知锚点真值后，取**同类目、规格冲突**的自家品组，
     并排除条码/型号/近重复，避免假负例；
  4. 文本清洗**保留规格**（规格是五金件"同品/非同品"的判别核心），只去品牌与营销词；
  5. 评估集优先用人工确认对，训练/评估 anchor 零重叠（硬断言，不满足直接报错）；
  6. 产物附 meta.json：源表指纹 + 口径 + 计数 + 重叠检查。

产物（output/model_train/v3/）：
  groups.jsonl       {gid, t, cat, brand}        自家品组检索库（线上同口径）
  pos_pairs.jsonl    {a, p, gid, src, conf}      正对（训练）
  triplet_v3.jsonl   {anchor, positive, negative, reason, sim}   MNR+显式难负
  hard_neg_v3.jsonl  {a, b, reason, human}       难负对（评估/拒识标定）
  eval_pairs.jsonl   {q, gid, src}               评估 query（与训练 anchor 零重叠）
  meta.json

跑法：python scripts/build_match_dataset_v3.py
"""
import io
import json
import os
import re
import sys
import time
import random
import collections

if getattr(sys.stdout, "encoding", "").lower() not in ("utf-8", "utf8"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUT = os.path.join(BASE, "output")
V3 = os.path.join(OUT, "model_train", "v3")

import pandas as pd
import numpy as np
import align as AL
import sales_text as STx
import text_norm as TN

# 清洗/规格/型号实现统一在 scripts/text_norm.py（与线上 vec_encoder/align 共用同一份代码）
v3_clean = TN.clean
spec_sig = TN.spec_sig
spec_conflict = TN.spec_conflict
model_tokens = TN.model_tokens
norm_bc = TN.norm_bc
nkey = TN.nkey

SEED = 42
VAL_ANCHOR_FRAC = 0.25      # 划入评估集的 anchor 比例（训练/评估 anchor 零重叠）
NEG_PER_POS = 3             # 每个正对挖几条难负
NEG_SIM_RATIO = 0.70        # 难负相似度下限 = 真对相似度 × 该比例（真对字面相似度本身可能只有 0.3）
NEG_SIM_MAX = 0.95          # 太高则可能是同一 SPU 的重复，排除
MODEL_DF_CAP = 20           # 型号 token 在自家品组中出现超过此数视为泛化 token，丢弃
SUSPECT_MARGIN = 0.02       # 同类目更相似者超过真对多少即判为疑似错标
DROP_SUSPECT_GAP = 0.10     # 分差≥此值的疑似错标直接从训练/评估中剔除
VERSION = "v3.0"
CLEAN_VERSION = TN.VERSION   # 清洗口径版本（与线上共用；改动必须重跑本脚本）

# --------------------------------------------------------------------------
# 加载自家库
# --------------------------------------------------------------------------
def load_own():
    own = pd.read_excel(os.path.join(OUT, "自家_归一化后.xlsx"), dtype=str)
    own["__s"] = own["月销量"].apply(lambda v: float(STx.eff(v) or 0)) if "月销量" in own.columns else 0.0
    own = own.sort_values("__s", ascending=False).drop_duplicates(subset=["SKUID"], keep="first")
    own["__gid"] = own["SPUID"].astype(str).str.strip()
    own["__cat"] = own.get("类目三级_归一", pd.Series("", index=own.index)).astype(str).str.strip()
    own["__brand"] = own.get("品牌（抽取）", pd.Series("", index=own.index)).astype(str).str.strip()
    own["__bc"] = own["商品条码"].apply(norm_bc) if "商品条码" in own.columns else ""

    sku2gid = dict(zip(own["SKUID"].astype(str), own["__gid"]))
    groups, gmeta = {}, {}
    for gid, g in own.groupby("__gid"):
        g = g.sort_values("__s", ascending=False)
        b = g.iloc[0]
        nm = str(b.get("商品名称") or "").strip()
        if not nm:
            continue
        groups[gid] = v3_clean(nm, b.get("__brand"))
        gmeta[gid] = {"cat": str(b.get("__cat") or ""), "brand": str(b.get("__brand") or ""),
                      "raw": nm}
    # 型号 token 的文档频率（按品组计）
    df_model = collections.Counter()
    for gid, g in groups.items():
        for m in model_tokens(g):
            df_model[m] += 1
    specific = {m for m, c in df_model.items() if c <= MODEL_DF_CAP}
    return own, sku2gid, groups, gmeta, specific


# --------------------------------------------------------------------------
# 加载竞品库（6 家标准化 + 示例竞店B美团导出）
# --------------------------------------------------------------------------
def load_comp():
    items = []
    for f in sorted(os.listdir(OUT)):
        if not (f.startswith("竞品_标准化_") and f.endswith(".xlsx")):
            continue
        d = pd.read_excel(os.path.join(OUT, f), dtype=str)
        cat = next((c for c in ("三级类目", "二级类目", "一级类目") if c in d.columns), None)
        for _, r in d.iterrows():
            nm = str(r.get("商品名称") or "").strip()
            if not nm:
                continue
            items.append({"name": nm, "store": f.replace("竞品_标准化_", "").split("_")[0],
                          "cat": str(r.get(cat) or "") if cat else "",
                          "bc": norm_bc(r.get("商品条码")), "src": f})
    p = os.path.join(BASE, "竞品_示例竞店B商品库_美团导出.xlsx")
    if os.path.exists(p):
        d = pd.read_excel(p, dtype=str)
        cat = next((c for c in ("美团后台三级类", "美团后台二级类", "美团后台一级类") if c in d.columns), None)
        for _, r in d.iterrows():
            nm = str(r.get("商品名字") or "").strip()
            if not nm:
                continue
            items.append({"name": nm, "store": "示例竞店B",
                          "cat": str(r.get(cat) or "") if cat else "", "bc": "", "src": os.path.basename(p)})
    return items


# --------------------------------------------------------------------------
# 人工确认标签
# --------------------------------------------------------------------------
def load_human(sku2gid, own_names_by_gid):
    pos, neg, report = [], [], {}
    p = os.path.join(OUT, "训练样本库.xlsx")
    if os.path.exists(p):
        ds = pd.read_excel(p, sheet_name="正样本", dtype=str)
        ds = ds[ds["自家名"].astype(str).str.strip().ne("") & ds["自家名"].notna()]
        for _, r in ds.iterrows():
            gid = sku2gid.get(str(r.get("自家ID") or "").strip(), "")
            if gid:
                pos.append({"a": str(r["竞品名"]).strip(), "gid": gid, "src": "人工确认(训练样本库)", "conf": 1.0})
        ng = pd.read_excel(p, sheet_name="难负样本", dtype=str)
        ng = ng[ng["竞品名"].notna() & ng["自家名"].notna()]
        for _, r in ng.iterrows():
            neg.append({"a": str(r["竞品名"]).strip(), "b": str(r["自家名"]).strip(),
                        "reason": "人工无对应", "human": True})
        report["训练样本库"] = {"正样本": int(len(ds)), "难负样本": int(len(ng))}

    p = os.path.join(OUT, "对齐汇总库.xlsx")
    if os.path.exists(p):
        d1 = pd.read_excel(p, sheet_name="SPU匹配层", dtype=str)
        hit = 0
        for _, r in d1.iterrows():
            nm = nkey(r.get("自家品名"))
            gid = own_names_by_gid.get(nm)
            if gid:
                pos.append({"a": str(r.get("竞品商品名") or "").strip(), "gid": gid,
                            "src": f"对齐汇总库SPU({r.get('确认来源')})", "conf": 0.95})
                hit += 1
        d2 = pd.read_excel(p, sheet_name="SKU匹配层", dtype=str)
        hit2 = 0
        for _, r in d2.iterrows():
            gid = sku2gid.get(str(r.get("自家SKUID") or "").strip(), "")
            if gid:
                pos.append({"a": str(r.get("竞品商品名") or "").strip(), "gid": gid,
                            "src": f"对齐汇总库SKU({r.get('匹配依据')})", "conf": 0.95})
                hit2 += 1
        report["对齐汇总库"] = {"SPU层命中": hit, "SKU层命中": hit2,
                               "SPU层总行": int(len(d1)), "SKU层总行": int(len(d2))}
    return pos, neg, report


# --------------------------------------------------------------------------
def main():
    rng = random.Random(SEED)   # 局部 RNG，保证可复现（不污染全局状态）
    os.makedirs(V3, exist_ok=True)
    t0 = time.time()
    print("=" * 88)
    print("加载自家库 / 竞品库")
    own, sku2gid, groups, gmeta, specific = load_own()
    print(f"  自家 SKU {len(own)}，品组 {len(groups)}，可用型号 token {len(specific)}")
    comp = load_comp()
    print(f"  竞品条目 {len(comp)}（{len(set(x['store'] for x in comp))} 家店）")

    own_names_by_gid = {}
    for gid, g in groups.items():
        own_names_by_gid[nkey(gmeta[gid]["raw"])] = gid
    # 兜底：用自家全名表
    for _, r in own.iterrows():
        own_names_by_gid.setdefault(nkey(r.get("商品名称")), r["__gid"])

    print("\n加载人工确认标签")
    human_pos, human_neg, hrep = load_human(sku2gid, own_names_by_gid)
    for k, v in hrep.items():
        print(f"  {k}: {v}")
    print(f"  人工正对原始 {len(human_pos)}，人工难负 {len(human_neg)}")

    # ---------------- L1 硬证据正对 ----------------
    print("\nL1a 条码相同 → 正对")
    bc2gid = collections.defaultdict(set)
    for gid, g in own.groupby("__gid"):
        for bc in g["__bc"]:
            if bc:
                bc2gid[bc].add(gid)
    bc2gid = {b: list(s)[0] for b, s in bc2gid.items() if len(s) == 1}   # 歧义条码丢弃
    l1a = []
    for it in comp:
        gid = bc2gid.get(it["bc"])
        if gid and gid in groups:
            l1a.append({"a": it["name"], "gid": gid, "src": "L1a条码", "conf": 1.0})
    print(f"  命中 {len(l1a)} 条（自家唯一条码 {len(bc2gid)}）")

    print("L1b 品牌+型号全等 → 正对")
    model2gids = collections.defaultdict(set)
    for gid, g in groups.items():
        for m in model_tokens(g):
            if m in specific:
                model2gids[m].add(gid)
    model2gids = {m: s for m, s in model2gids.items() if len(s) <= 5}   # 一个型号最多对 5 个品组
    l1b, seen_l1b = [], set()
    for it in comp:
        toks = model_tokens(it["name"]) & set(model2gids)
        if not toks:
            continue
        for m in sorted(toks):                      # sorted：避免 set 迭代顺序导致不可复现
            for gid in sorted(model2gids[m]):
                k = (nkey(it["name"]), gid)
                if k in seen_l1b:
                    continue
                seen_l1b.add(k)
                l1b.append({"a": it["name"], "gid": gid, "src": f"L1b型号:{m}", "conf": 0.9})
    print(f"  命中 {len(l1b)} 条（可用型号 {len(model2gids)} 个）")

    # ---------------- 汇总正对 ----------------
    all_pos = human_pos + l1a + l1b
    seen, pos = set(), []
    for r in all_pos:
        a = str(r["a"]).strip()
        if not a or r["gid"] not in groups:
            continue
        k = (nkey(a), r["gid"])
        if k in seen:
            continue
        seen.add(k)
        pos.append(r)
    print(f"\n正对合计（去重后）{len(pos)} 条")

    # ---------------- 训练/评估切分（anchor 零重叠）----------------
    # 说明：自家库已按 SPU 归一化（同一 SPU 内所有 SKU 共用同一个「商品名称」，
    # 实测 3708 个品组中 0 个存在同品多挂法），故无法从自家库做"同品挂法"增广。
    # 正对规模只能靠 L1/L3 硬证据；扩量需另接 L2 结构等价 / L4 LLM 辅助标注。
    all_anchor_keys = sorted(set(nkey(r["a"]) for r in pos))
    rng.shuffle(all_anchor_keys)
    eval_set = set(a for a in all_anchor_keys if rng.random() < VAL_ANCHOR_FRAC)
    eval_rows = [r for r in pos if nkey(r["a"]) in eval_set]
    train_rows = [r for r in pos if nkey(r["a"]) not in eval_set]

    # ---------------- 索引与检索基座（自检 + 难负挖掘共用）----------------
    from sklearn.feature_extraction.text import TfidfVectorizer
    gids = list(groups.keys())
    gidx = {g: i for i, g in enumerate(gids)}
    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=2)
    G = vec.fit_transform([groups[g] for g in gids])
    cat2gids = collections.defaultdict(list)
    for gid in gids:
        if gmeta[gid]["cat"]:
            cat2gids[gmeta[gid]["cat"]].append(gid)
    gid2models = {g: model_tokens(groups[g]) & specific for g in gids}
    gid2bc = collections.defaultdict(set)
    for _, r in own.iterrows():
        if r["__bc"]:
            gid2bc[r["__gid"]].add(r["__bc"])

    def banned(gid_cand, a_models, gid_pos):
        """排除疑似同品的候选：共享型号 / 共享条码。"""
        if gid_cand == gid_pos:
            return True
        if gid2models[gid_cand] & a_models:
            return True
        if gid2bc[gid_cand] & gid2bc[gid_pos]:
            return True
        return False

    def sims_of(a, cat):
        cand = cat2gids.get(cat, [])
        if len(cand) < 2:
            return None, None, None
        av = vec.transform([a])
        ci = [gidx[g] for g in cand]
        s = np.asarray((G[ci] @ av.T).todense()).ravel()
        return cand, {g: i for i, g in enumerate(cand)}, s

    # ---------------- 标签自检：疑似错标的正对 ----------------
    # 判据：同类目下存在比"真对"字面更相似的候选 → 该正对很可能标错（人工确认也会错）
    print("\n标签自检（同类目更相似者 > 真对 → 疑似错标）")
    suspects, drop_keys = [], set()
    for r in pos:
        a = v3_clean(str(r["a"]).strip())
        gid_pos = r["gid"]
        cand, cpos, sims = sims_of(a, gmeta[gid_pos]["cat"])
        if cand is None or gid_pos not in cpos:
            continue
        a_models = model_tokens(a) & specific
        sp = float(sims[cpos[gid_pos]])
        best_g, best_s = None, -1.0
        for g in cand:
            if banned(g, a_models, gid_pos):
                continue
            s = float(sims[cpos[g]])
            if s > best_s:
                best_g, best_s = g, s
        if best_g and best_s > sp + SUSPECT_MARGIN:
            rec = {"a": a, "p": groups[gid_pos], "gid": gid_pos, "src": r["src"],
                   "sim_pos": round(sp, 4), "sim_other": round(best_s, 4),
                   "gap": round(best_s - sp, 4), "other": groups[best_g], "other_gid": best_g}
            suspects.append(rec)
            if (best_s - sp) >= DROP_SUSPECT_GAP:
                drop_keys.add((nkey(r["a"]), gid_pos))
    print(f"  疑似错标 {len(suspects)} / {len(pos)} = {len(suspects)/max(len(pos),1):.1%}")
    print(f"  其中分差≥{DROP_SUSPECT_GAP} 的高置信错标 {len(drop_keys)} 条 → 从训练/评估中剔除")
    if drop_keys:
        before_t, before_e = len(train_rows), len(eval_rows)
        train_rows = [r for r in train_rows if (nkey(r["a"]), r["gid"]) not in drop_keys]
        eval_rows = [r for r in eval_rows if (nkey(r["a"]), r["gid"]) not in drop_keys]
        print(f"  剔除后：训练对 {before_t}→{len(train_rows)}，评估对 {before_e}→{len(eval_rows)}")

    ev_a = set(nkey(r["a"]) for r in eval_rows)
    tr_a = set(nkey(r["a"]) for r in train_rows)
    overlap = ev_a & tr_a
    print(f"\n训练 anchor {len(tr_a)}（对 {len(train_rows)}）")
    print(f"评估 anchor {len(ev_a)}（对 {len(eval_rows)}），重叠 {len(overlap)}")
    if overlap:
        raise SystemExit(f"❌ 训练/评估 anchor 重叠 {len(overlap)} 条，拒绝产出")
    print("  ✅ anchor 零重叠断言通过")

    # ---------------- L5 难负挖掘 ----------------
    print("\nL5 难负挖掘（同类目 + 与真对近分 + 排除同品）")
    # 同类别目下"其它正对的自家名"（Tier B 用：换负例，保证是不同 SPU）
    pos_by_cat = collections.defaultdict(list)
    for r in train_rows:
        c = gmeta[r["gid"]]["cat"]
        if c:
            pos_by_cat[c].append(r["gid"])

    triplets = []
    hard_eval = []
    n_a = n_b = 0
    for r in train_rows:
        a = v3_clean(str(r["a"]).strip())
        gid_pos = r["gid"]
        cat = gmeta[gid_pos]["cat"]
        cand, cpos, sims = sims_of(a, cat)
        if cand is None:
            continue
        a_models = model_tokens(a) & specific
        order = np.argsort(sims)[::-1]
        sim_pos = float(sims[cpos[gid_pos]]) if gid_pos in cpos else 0.0
        # 难负的相似度下限：跟着真对走（真对本身字面相似度可能只有 0.3）
        floor = max(0.10, sim_pos * NEG_SIM_RATIO)
        picked = []
        # Tier A：同类目 + 相似度不低于真对的一定比例 + 规格冲突
        for k in order:
            if len(picked) >= NEG_PER_POS:
                break
            g = cand[k]
            s = float(sims[k])
            if s > NEG_SIM_MAX or s < floor or banned(g, a_models, gid_pos):
                continue
            conf = spec_conflict(a, groups[g])
            if not conf:
                continue
            picked.append({"anchor": a, "positive": groups[gid_pos], "negative": groups[g],
                           "tier": "A规格冲突", "reason": ";".join(conf), "sim": round(s, 4),
                           "sim_pos": round(sim_pos, 4), "gid_pos": gid_pos, "gid_neg": g})
            n_a += 1
        # Tier B：同类目下另一条正对的自家品组（换负例，必为不同 SPU）
        if len(picked) < NEG_PER_POS:
            others = sorted(g for g in pos_by_cat[cat] if not banned(g, a_models, gid_pos))
            rng.shuffle(others)
            for g in others:
                if len(picked) >= NEG_PER_POS:
                    break
                s = float(sims[cpos[g]]) if g in cpos else 0.0
                if s > NEG_SIM_MAX:
                    continue
                picked.append({"anchor": a, "positive": groups[gid_pos], "negative": groups[g],
                               "tier": "B同类目换负", "reason": "同类目不同SPU",
                               "sim": round(s, 4), "sim_pos": round(sim_pos, 4),
                               "gid_pos": gid_pos, "gid_neg": g})
                n_b += 1
        triplets.extend(picked)
    print(f"  生成三元组 {len(triplets)} 条（Tier A 规格冲突 {n_a} + Tier B 同类目换负 {n_b}），"
          f"覆盖 {len(set(t['anchor'] for t in triplets))} 个 anchor")

    for h in human_neg:
        hard_eval.append({"a": v3_clean(h["a"]), "b": v3_clean(h["b"]),
                          "reason": h["reason"], "human": True})
    for t in triplets:
        hard_eval.append({"a": t["anchor"], "b": t["negative"], "reason": t["reason"], "human": False})
    print(f"  难负对合计 {len(hard_eval)} 条（人工 {len(human_neg)} + 挖掘 {len(triplets)}）")

    # ---------------- 落盘 ----------------
    def dump(name, rows):
        # 排序后落盘：输出与内部迭代顺序无关，保证逐字节可复现
        rows = sorted(rows, key=lambda x: json.dumps(x, ensure_ascii=False, sort_keys=True))
        with open(os.path.join(V3, name), "w", encoding="utf-8") as f:
            for x in rows:
                f.write(json.dumps(x, ensure_ascii=False) + "\n")
        print(f"  写 {name}: {len(rows)} 行")

    print("\n落盘")
    dump("groups.jsonl", [{"gid": g, "t": groups[g], "cat": gmeta[g]["cat"],
                           "brand": gmeta[g]["brand"]} for g in gids])
    dump("pos_pairs.jsonl", [{"a": v3_clean(r["a"]), "p": r.get("aug") or groups[r["gid"]],
                              "gid": r["gid"], "src": r["src"], "conf": r["conf"]} for r in train_rows])
    dump("eval_pairs.jsonl", [{"q": v3_clean(r["a"]), "gid": r["gid"],
                               "src": r["src"], "conf": r["conf"]} for r in eval_rows])
    dump("triplet_v3.jsonl", triplets)
    dump("hard_neg_v3.jsonl", hard_eval)
    dump("suspect_pos.jsonl", suspects)

    # ---------------- meta ----------------
    srcs = {}
    for fn in ("自家_归一化后.xlsx", "训练样本库.xlsx", "对齐汇总库.xlsx",
               "竞品_示例竞店B商品库_美团导出.xlsx"):
        p = os.path.join(OUT, fn) if fn != "竞品_示例竞店B商品库_美团导出.xlsx" else os.path.join(BASE, fn)
        if os.path.exists(p):
            st = os.stat(p)
            srcs[fn] = {"size": st.st_size, "mtime": int(st.st_mtime)}
    for f in sorted(os.listdir(OUT)):
        if f.startswith("竞品_标准化_") and f.endswith(".xlsx"):
            st = os.stat(os.path.join(OUT, f))
            srcs[f] = {"size": st.st_size, "mtime": int(st.st_mtime)}

    meta = {
        "version": VERSION,
        "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "script": "scripts/build_match_dataset_v3.py",
        "task": "竞品商品名 → 自家品组（SPUID）对齐",
        "recipe": {
            "positive": ["L1a 条码相同", "L1b 品牌+型号全等", "L3 人工确认"],
            "negative": ["TierA 同类目 + 相似度≥真对×%.2f + 规格冲突" % NEG_SIM_RATIO,
                         "TierB 同类目另一正对的自家品组（保证不同 SPU）",
                         "两者均排除共享型号/条码的候选（防假负）"],
            "clean": "v3_clean：去品牌与营销词，**保留规格数字**",
            "spec_norm": "长度→mm、重量→g、体积→ml、电压→V、功率→W、电流→A",
            "split": "anchor 级 %.0f%% 划入评估集；训练/评估 anchor 零重叠（硬断言）" % (VAL_ANCHOR_FRAC * 100),
        },
        "counts": {
            "own_sku": int(len(own)), "own_groups": len(groups), "comp_items": len(comp),
            "human_pos_raw": len(human_pos), "human_neg": len(human_neg),
            "l1a_barcode": len(l1a), "l1b_model": len(l1b),
            "pos_total": len(pos), "pos_train": len(train_rows),
            "eval_queries": len(eval_rows),
            "suspect_pos": len(suspects), "suspect_dropped": len(drop_keys),
            "triplets": len(triplets), "hard_neg": len(hard_eval),
            "train_anchors": len(tr_a), "eval_anchors": len(ev_a),
            "anchor_overlap": len(overlap),
        },
        "human_label_report": hrep,
        "sources": srcs,
        "elapsed_sec": round(time.time() - t0, 1),
    }
    with open(os.path.join(V3, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print(f"  写 meta.json")

    print("\n" + "=" * 88)
    print("汇总")
    for k, v in meta["counts"].items():
        print(f"  {k:20s} {v}")
    print(f"\n完成，用时 {meta['elapsed_sec']}s → {V3}")


if __name__ == "__main__":
    main()

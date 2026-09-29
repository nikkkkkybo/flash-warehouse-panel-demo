# -*- coding: utf-8 -*-
"""训练数据 v4：v3 基础上接 复核裁决 + L2 机判 + L4 LLM 裁决（LLM 可后补）。

输入：
  v3/  pos_pairs / eval_pairs / groups / hard_neg_v3（人工难负）+ meta
  v4/  review_verdicts.jsonl（疑似错标复核：keep/drop/flip）
       l2_pos.jsonl（L2 结构等价机判正对 222）
       l2_neg.jsonl（L2 规格冲突机判难负 11991）
       llm_verdicts.jsonl（L4 LLM 裁决，可后补——有则并入，无则跳过）

处理规则：
  - v3 正对中，复核判定 drop → 剔除；机器自动来源 → 必须 LLM 确认才保留
  - LLM same=True → 正对；same=False → 难负（进损失 Tier C + 评估难负集）
  - 评估集 anchor 级 25% 切分（零重叠硬断言），跨店统计写入 meta

产物（output/model_train/v4/）：
  groups.jsonl / pos_pairs.jsonl / eval_pairs.jsonl / triplet_v4.jsonl /
  hard_neg_v4.jsonl / meta.json

跑法：python scripts/build_match_dataset_v4.py
"""
import collections
import io
import json
import os
import random
import re
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUT = os.path.join(BASE, "output")
V3D = os.path.join(OUT, "model_train", "v3")
V4 = os.path.join(OUT, "model_train", "v4")

import numpy as np
import text_norm as TN

SEED = 42
VAL_ANCHOR_FRAC = 0.25
NEG_PER_POS = 3
SIM_MAX = 0.95
VERSION = "v4.0"


def jl(path):
    p = path if os.path.isabs(path) else os.path.join(V4, path)
    if not os.path.exists(p):
        return []
    return [json.loads(x) for x in open(p, encoding="utf-8") if x.strip()]


def jl3(fn):
    return jl(os.path.join(V3D, fn))


def dump(fn, rows):
    rows = sorted(rows, key=lambda x: json.dumps(x, ensure_ascii=False, sort_keys=True))
    with open(os.path.join(V4, fn), "w", encoding="utf-8") as f:
        for x in rows:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    print(f"  写 {fn}: {len(rows)} 行")


def load_review():
    rev = {}
    for r in jl("review_verdicts.jsonl"):
        rev[(r["a_key"], r["gid"])] = r
    return rev


def load_llm():
    llm_pos, llm_neg, llm_stat = {}, {}, collections.Counter()
    for r in jl("llm_verdicts.jsonl"):
        key = (r["key"][0], r["gid"])
        if r.get("same") is True:
            llm_pos[key] = r
        elif r.get("same") is False:
            llm_neg[key] = r
        else:
            llm_stat["null"] += 1
    llm_stat["pos"] = len(llm_pos)
    llm_stat["neg"] = len(llm_neg)
    return llm_pos, llm_neg, llm_stat


def main():
    rng = random.Random(SEED)
    os.makedirs(V4, exist_ok=True)
    t0 = time.time()
    groups_rows = jl3("groups.jsonl")
    groups = {r["gid"]: r["t"] for r in groups_rows}
    gcat = {r["gid"]: r.get("cat", "") for r in groups_rows}
    print(f"品组 {len(groups)}")

    rev = load_review()
    llm_pos, llm_neg, llm_stat = load_llm()
    print(f"复核裁决 {len(rev)} 条，LLM 裁决 {dict(llm_stat)}")

    # ---------- 正对池 ----------
    pos, stats = [], collections.Counter()
    dropped = flipped = machine_pending = 0
    for fn, key in (("pos_pairs.jsonl", "a"), ("eval_pairs.jsonl", "q")):
        for r in jl3(fn):
            a = r[key]
            gid = r["gid"]
            if gid not in groups:
                continue
            ak = TN.nkey(a)
            src = r.get("src", "")
            vd = rev.get((ak, gid))
            if vd and vd.get("verdict") == "drop":
                dropped += 1
                continue
            if vd and vd.get("verdict") == "flip" and vd.get("other_gid"):
                gid = vd["other_gid"]
                flipped += 1
            if "机器自动" in src:
                if (ak, gid) in llm_pos:
                    conf = 0.8
                    stats["机器自动+LLM确认"] += 1
                else:
                    machine_pending += 1
                    continue
            else:
                conf = float(r.get("conf") or 0.95)
            pos.append({"a": a, "p": groups[gid], "gid": gid, "src": src, "conf": conf})
    stats["v3保留"] = len(pos)
    stats["复核剔除"] = dropped
    stats["复核改判"] = flipped
    stats["机器自动待LLM"] = machine_pending

    n_before = len(pos)
    for r in jl("l2_pos.jsonl"):
        if r["gid"] in groups:
            pos.append({"a": r["a"], "p": r["p"], "gid": r["gid"],
                        "src": f"L2:{r.get('reason', '')}", "conf": 0.85})
    stats["L2机判"] = len(pos) - n_before

    n_before = len(pos)
    seen = set((TN.nkey(r["a"]), r["gid"]) for r in pos)
    for (ak, gid), r in llm_pos.items():
        if gid in groups and (ak, gid) not in seen:
            pos.append({"a": r["a"], "p": groups[gid], "gid": gid,
                        "src": "L4:LLM同款", "conf": 0.8})
            seen.add((ak, gid))
    stats["LLM同款"] = len(pos) - n_before

    # 同 (anchor,gid) 去重，同 anchor 多 gid 保留（一竞品可对同品组唯一；多 gid 属可疑，取 conf 最高）
    by_a = collections.defaultdict(list)
    for r in pos:
        by_a[TN.nkey(r["a"])].append(r)
    clean_pos = []
    multi_gid = 0
    for ak, rs in by_a.items():
        gids = set(r["gid"] for r in rs)
        if len(gids) > 1:
            multi_gid += 1
            best = max(rs, key=lambda r: (r["conf"], r["gid"]))
            clean_pos.append(best)
        else:
            clean_pos.append(rs[0])
    pos = clean_pos
    stats["同anchor多gid冲突"] = multi_gid
    stats["正对合计"] = len(pos)
    print("正对池：", dict(stats))
    if not pos:
        raise SystemExit("无正对")

    # ---------- 切分（anchor 零重叠）----------
    aks = sorted(by_a.keys())
    rng.shuffle(aks)
    eval_set = set(a for a in aks if rng.random() < VAL_ANCHOR_FRAC)
    eval_rows = [r for r in pos if TN.nkey(r["a"]) in eval_set]
    train_rows = [r for r in pos if TN.nkey(r["a"]) not in eval_set]
    ev_a = set(TN.nkey(r["a"]) for r in eval_rows)
    tr_a = set(TN.nkey(r["a"]) for r in train_rows)
    overlap = ev_a & tr_a
    if overlap:
        raise SystemExit(f"❌ 训练/评估 anchor 重叠 {len(overlap)}，拒绝产出")
    print(f"训练 {len(train_rows)} 对 / 评估 {len(eval_rows)} 对，重叠 0 ✅")

    # ---------- 三元组 ----------
    l2_neg = jl("l2_neg.jsonl")
    neg_by_a = collections.defaultdict(list)
    for r in l2_neg:
        neg_by_a[TN.nkey(r["a"])].append(r)
    for (ak, gid), r in llm_neg.items():
        neg_by_a[ak].append({"a": r["a"], "p": r.get("p", ""), "gid": gid,
                             "reason": f"LLM非同款:{r.get('reason', '')}", "sim": r.get("sim")})
    pos_by_cat = collections.defaultdict(list)
    for r in train_rows:
        if gcat.get(r["gid"]):
            pos_by_cat[gcat[r["gid"]]].append(r["gid"])

    triplets = []
    nA = nB = nC = 0
    for r in train_rows:
        a, gid_pos = r["a"], r["gid"]
        picked = []
        for nr in neg_by_a[TN.nkey(a)]:
            g = nr["gid"]
            if g == gid_pos or g not in groups or g == "":
                continue
            if nr.get("sim") and float(nr["sim"]) > SIM_MAX:
                continue
            picked.append({"anchor": a, "positive": groups[gid_pos], "negative": groups[g],
                           "tier": "C规则/LLM难负", "reason": nr.get("reason", ""),
                           "sim": nr.get("sim"), "gid_pos": gid_pos, "gid_neg": g})
            nC += 1
            if len(picked) >= NEG_PER_POS:
                break
        if len(picked) < NEG_PER_POS and gcat.get(gid_pos):
            others = [g for g in pos_by_cat[gcat[gid_pos]] if g != gid_pos]
            rng.shuffle(others)
            for g in others:
                if len(picked) >= NEG_PER_POS:
                    break
                picked.append({"anchor": a, "positive": groups[gid_pos], "negative": groups[g],
                               "tier": "B同类目换负", "reason": "同类目不同SPU",
                               "gid_pos": gid_pos, "gid_neg": g})
                nB += 1
        triplets.extend(picked)
    print(f"三元组 {len(triplets)}（C 规则/LLM {nA + nC} + B 同类目 {nB}）")

    # ---------- 难负评估集 ----------
    hard = []
    for r in jl3("hard_neg_v3.jsonl"):
        if r.get("human"):
            hard.append({"a": r["a"], "b": r["b"], "reason": r.get("reason", ""), "human": True})
    for r in l2_neg:
        hard.append({"a": r["a"], "b": r["p"], "reason": r.get("reason", ""), "human": False})
    for (ak, gid), r in llm_neg.items():
        if r.get("p"):
            hard.append({"a": r["a"], "b": r["p"], "reason": f"LLM:{r.get('reason', '')}", "human": False})
    print(f"难负评估集 {len(hard)}")

    # ---------- 落盘 ----------
    stores = collections.Counter()
    dump("groups.jsonl", groups_rows)
    dump("pos_pairs.jsonl", train_rows)
    dump("eval_pairs.jsonl", eval_rows)
    dump("triplet_v4.jsonl", triplets)
    dump("hard_neg_v4.jsonl", hard)

    src_f = {}
    for fn in ("review_verdicts.jsonl", "l2_pos.jsonl", "l2_neg.jsonl", "llm_verdicts.jsonl"):
        p = os.path.join(V4, fn)
        if os.path.exists(p):
            st = os.stat(p)
            src_f[fn] = {"size": st.st_size, "mtime": int(st.st_mtime)}
    for fn in ("pos_pairs.jsonl", "eval_pairs.jsonl", "meta.json"):
        p = os.path.join(V3D, fn)
        if os.path.exists(p):
            st = os.stat(p)
            src_f["v3/" + fn] = {"size": st.st_size, "mtime": int(st.st_mtime)}

    meta = {"version": VERSION, "generated_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "script": "scripts/build_match_dataset_v4.py",
            "base": "v3（build_match_dataset_v3.py 产物）",
            "inputs": {"review_verdicts": len(rev), "llm_verdicts": dict(llm_stat),
                       "l2_pos": stats.get("L2机判", 0)},
            "counts": {k: v for k, v in stats.items()},
            "split": {"train_pairs": len(train_rows), "eval_pairs": len(eval_rows),
                      "train_anchors": len(tr_a), "eval_anchors": len(ev_a),
                      "anchor_overlap": 0, "val_frac": VAL_ANCHOR_FRAC},
            "triplets": {"total": len(triplets), "tierC": nA + nC, "tierB": nB},
            "hard_neg": len(hard), "extra_sources": src_f,
            "elapsed_sec": round(time.time() - t0, 1)}
    with open(os.path.join(V4, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, ensure_ascii=False, indent=1)
    print("  写 meta.json")
    print(f"\n完成 {meta['elapsed_sec']}s → {V4}")


if __name__ == "__main__":
    main()

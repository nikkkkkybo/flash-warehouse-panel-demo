# -*- coding: utf-8 -*-
"""候选挖掘（v4 扩量第 1 步）：竞品 → 自家品组候选对，L2 预判 + LLM 队列。

输入：自家品组库（v3/groups.jsonl）+ 竞品库（build_match_dataset_v3.load_comp）
排除：v3 已确认正对的 anchor（不重复标注）
输出（output/model_train/v4/）：
  cand_all.jsonl   全部候选（检索 top-K，含 sim），溯源用
  l2_pos.jsonl     L2 结构等价 → 机器正对（不花钱直接用）
  l2_neg.jsonl     L2 规格冲突 → 机器难负
  llm_queue.jsonl  L2 覆盖不了 → L4 LLM 打标队列（按 sim 降序，限额）
跑法：python scripts/mine_candidates.py [--cap 6000]
"""
import collections
import io
import json
import os
import sys

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUT = os.path.join(BASE, "output")
V3 = os.path.join(OUT, "model_train", "v3")
V4 = os.path.join(OUT, "model_train", "v4")

import numpy as np
import text_norm as TN
import name_parse as NP
import build_match_dataset_v3 as V3B

TOPK = 10
SIM_MIN, SIM_MAX = 0.25, 0.95
DEFAULT_CAP = 6000


def jl(fn, d=V3):
    p = os.path.join(d, fn)
    if not os.path.exists(p):
        return []
    return [json.loads(x) for x in open(p, encoding="utf-8") if x.strip()]


def dump(fn, rows):
    rows = sorted(rows, key=lambda x: json.dumps(x, ensure_ascii=False, sort_keys=True))
    with open(os.path.join(V4, fn), "w", encoding="utf-8") as f:
        for x in rows:
            f.write(json.dumps(x, ensure_ascii=False) + "\n")
    print(f"  写 {fn}: {len(rows)} 行")


def main():
    cap = DEFAULT_CAP
    if "--cap" in sys.argv:
        cap = int(sys.argv[sys.argv.index("--cap") + 1])
    os.makedirs(V4, exist_ok=True)

    rows = jl("groups.jsonl")
    groups = {r["gid"]: r["t"] for r in rows}
    gids = [r["gid"] for r in rows]
    print(f"品组 {len(gids)}")

    have = set()
    for fn, key in (("pos_pairs.jsonl", "a"), ("eval_pairs.jsonl", "q")):
        for r in jl(fn):
            have.add(TN.nkey(r[key]))
    print(f"v3 已有正对 anchor {len(have)} 个（跳过）")

    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.metrics.pairwise import linear_kernel
    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=2)
    G = vec.fit_transform([groups[g] for g in gids])

    comp = V3B.load_comp()
    print(f"竞品条目 {len(comp)}")

    l2_pos, l2_neg, queue, n_skip, n_low = [], [], [], 0, 0
    seen_pair = set()
    queue_keys = set()
    for it in comp:
        a_raw = it["name"]
        ak = TN.nkey(a_raw)
        if ak in have:
            n_skip += 1
            continue
        a = TN.clean(a_raw)
        if len(a) < 4:
            continue
        av = vec.transform([a])
        s = linear_kernel(G, av).ravel()
        order = np.argsort(s)[::-1][:TOPK]
        for k in order:
            sim = float(s[k])
            if sim < SIM_MIN or sim > SIM_MAX:
                if sim < SIM_MIN:
                    n_low += 1
                break
            gid = gids[k]
            pk = (ak, gid)
            if pk in seen_pair:
                continue
            seen_pair.add(pk)
            p = groups[gid]
            v, why = NP.equivalent(a, p)
            rec = {"a": a, "p": p, "gid": gid, "sim": round(sim, 4),
                   "store": it["store"], "why": why}
            if v == "pos":
                rec["reason"] = why
                l2_pos.append(rec)
            elif v == "neg":
                rec["reason"] = why
                l2_neg.append(rec)
            else:
                qk = (ak, gid)
                if qk not in queue_keys and len(queue) < cap and sim >= 0.30:
                    queue_keys.add(qk)
                    rec["brand"] = it.get("brand", "") if isinstance(it.get("brand", ""), str) else ""
                    queue.append(rec)
    print(f"跳过已有 anchor {n_skip}，低于下限 {n_low}")
    print(f"L2 正对 {len(l2_pos)}，L2 难负 {len(l2_neg)}，LLM 队列 {len(queue)}")
    print(f"LLM 队列 anchor 唯一 {len(set(TN.nkey(x['a']) for x in queue))}")

    dump("cand_stats.json", [{"l2_pos": len(l2_pos), "l2_neg": len(l2_neg),
                              "llm_queue": len(queue), "skipped_anchors": n_skip}])
    dump("l2_pos.jsonl", l2_pos)
    dump("l2_neg.jsonl", l2_neg)
    dump("llm_queue.jsonl", queue)


if __name__ == "__main__":
    main()

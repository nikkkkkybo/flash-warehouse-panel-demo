# -*- coding: utf-8 -*-
"""疑似错标复核（质检方案 §7.2 遗留任务②）。

对象：
  A. output/model_train/v3/suspect_pos.jsonl   466 条疑似错标正对
  B. output/对齐汇总库.xlsx SPU匹配层 确认来源=机器自动 165 条

方法（结构化判据，不靠"字面更像"单指标）：
  1. 对 (竞品名, 原标注自家名) 与 (竞品名, 更像自家名) 各做 L2 等价判定；
  2. 证据分级：条码 > 人工确认 > 型号 > 机器自动。弱证据源被"更像"质疑时
     仅降级不剔除；强证据源必须有双向结构证据才改判；
  3. 判定建议：keep(保留) / drop(剔除) / flip(改判更像者) / manual(待人工)。

产物：
  output/复核_疑似错标_20260909.xlsx          人审工作表（两个 sheet）
  output/model_train/v4/review_verdicts.jsonl 机器裁决（v4 生成器消费）
跑法：python scripts/review_suspects.py
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
import pandas as pd
import name_parse as NP
import text_norm as TN

# 证据强度：数值越大越可信
SRC_RANK = [
    ("条码", 5), ("人工确认", 5), ("人工手动指定", 5),
    ("人工无对应", 5), ("型号", 4), ("类目+品牌+型号+规格", 4),
    ("机器自动", 2), ("L1b", 4),
]


def src_rank(src):
    for k, r in SRC_RANK:
        if k in str(src):
            return r
    return 3


def load_groups():
    rows = [json.loads(x) for x in open(os.path.join(V3, "groups.jsonl"), encoding="utf-8") if x.strip()]
    groups = {r["gid"]: r["t"] for r in rows}
    gcat = {r["gid"]: r.get("cat", "") for r in rows}
    return rows, groups, gcat


def build_index(gnames):
    from sklearn.feature_extraction.text import TfidfVectorizer
    vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4), min_df=2)
    G = vec.fit_transform(gnames)
    return vec, G


def better_candidates(a, gid_pos, vec, G, gids, gcat, topn=5):
    """同类目下比 gid_pos 更像 a 的品组（供 165 复核；466 已带 other 字段）。"""
    from sklearn.metrics.pairwise import linear_kernel
    cat = gcat.get(gid_pos, "")
    cand = [g for g in gids if gcat.get(g, "") == cat and g != gid_pos]
    if len(cand) < 2:
        return []
    av = vec.transform([a])
    ci = [gids.index(g) for g in cand]
    s = linear_kernel(G[ci], av).ravel()
    order = np.argsort(s)[::-1][:topn]
    return [(cand[k], float(s[k])) for k in order]


def judge(a, p, other, other_sim, src):
    """返回 (建议, 证据说明)。a/p/other 均为清洗后文本。"""
    v_p, why_p = NP.equivalent(a, p)
    v_o, why_o = NP.equivalent(a, other) if other else (None, "")
    rank = src_rank(src)
    ev = [f"L2(竞品,原标)={NP.verdict_label(v_p)}:{why_p}" if why_p else "L2(竞品,原标)=不确定",
          f"L2(竞品,更像)={NP.verdict_label(v_o)}:{why_o}" if other and why_o else
          (f"L2(竞品,更像)={NP.verdict_label(v_o)}" if other else "无更像者")]
    ev.append(f"字面分差={other_sim}" if other_sim is not None else "")
    if v_p == "pos":
        return "keep", ";".join(x for x in ev if x), v_p
    if v_p == "neg":
        return "drop", ";".join(x for x in ev if x) + " ← 原对规格冲突", v_p
    if v_o == "pos" and v_p != "pos":
        if rank >= 5:
            # 强证据（条码/人工）被挑战：只在"更像者结构等价 + 原对规格冲突"时才改判，
            # 条码/人工一般比字面更像可信，交人工复核
            return "manual", ";".join(x for x in ev if x) + " ← 强证据被更像者挑战，请人工裁决", None
        return "flip", ";".join(x for x in ev if x), v_o
    if other_sim is not None and other_sim > 0.55 and rank <= 2:
        return "manual", ";".join(x for x in ev if x), None
    return "keep", ";".join(x for x in ev if x), None


def main():
    os.makedirs(V4, exist_ok=True)
    rows, groups, gcat = load_groups()
    gids = [r["gid"] for r in rows]
    vec, G = build_index([r["t"] for r in rows])

    verdicts = {}
    out_a, out_b = [], []

    print("== A. suspect_pos 466 条复核 ==")
    for r in [json.loads(x) for x in open(os.path.join(V3, "suspect_pos.jsonl"), encoding="utf-8") if x.strip()]:
        a, p = r["a"], r["p"]
        other = r.get("other", "")
        adv, ev, _ = judge(a, p, other, r.get("gap"), r.get("src", ""))
        verdicts[(TN.nkey(a), r["gid"])] = {"verdict": adv, "ev": ev}
        out_a.append({"判定建议": adv, "竞品名": a, "原标注自家名": p, "更像者": other,
                      "分差": r.get("gap"), "来源": r.get("src", ""), "证据": ev})

    print("== B. 对齐汇总库 机器自动 SPU 165 条复核 ==")
    d1 = pd.read_excel(os.path.join(OUT, "对齐汇总库.xlsx"), sheet_name="SPU匹配层", dtype=str)
    ma = d1[d1["确认来源"] == "机器自动"]
    for _, r in ma.iterrows():
        a_raw = str(r.get("竞品商品名") or "").strip()
        own_name = str(r.get("自家品名") or "").strip()
        # 汇总库的 SPU_O ID 与 groups 的数字 SPUID 不同源，用自家品名回查 gid
        nk = TN.nkey(own_name)
        gid = next((g for g, t in groups.items() if TN.nkey(t) == nk), "")
        if not gid:
            # 品组名与汇总库品名可能不一致（SPU 内主挂法），退回字符串包含
            gid = next((g for g, t in groups.items()
                        if nk[:12] and nk[:12] in TN.nkey(t)), "")
        if not gid:
            out_b.append({"判定建议": "manual", "竞品名": TN.clean(a_raw), "原标注自家名": own_name,
                          "更像者": "", "分差": "", "来源": "机器自动", "证据": "品组库中找不到该自家品名"})
            continue
        a = TN.clean(a_raw)
        p = groups[gid]
        cands = better_candidates(a, gid, vec, G, gids, gcat, topn=3)
        other, other_sim = (cands[0][0], round(cands[0][1], 3)) if cands else ("", None)
        other_t = groups.get(other, "") if other else ""
        adv, ev, _ = judge(a, p, other_t, other_sim, "机器自动")
        verdicts[(TN.nkey(a), gid)] = {"verdict": adv, "ev": ev,
                                       "other_gid": other if adv == "flip" else "",
                                       "own_name": own_name}
        out_b.append({"判定建议": adv, "竞品名": a, "原标注自家名": p, "更像者": other_t,
                      "分差": other_sim, "来源": "机器自动", "证据": ev})

    for name, rowsx in (("疑似错标466复核", out_a), ("机器自动SPU165复核", out_b)):
        df = pd.DataFrame(rowsx)
        print(f"  {name}: {df['判定建议'].value_counts().to_dict()}")

    with pd.ExcelWriter(os.path.join(OUT, "复核_疑似错标_20260909.xlsx"), engine="openpyxl") as w:
        pd.DataFrame(out_a).to_excel(w, sheet_name="疑似错标466", index=False)
        pd.DataFrame(out_b).to_excel(w, sheet_name="机器自动SPU165", index=False)
    with open(os.path.join(V4, "review_verdicts.jsonl"), "w", encoding="utf-8") as f:
        for (ak, gid), v in sorted(verdicts.items()):
            f.write(json.dumps({"a_key": ak, "gid": gid, **v}, ensure_ascii=False) + "\n")
    print(f"复核完成 → xlsx(2 sheets) + review_verdicts.jsonl({len(verdicts)})")


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""L4 LLM 辅助标注（v4 扩量第 2 步）：竞品名 × top-3 自家候选 → 是否同款。

省钱设计：按 anchor 分组，一次调用判最多 3 个候选（2361 次调用 vs 6000 次）。
断点续跑：结果缓存 output/model_train/v4/llm_verdicts.jsonl（追加式，重启跳过已判）。
成本可见：走 ai_client.call_llm（决策日志 output/ai_call_log.jsonl）。

跑法：python scripts/llm_label.py [--limit N] [--dry]
"""
import io
import json
import os
import sys
import time

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUT = os.path.join(BASE, "output")
V4 = os.path.join(OUT, "model_train", "v4")
CACHE = os.path.join(V4, "llm_verdicts.jsonl")

import text_norm as TN

SYS = ("你是五金建材/日百商品对齐专家。判断竞品商品名与候选自家商品名是否为"
       "**同款商品**：同品牌且同型号且同规格，功能用途一致。"
       "仅包装数量/销售挂法/标题写法不同 → 同款；规格不同（口径/伏数/脚数/尺寸/斤两）→ 不同款；"
       "同类目但不同品 → 不同款；品牌不同且型号无法对应 → 不确定偏不同款。")


def load_cache():
    done = {}
    if os.path.exists(CACHE):
        for x in open(CACHE, encoding="utf-8"):
            x = x.strip()
            if not x:
                continue
            r = json.loads(x)
            done[r["key"]] = r
    return done


def group_queue():
    q = [json.loads(x) for x in open(os.path.join(V4, "llm_queue.jsonl"), encoding="utf-8") if x.strip()]
    by_a = {}
    for r in q:
        by_a.setdefault(TN.nkey(r["a"]), {"a": r["a"], "store": r.get("store", ""), "cands": []})
        by_a[TN.nkey(r["a"])]["cands"].append(r)
    for v in by_a.values():
        v["cands"] = sorted(v["cands"], key=lambda x: -x["sim"])[:3]
    return sorted(by_a.values(), key=lambda x: x["a"])


def judge_anchor(item, ai):
    import ai_client
    lines = [f"竞品商品名：{item['a']}"]
    for i, c in enumerate(item["cands"]):
        lines.append(f"候选{i}：{c['p']}（自家类目同名品组，字面相似度{c['sim']}）")
    msgs = [{"role": "system", "content": SYS},
            {"role": "user", "content": "\n".join(lines) +
             '\n\n输出 JSON：{"verdicts":[{"idx":0,"same":true,"reason":"10字内"}]} '
             "same 取 true/false/null(不确定)，每个候选一条，按 idx 顺序。"}]
    txt = ai_client.call_llm("v4_label", msgs, json_mode=True, temperature=0.1,
                             max_tokens=200, config=ai)
    data = json.loads(txt)
    out = {}
    for v in data.get("verdicts", []):
        i = int(v.get("idx", -1))
        if 0 <= i < len(item["cands"]):
            s = v.get("same")
            out[item["cands"][i]["gid"]] = {"same": None if s is None else bool(s),
                                            "reason": str(v.get("reason", ""))[:40]}
    return out


def main():
    limit = None
    dry = "--dry" in sys.argv
    if "--limit" in sys.argv:
        limit = int(sys.argv[sys.argv.index("--limit") + 1])
    import ai_client
    ai = ai_client.load_config()
    if (ai.get("driver") or "real") == "real" and not dry:
        print(f"真实 LLM：{ai.get('model')}（费用进 ai_call_log.jsonl）")
    items = group_queue()
    done = load_cache()
    pending = [x for x in items if x["a"] not in {k[0] for k in done}]
    print(f"anchor 总 {len(items)}，已判 {len(items) - len(pending)}，待判 {len(pending)}"
          + (f"，本次 limit={limit}" if limit else ""))
    if dry:
        for x in pending[:3]:
            print(json.dumps(x, ensure_ascii=False)[:300])
        return
    fh = open(CACHE, "a", encoding="utf-8")
    n_ok = n_err = 0
    t0 = time.time()
    for n, item in enumerate(pending):
        if limit and n >= limit:
            break
        try:
            res = judge_anchor(item, ai)
            ak = TN.nkey(item["a"])
            for c in item["cands"]:
                v = res.get(c["gid"], {"same": None, "reason": "无返回"})
                fh.write(json.dumps({"key": [ak, c["gid"]], "a": item["a"],
                                     "gid": c["gid"], "p": c["p"],
                                     "same": v["same"], "reason": v["reason"],
                                     "sim": c["sim"], "store": item.get("store", "")},
                                    ensure_ascii=False) + "\n")
            fh.flush()
            n_ok += 1
        except Exception as e:
            n_err += 1
            print(f"  [ERR] {item['a'][:24]}: {e}")
        if (n + 1) % 50 == 0:
            el = time.time() - t0
            print(f"  {n + 1}/{len(pending)} 用时{el:.0f}s ETA {(el / (n + 1) * (len(pending) - n - 1)) / 60:.1f}min")
    fh.close()
    print(f"完成：成功 {n_ok}，失败 {n_err} → {CACHE}")


if __name__ == "__main__":
    main()

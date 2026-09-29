# -*- coding: utf-8 -*-
"""把 agent_verdicts/batch_*.json 物化为 llm_verdicts.jsonl（与 llm_label.py 同格式）。"""
import collections
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import text_norm as TN

V4 = os.path.join("output", "model_train", "v4")
QUEUE = os.path.join(V4, "queue_compact.jsonl")
OUT = os.path.join(V4, "llm_verdicts.jsonl")


def main():
    q = {}
    with open(QUEUE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                r = json.loads(line)
                q[r["i"]] = r

    stat = collections.Counter()
    rows = []
    for fp in sorted(glob.glob(os.path.join(V4, "agent_verdicts", "batch_*.json"))):
        with open(fp, encoding="utf-8") as f:
            doc = json.load(f)
        for b in doc.get("b", []):
            idx = b["i"]
            item = q.get(idx)
            if item is None:
                print(f"[WARN] {os.path.basename(fp)} i={idx} 不在队列")
                stat["orphan"] += 1
                continue
            cs = item["cs"]
            vs = b.get("v", [])
            rs = b.get("r", [])
            if len(vs) != len(cs):
                # v 数组与候选数不符（旧批次与队列错位）→ 整锚点弃用为 null，不注入错位标签
                print(f"[NULL] i={idx} v={len(vs)} vs cands={len(cs)}")
                for c in cs:
                    rows.append({"key": [TN.nkey(item["a"]), c["gid"]],
                                 "a": item["a"], "gid": c["gid"], "p": c["p"],
                                 "same": None, "reason": "batch错位弃用",
                                 "sim": c["sim"], "store": "",
                                 "labeler": "agent-LLM", "i": idx})
                    stat["null"] += 1
                continue
            for k, c in enumerate(cs):
                v = vs[k]
                if v == 1:
                    same, stat["pos"] = True, stat["pos"] + 1
                elif v == 0:
                    same, stat["neg"] = False, stat["neg"] + 1
                else:
                    same, stat["null"] = None, stat["null"] + 1
                reason = rs[k] if k < len(rs) else ""
                rows.append({"key": [TN.nkey(item["a"]), c["gid"]],
                             "a": item["a"], "gid": c["gid"], "p": c["p"],
                             "same": same, "reason": reason, "sim": c["sim"],
                             "store": "", "labeler": "agent-LLM", "i": idx})
    with open(OUT, "w", encoding="utf-8") as f:
        for r in rows:
            f.write(json.dumps(r, ensure_ascii=False) + "\n")
    print(f"写 {OUT}: {len(rows)} 行  pos={stat['pos']} neg={stat['neg']} null={stat['null']} orphan={stat['orphan']}")


if __name__ == "__main__":
    main()

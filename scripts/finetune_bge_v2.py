# -*- coding: utf-8 -*-
"""小模型微调 v2：品级方向 MNR 轻校准（吸取 v1 类目词方向塌缩的教训）。

配方变化（vs v1）：
  - 数据：output/model_train/v2/mnr_v2.jsonl（竞品全名↔自家全名，无类目短词）
  - 损失：只用 MNR（去掉三元组，避免过度挤压塌空间）
  - 步幅：LR 1e-5、2 轮、batch 16（192 对，轻校准不重铸空间）
  - 验证：黄金集留出组 recall@1/5/10（考什么显示什么），难负分离仅参考
  - 保存：仅当验证 recall@5 超过基线才存 bge-small-zh-v1.5-ft2
红线：不放 FINETUNED 标记、不动线上 model.onnx。
跑法：.ftvenv\\Scripts\\python.exe scripts\\finetune_bge_v2.py
"""
import json
import os
import sys
import time
import threading
import traceback
import random
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V2 = os.path.join(BASE, "output", "model_train", "v2")
MNR_PATH = os.path.join(V2, "mnr_v2.jsonl")
VAL_PATH = os.path.join(V2, "val_pairs_v2.jsonl")
GROUPS_PATH = os.path.join(V2, "groups_v2.jsonl")
HARD_PATH = os.path.join(V2, "hard_neg_v2.jsonl")
TRAIN_DIR = os.path.join(BASE, "output", "model_train")
STATUS_PATH = os.path.join(TRAIN_DIR, "train_status.json")
LOG_PATH = os.path.join(TRAIN_DIR, "train_log.txt")
SRC_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-hf")
DST_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft2")

EPOCHS = 2
BATCH = 16
LR = 1e-5
MAXLEN = 48

_status_lock = threading.Lock()
_start_ts = time.time()


def now():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def log(msg):
    line = f"[{now()}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def write_status(**kw):
    add_eval = kw.pop("add_eval", None)
    with _status_lock:
        try:
            with open(STATUS_PATH, encoding="utf-8") as f:
                st = json.load(f)
        except Exception:
            st = {}
        st["updated_at"] = now()
        st["updated_ts"] = time.time()
        st.update(kw)
        if add_eval:
            hist = [h for h in (st.get("eval_history") or []) if h.get("epoch") != add_eval.get("epoch")]
            hist.append(add_eval)
            st["eval_history"] = hist
        tmp = f"{STATUS_PATH}.{os.getpid()}.tmp"
        for attempt in range(5):
            try:
                with open(tmp, "w", encoding="utf-8") as f:
                    json.dump(st, f, ensure_ascii=False)
                os.replace(tmp, STATUS_PATH)
                break
            except PermissionError:
                if attempt == 4:
                    raise
                time.sleep(0.2 * (attempt + 1))


def jl(path):
    with open(path, encoding="utf-8") as f:
        return [json.loads(x) for x in f if x.strip()]


def norm(emb):
    import numpy as np
    return emb / (np.linalg.norm(emb, axis=1, keepdims=True) + 1e-12)


def evaluate(model, val, groups, hard):
    import numpy as np
    gv = norm(model.encode([g["t"] for g in groups], batch_size=64,
                           show_progress_bar=False, convert_to_numpy=True))
    qv = norm(model.encode([v["q"] for v in val], batch_size=64,
                           show_progress_bar=False, convert_to_numpy=True))
    gidx = {g["gid"]: j for j, g in enumerate(groups)}
    r = {"1": 0, "5": 0, "10": 0}
    cos_true = []
    for k, v in enumerate(val):
        sims = gv @ qv[k]
        j0 = gidx[v["gid"]]
        rank = int(np.where(np.argsort(sims)[::-1] == j0)[0][0])
        cos_true.append(float(sims[j0]))
        for kk in r:
            if rank < int(kk):
                r[kk] += 1
    n = len(val)
    pos_med = float(np.median(cos_true))
    neg_med, gate = None, None
    if hard:
        av = norm(model.encode([h["a"] for h in hard], batch_size=64,
                               show_progress_bar=False, convert_to_numpy=True))
        bv = norm(model.encode([h["b"] for h in hard], batch_size=64,
                               show_progress_bar=False, convert_to_numpy=True))
        negs = (av * bv).sum(axis=1)
        neg_med = float(np.median(negs))
        gate = float((negs > 0.62).mean())
    return {"recall1": round(r["1"] / n, 4), "recall5": round(r["5"] / n, 4),
            "recall10": round(r["10"] / n, 4), "pos_med": round(pos_med, 4),
            "neg_med": None if neg_med is None else round(neg_med, 4),
            "gate_ratio": None if gate is None else round(gate, 4),
            "n_val": n, "n_hard": len(hard)}


def main():
    global _start_ts
    _start_ts = time.time()
    os.makedirs(TRAIN_DIR, exist_ok=True)
    write_status(status="running", phase="初始化(v2)", epoch=0, total_epochs=EPOCHS,
                 step=0, total_steps=0, progress=0.0, loss=None,
                 loss_history=[], eval_history=[], err="",
                 started_at=now(), finished_at=None,
                 src_model="scripts/models/bge-small-zh-v1.5-hf",
                 dst_model="scripts/models/bge-small-zh-v1.5-ft2",
                 pid=os.getpid(), version="v2", saved=False,
                 recipe={"loss": "MNR", "lr": LR, "epochs": EPOCHS,
                         "batch": BATCH, "data": "品级对（无类目短词）"})
    log(f"v2 微调启动 pid={os.getpid()} lr={LR} epochs={EPOCHS} batch={BATCH}")

    import torch
    from torch.utils.data import DataLoader
    from sentence_transformers import (SentenceTransformer, models, losses,
                                       InputExample, SentencesDataset)
    from transformers import get_linear_schedule_with_warmup

    train = [(d["a"], d["p"]) for d in jl(MNR_PATH)]
    val = jl(VAL_PATH)
    groups = jl(GROUPS_PATH)
    hard = jl(HARD_PATH)
    log(f"训练对 {len(train)}，验证对 {len(val)}，品组库 {len(groups)}，难负 {len(hard)}")

    write_status(phase="加载底座模型")
    word = models.Transformer(SRC_MODEL, max_seq_length=MAXLEN)
    pool = models.Pooling(word.get_word_embedding_dimension(),
                          pooling_mode_cls_token=True)
    model = SentenceTransformer(modules=[word, pool])

    write_status(phase="基线评估")
    baseline = evaluate(model, val, groups, hard)
    log(f"基线：recall@1={baseline['recall1']} @5={baseline['recall5']} "
        f"@10={baseline['recall10']} 真对中位={baseline['pos_med']} 难负中位={baseline['neg_med']}")
    write_status(baseline=baseline, add_eval={"epoch": 0, **baseline})

    ds = SentencesDataset([InputExample(texts=list(t)) for t in train], model)
    dl = DataLoader(ds, shuffle=True, batch_size=BATCH,
                    collate_fn=model.smart_batching_collate)
    loss_fn = losses.MultipleNegativesRankingLoss(model)
    opt = torch.optim.AdamW(loss_fn.parameters(), lr=LR, weight_decay=0.01)
    total_steps = len(dl) * EPOCHS
    warmup = max(4, int(total_steps * 0.1))
    sched = get_linear_schedule_with_warmup(opt, warmup, total_steps)
    log(f"总步数 {total_steps}，warmup {warmup}")
    write_status(phase="训练中", steps_per_epoch=len(dl), total_steps=total_steps)

    step_no = 0
    best_r5 = baseline["recall5"]
    for ep in range(1, EPOCHS + 1):
        s, c = 0.0, 0
        for features, labels in dl:
            loss = loss_fn(features, labels)
            opt.zero_grad()
            loss.backward()
            torch.nn.utils.clip_grad_norm_(loss_fn.parameters(), 1.0)
            opt.step()
            sched.step()
            step_no += 1
            lv = float(loss.item())
            s += lv
            c += 1
            rec = {"step": step_no, "epoch": ep, "loss": round(lv, 4)}
            frac = step_no / total_steps
            eta = int((time.time() - _start_ts) / frac * (1 - frac)) if frac > 0.01 else None
            write_status(step=step_no, epoch=ep, loss=rec["loss"],
                         progress=round(frac, 4), add_loss=[rec],
                         eta_sec=eta, phase="训练中")
        avg = round(s / max(c, 1), 4)
        log(f"epoch {ep} 完成 平均 loss={avg:.4f}，黄金集评估中...")
        write_status(phase="评估中", epoch=ep, avg_epoch_loss=avg)
        m = evaluate(model, val, groups, hard)
        log(f"epoch {ep} 验证：recall@1={m['recall1']} @5={m['recall5']} "
            f"@10={m['recall10']} 真对中位={m['pos_med']} 难负中位={m['neg_med']}")
        write_status(add_eval={"epoch": ep, **m})
        if m["recall5"] > best_r5:
            best_r5 = m["recall5"]
            model.save(DST_MODEL)
            log(f"验证 recall@5 提升，保存 → scripts/models/bge-small-zh-v1.5-ft2")
            write_status(saved=True, best_recall5=best_r5)
        write_status(phase="训练中" if ep < EPOCHS else "完成",
                     epoch=ep, progress=round(ep / EPOCHS, 4))

    if os.path.exists(os.path.join(DST_MODEL, "config.json")):
        log(f"完成。最优验证 recall@5={best_r5}（基线 {baseline['recall5']}）。"
            "下一步：导出 ONNX 实测全量黄金集后再议放标记。")
    else:
        log(f"完成但验证无提升（基线 @5={baseline['recall5']}），未保存模型，线上维持原版。")
    write_status(status="done", phase="完成", progress=1.0, step=total_steps,
                 finished_at=now(), best_recall5=best_r5,
                 baseline_recall5=baseline["recall5"],
                 onnx_ready=os.path.exists(os.path.join(DST_MODEL, "config.json")))


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        log("训练异常退出：\n" + err)
        write_status(status="error", phase="异常", err=err[-2000:],
                     finished_at=now())
        sys.exit(1)

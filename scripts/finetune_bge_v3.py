# -*- coding: utf-8 -*-
"""小模型微调 v3（正式）：v3 品级数据（L1 硬证据正对 + L5 难负进损失）。

配方 = 已验证探针 _train_v3_probe.py：MNR(534 正对) + TripletLoss(1389 显式难负,
margin 0.2)，bge-small-zh-v1.5-hf，batch 16，lr 2e-5，3 轮，固定种子 42。

验收三门槛（缺一不可，单一指标不采信——防 v1/退化解假阳性）：
  recall@5 ≥ 0.85 且 无匹配误报(难负超0.62) ≤ 0.10 且 AUC > 0.8
产物：
  - 通过保优保存 → scripts/models/bge-small-zh-v1.5-ft3/（含 metrics.json）
  - metrics.json：数据指纹 + 模型 hash + 逐轮指标 + 验收判定（FINETUNED 标记
    只在人工确认且验收通过后由脚本写入，见 mark_finetuned.py）
红线：不动线上 model.onnx，不放 FINETUNED 标记。
跑法：.ftvenv\\Scripts\\python.exe scripts\\finetune_bge_v3.py
"""
import hashlib
import json
import os
import re
import sys
import time
import threading
import traceback
import random
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
V3 = os.environ.get("TRAIN_DATA_DIR") or os.path.join(BASE, "output", "model_train", "v3")
TRAIN_DIR = os.path.join(BASE, "output", "model_train")
STATUS_PATH = os.path.join(TRAIN_DIR, "train_status.json")
LOG_PATH = os.path.join(TRAIN_DIR, "train_log.txt")
SRC_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-hf")
DST_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft3")

EPOCHS, BATCH, LR, MAXLEN, MARGIN, SEED = int(os.environ.get("FT_EPOCHS", "3")), 16, 2e-5, 48, 0.2, 42
GATE_R5, GATE_FP, GATE_AUC = 0.85, 0.10, 0.80

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
    add_loss = kw.pop("add_loss", None)
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
        if add_loss:
            st["loss_history"] = (st.get("loss_history") or [])[-2000:] + list(add_loss)
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


def jl(fn):
    p = os.path.join(V3, fn)
    if not os.path.exists(p):
        alt = re.sub(r"_v3\.jsonl$", "_v4.jsonl", fn)
        p2 = os.path.join(V3, alt)
        if os.path.exists(p2):
            p = p2
    with open(p, encoding="utf-8") as f:
        return [json.loads(x) for x in f if x.strip()]


def accepted(m):
    return (m["recall5"] >= GATE_R5 and m["neg_gt_062"] <= GATE_FP
            and m["auc"] > GATE_AUC)


def evaluate(model, gnames, ev, gidx, hn, pos):
    import numpy as np
    enc = lambda ts: model.encode(ts, batch_size=64, show_progress_bar=False,
                                  convert_to_numpy=True, normalize_embeddings=True)
    gv, qv = enc(gnames), enc([(e.get("q") or e["a"]) for e in ev])
    ranks, cos_t = [], []
    for k, e in enumerate(ev):
        s = gv @ qv[k]
        j0 = gidx[e["gid"]]
        ranks.append(int(np.where(np.argsort(s)[::-1] == j0)[0][0]))
        cos_t.append(float(s[j0]))
    ncos = (enc([h["a"] for h in hn]) * enc([h["b"] for h in hn])).sum(axis=1)
    pcos = (enc([p["a"] for p in pos]) * enc([p["p"] for p in pos])).sum(axis=1)
    posv = np.r_[cos_t, pcos]
    sc = np.r_[posv, ncos]
    lab = np.r_[np.ones(len(posv)), np.zeros(len(ncos))]
    o = np.argsort(sc)
    rk = np.empty_like(o)
    rk[o] = np.arange(len(sc))
    auc = float((rk[lab == 1].sum() - len(posv) * (len(posv) - 1) / 2)
                / (len(posv) * len(ncos)))
    n = len(ev)
    return {"recall1": round(sum(1 for r in ranks if r < 1) / n, 4),
            "recall5": round(sum(1 for r in ranks if r < 5) / n, 4),
            "recall10": round(sum(1 for r in ranks if r < 10) / n, 4),
            "pos_med": round(float(np.median(posv)), 4),
            "neg_med": round(float(np.median(ncos)), 4),
            "neg_gt_062": round(float((ncos > 0.62).mean()), 4),
            "auc": round(auc, 4), "n_eval": n, "n_hard": len(ncos),
            "accepted": accepted({"recall5": sum(1 for r in ranks if r < 5) / n,
                                  "neg_gt_062": float((ncos > 0.62).mean()),
                                  "auc": auc})}


def model_sha256():
    h = hashlib.sha256()
    for name in ("model.safetensors", "pytorch_model.bin"):
        p = os.path.join(DST_MODEL, name)
        if os.path.exists(p):
            with open(p, "rb") as f:
                for chunk in iter(lambda: f.read(1 << 20), b""):
                    h.update(chunk)
            return h.hexdigest()
    return ""


def main():
    global _start_ts
    _start_ts = time.time()
    os.makedirs(TRAIN_DIR, exist_ok=True)
    write_status(status="running", phase="初始化(v3)", version="v3", saved=False,
                 epoch=0, total_epochs=EPOCHS, step=0, total_steps=0,
                 progress=0.0, loss=None, loss_history=[], eval_history=[],
                 err="", started_at=now(), finished_at=None,
                 src_model="scripts/models/bge-small-zh-v1.5-hf",
                 dst_model="scripts/models/bge-small-zh-v1.5-ft3",
                 pid=os.getpid(),
                 recipe={"loss": "MNR+Triplet(0.2)", "lr": LR, "epochs": EPOCHS,
                         "batch": BATCH, "seed": SEED,
                         "gates": f"recall@5>={GATE_R5} & 误报<={GATE_FP} & AUC>{GATE_AUC}"})
    log(f"v3 正式微调启动 pid={os.getpid()} lr={LR} epochs={EPOCHS} seed={SEED}")

    import torch
    from torch.utils.data import DataLoader
    from sentence_transformers import (SentenceTransformer, models, losses,
                                       InputExample, SentencesDataset)
    from transformers import get_linear_schedule_with_warmup
    from itertools import zip_longest

    random.seed(SEED)
    torch.manual_seed(SEED)

    groups, pos = jl("groups.jsonl"), jl("pos_pairs.jsonl")
    ev, tri, hn = jl("eval_pairs.jsonl"), jl("triplet_v3.jsonl"), jl("hard_neg_v3.jsonl")
    gnames = [g["t"] for g in groups]
    gidx = {g["gid"]: i for i, g in enumerate(groups)}
    log(f"正对 {len(pos)}，三元组 {len(tri)}，评估 {len(ev)}，难负 {len(hn)}，品组 {len(groups)}")
    write_status(n_pos=len(pos), n_tri=len(tri), n_eval=len(ev), n_hard=len(hn))

    write_status(phase="加载底座模型")
    word = models.Transformer(SRC_MODEL, max_seq_length=MAXLEN)
    pool = models.Pooling(word.get_word_embedding_dimension(),
                          pooling_mode_cls_token=True)
    model = SentenceTransformer(modules=[word, pool])

    write_status(phase="基线评估")
    base = evaluate(model, gnames, ev, gidx, hn, pos)
    log(f"基线 recall@1={base['recall1']} @5={base['recall5']} @10={base['recall10']} "
        f"误报={base['neg_gt_062']} AUC={base['auc']}")
    write_status(baseline=base, add_eval={"epoch": 0, **base})

    ds = SentencesDataset([InputExample(texts=[p["a"], p["p"]]) for p in pos], model)
    dl = DataLoader(ds, shuffle=True, batch_size=BATCH,
                    collate_fn=model.smart_batching_collate)
    dst = SentencesDataset(
        [InputExample(texts=[t["anchor"], t["positive"], t["negative"]]) for t in tri],
        model)
    dl_t = DataLoader(dst, shuffle=True, batch_size=BATCH,
                      collate_fn=model.smart_batching_collate)
    loss_mnr = losses.MultipleNegativesRankingLoss(model)
    loss_tri = losses.TripletLoss(model,
                                  distance_metric=losses.TripletDistanceMetric.COSINE,
                                  triplet_margin=MARGIN)
    opt = torch.optim.AdamW(list(loss_mnr.parameters()) + list(loss_tri.parameters()),
                            lr=LR, weight_decay=0.01)
    steps_per_epoch = max(len(dl), len(dl_t))
    total = steps_per_epoch * EPOCHS
    sched = get_linear_schedule_with_warmup(opt, max(10, int(total * 0.1)), total)
    log(f"每轮 {steps_per_epoch} 步（MNR {len(dl)} + 三元组 {len(dl_t)}），总 {total} 步")
    write_status(phase="训练中", steps_per_epoch=steps_per_epoch, total_steps=total)

    step = 0
    best = dict(base)
    best_ep = 0
    for ep in range(1, EPOCHS + 1):
        s, c = 0.0, 0
        for pair in zip_longest(dl, dl_t):
            lm = loss_mnr(pair[0][0], pair[0][1]) if pair[0] is not None else None
            lt = loss_tri(pair[1][0], pair[1][1]) if pair[1] is not None else None
            loss = sum(x for x in (lm, lt) if x is not None)
            opt.zero_grad()
            loss.backward()
            torch.nn.utils.clip_grad_norm_(
                list(loss_mnr.parameters()) + list(loss_tri.parameters()), 1.0)
            opt.step()
            sched.step()
            step += 1
            lv = float(loss.item())
            s += lv
            c += 1
            frac = step / total
            eta = int((time.time() - _start_ts) / frac * (1 - frac)) if frac > 0.01 else None
            write_status(step=step, epoch=ep, loss=round(lv, 4),
                         progress=round(frac, 4),
                         add_loss=[{"step": step, "epoch": ep, "loss": round(lv, 4)}],
                         eta_sec=eta, phase="训练中")
            if step % 20 == 0:
                log(f"  epoch {ep} step {step}/{total} loss={lv:.4f}")
        avg = round(s / max(c, 1), 4)
        log(f"epoch {ep} 完成 平均 loss={avg:.4f}，考试中...")
        write_status(phase="评估中", epoch=ep, avg_epoch_loss=avg)
        m = evaluate(model, gnames, ev, gidx, hn, pos)
        log(f"epoch {ep} 验收：recall@5={m['recall5']} 误报={m['neg_gt_062']} "
            f"AUC={m['auc']} → {'✅达标' if m['accepted'] else '未达标'}")
        write_status(add_eval={"epoch": ep, "avg_loss": avg, **m})
        better = (m["recall5"] > best["recall5"]
                  or (m["recall5"] == best["recall5"] and m["accepted"] and not best.get("accepted")))
        if better:
            best, best_ep = dict(m), ep
            model.save(DST_MODEL)
            log(f"保优保存 → scripts/models/bge-small-zh-v1.5-ft3")
            write_status(saved=True)
        write_status(phase="训练中" if ep < EPOCHS else "完成",
                     epoch=ep, progress=round(ep / EPOCHS, 4))

    acc = bool(best.get("accepted"))
    if best_ep > 0:
        meta_src = os.path.join(V3, "meta.json")
        metrics = {"version": "v3", "finished_at": now(),
                   "recipe": {"loss": "MNR+Triplet", "margin": MARGIN, "lr": LR,
                              "epochs": EPOCHS, "batch": BATCH, "seed": SEED,
                              "maxlen": MAXLEN},
                   "data": {"dir": os.path.relpath(V3, BASE).replace("\\", "/"),
                            "n_pos": len(pos), "n_tri": len(tri),
                            "n_eval": len(ev), "n_hard": len(hn),
                            "meta": json.load(open(meta_src, encoding="utf-8"))
                            if os.path.exists(meta_src) else None},
                   "best_epoch": best_ep,
                   "best_metrics": best,
                   "baseline": base,
                   "gates": {"recall5": GATE_R5, "false_pos": GATE_FP, "auc": GATE_AUC},
                   "accepted": acc,
                   "model_sha256": model_sha256()}
        with open(os.path.join(DST_MODEL, "metrics.json"), "w", encoding="utf-8") as f:
            json.dump(metrics, f, ensure_ascii=False, indent=1)
        log(f"metrics.json 已写入。最优 epoch{best_ep}: recall@5={best['recall5']} "
            f"误报={best['neg_gt_062']} AUC={best['auc']} 验收={'通过' if acc else '未通过'}")
    else:
        log("所有轮次均未超过基线，未保存模型，线上维持原版。")
    write_status(status="done", phase="完成", progress=1.0, step=total,
                 finished_at=now(), saved=best_ep > 0,
                 best_recall5=best["recall5"], baseline_recall5=base["recall5"],
                 accepted=acc, onnx_ready=False)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        log("训练异常退出：\n" + err)
        write_status(status="error", phase="异常", err=err[-2000:],
                     finished_at=now())
        sys.exit(1)

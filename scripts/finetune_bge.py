# -*- coding: utf-8 -*-
"""
小模型微调（P2.4 微调 v1）：对比学习微调 BGE-small-zh-v1.5。
用 .ftvenv 环境运行：.ftvenv\\Scripts\\python.exe scripts\\finetune_bge.py
数据：题词向量库/mnr_train.jsonl（anchor-positive 对，MNR 损失）。
评估：题词向量库/bi_encoder_train.jsonl（label 0/1），训练前基线 + 每轮复评。
进度：output/model_train/train_status.json（监控面板 train_panel.py 轮询）。
日志：output/model_train/train_log.txt。
产物：scripts/models/bge-small-zh-v1.5-ft/（margin 提升才保存）。
红线：只产候选模型，不自动放 FINETUNED 标记、不动 model.onnx。
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
DATA_DIR = os.path.join(BASE, "题词向量库")
MNR_PATH = os.path.join(DATA_DIR, "mnr_train.jsonl")
BI_PATH = os.path.join(DATA_DIR, "bi_encoder_train.jsonl")
TRAIN_DIR = os.path.join(BASE, "output", "model_train")
STATUS_PATH = os.path.join(TRAIN_DIR, "train_status.json")
LOG_PATH = os.path.join(TRAIN_DIR, "train_log.txt")
SRC_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-hf")
DST_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft")

EPOCHS = 3
BATCH = 16
LR = 2e-5
MAXLEN = 48
EVAL_CAP = 800
GATE_T = 0.62

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
    add_loss = kw.pop("add_loss", None)
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
        if add_loss:
            st["loss_history"] = (st.get("loss_history") or [])[-2000:] + list(add_loss)
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


def load_mnr():
    ex = []
    with open(MNR_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            ex.append((d.get("anchor", ""), d.get("positive", "")))
    return ex


def load_triplet():
    ex = []
    path = os.path.join(DATA_DIR, "triplet_train.jsonl")
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            ex.append((d.get("anchor", ""), d.get("positive", ""), d.get("negative", "")))
    return ex


def load_bi(cap):
    pos, neg = [], []
    with open(BI_PATH, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            d = json.loads(line)
            item = (d.get("sentence1", ""), d.get("sentence2", ""))
            (pos if int(d.get("label", 0)) == 1 else neg).append(item)
    random.shuffle(pos)
    random.shuffle(neg)
    return pos[:cap // 2], neg[:cap // 2]


def evaluate(model, pos, neg):
    import numpy as np
    texts = []
    for a, b in pos + neg:
        texts.extend([a, b])
    emb = model.encode(texts, batch_size=64, show_progress_bar=False,
                       convert_to_numpy=True)
    emb = emb / (np.linalg.norm(emb, axis=1, keepdims=True) + 1e-12)
    pos_c, neg_c = [], []
    row = 0
    for _ in pos:
        pos_c.append(float(np.dot(emb[row], emb[row + 1])))
        row += 2
    for _ in neg:
        neg_c.append(float(np.dot(emb[row], emb[row + 1])))
        row += 2
    p, n = float(np.mean(pos_c)), float(np.mean(neg_c))
    acc = float(np.mean([c >= GATE_T for c in pos_c] + [c < GATE_T for c in neg_c]))
    return {"pos_cos": round(p, 4), "neg_cos": round(n, 4),
            "margin": round(p - n, 4), "acc": round(acc, 4),
            "n_pos": len(pos_c), "n_neg": len(neg_c)}


def main():
    global _start_ts
    _start_ts = time.time()
    os.makedirs(TRAIN_DIR, exist_ok=True)
    write_status(status="running", phase="初始化", epoch=0, total_epochs=EPOCHS,
                 step=0, total_steps=0, progress=0.0, loss=None,
                 loss_history=[], eval_history=[], err="",
                 started_at=now(), finished_at=None,
                 src_model="scripts/models/bge-small-zh-v1.5-hf",
                 dst_model="scripts/models/bge-small-zh-v1.5-ft",
                 pid=os.getpid(), gate_t=GATE_T)
    log(f"微调启动 pid={os.getpid()} epochs={EPOCHS} batch={BATCH} lr={LR}")

    import torch
    from torch.utils.data import DataLoader
    from sentence_transformers import (SentenceTransformer, models, losses,
                                       InputExample, SentencesDataset)
    from transformers import get_linear_schedule_with_warmup

    write_status(phase="加载底座模型")
    word = models.Transformer(SRC_MODEL, max_seq_length=MAXLEN)
    pool = models.Pooling(word.get_word_embedding_dimension(),
                          pooling_mode_cls_token=True)
    model = SentenceTransformer(modules=[word, pool])
    log(f"底座加载完成 dim={word.get_word_embedding_dimension()} device=cpu")

    examples = load_mnr()
    log(f"MNR 训练对 {len(examples)} 条")
    write_status(phase="基线评估", mnr_count=len(examples))

    pos, neg = load_bi(EVAL_CAP)
    baseline = evaluate(model, pos, neg)
    log(f"基线评估：真对 cos={baseline['pos_cos']} 难负 cos={baseline['neg_cos']} "
        f"margin={baseline['margin']} acc@{GATE_T}={baseline['acc']}")
    write_status(baseline=baseline, add_eval={"epoch": 0, **baseline})

    ds = SentencesDataset([InputExample(texts=list(t)) for t in examples], model)
    dl = DataLoader(ds, shuffle=True, batch_size=BATCH,
                    collate_fn=model.smart_batching_collate)
    tri_examples = load_triplet()
    tri_fn = losses.TripletLoss(model, distance_metric=losses.TripletDistanceMetric.COSINE)
    dl_tri = DataLoader(
        SentencesDataset([InputExample(texts=list(t)) for t in tri_examples], model),
        shuffle=True, batch_size=BATCH, collate_fn=model.smart_batching_collate)
    log(f"三元组训练 {len(tri_examples)} 条")
    loss_fn = losses.MultipleNegativesRankingLoss(model)
    opt = torch.optim.AdamW(list(loss_fn.parameters()) + list(tri_fn.parameters()),
                            lr=LR, weight_decay=0.01)
    steps_per_epoch = max(len(dl), len(dl_tri))
    total_steps = steps_per_epoch * EPOCHS
    warmup = max(10, int(total_steps * 0.1))
    sched = get_linear_schedule_with_warmup(opt, warmup, total_steps)
    log(f"总步数 {total_steps}（每轮 MNR {len(dl)} + 三元组 {len(dl_tri)}），warmup {warmup}")
    write_status(phase="训练中", steps_per_epoch=steps_per_epoch,
                 total_steps=total_steps, tri_count=len(tri_examples))

    step_no = 0
    try:
        with open(STATUS_PATH, encoding="utf-8") as f:
            prev_best = json.load(f).get("best_margin") or 0.0
    except Exception:
        prev_best = 0.0
    best_margin = max(baseline["margin"], prev_best)
    write_status(best_margin=best_margin, prev_best_margin=prev_best)
    for ep in range(1, EPOCHS + 1):
        ep_loss_sum, ep_loss_n = 0.0, 0
        from itertools import zip_longest
        for i, pair in enumerate(zip_longest(dl, dl_tri)):
            loss_m = loss_fn(pair[0][0], pair[0][1]) if pair[0] is not None else None
            loss_t = tri_fn(pair[1][0], pair[1][1]) if pair[1] is not None else None
            loss = sum(l for l in (loss_m, loss_t) if l is not None)
            opt.zero_grad()
            loss.backward()
            torch.nn.utils.clip_grad_norm_(loss_fn.parameters(), 1.0)
            opt.step()
            sched.step()
            step_no += 1
            lv = float(loss.item())
            ep_loss_sum += lv
            ep_loss_n += 1
            rec = {"step": step_no, "epoch": ep, "loss": round(lv, 4)}
            eta = None
            if step_no % 2 == 0 or step_no == total_steps:
                frac = step_no / total_steps
                elapsed = time.time() - _start_ts
                if frac > 0.01:
                    eta = int(elapsed / frac * (1 - frac))
                write_status(step=step_no, epoch=ep, loss=rec["loss"],
                             progress=round(frac, 4), add_loss=[rec],
                             eta_sec=eta, phase="训练中")
            if step_no % 10 == 0:
                log(f"epoch {ep}/{EPOCHS} step {step_no}/{total_steps} loss={lv:.4f}")
        avg = round(ep_loss_sum / max(ep_loss_n, 1), 4)
        log(f"epoch {ep} 完成，平均 loss={avg:.4f}，评估中...")
        write_status(phase="评估中", epoch=ep, avg_epoch_loss=avg)
        m = evaluate(model, pos, neg)
        log(f"epoch {ep} 评估：真对={m['pos_cos']} 难负={m['neg_cos']} "
            f"margin={m['margin']} acc@{GATE_T}={m['acc']}")
        write_status(add_eval={"epoch": ep, **m})
        if m["margin"] > best_margin:
            best_margin = m["margin"]
            model.save(DST_MODEL)
            log("margin 提升，保存模型 → scripts/models/bge-small-zh-v1.5-ft")
            write_status(saved=True)
        write_status(phase="训练中" if ep < EPOCHS else "完成",
                     epoch=ep, progress=round(ep / EPOCHS, 4))

    if not os.path.exists(os.path.join(DST_MODEL, "config.json")):
        model.save(DST_MODEL)
        log("保存最终模型 → scripts/models/bge-small-zh-v1.5-ft")

    final = evaluate(model, pos, neg)
    log(f"最终评估：真对={final['pos_cos']} 难负={final['neg_cos']} "
        f"margin={final['margin']}（基线 margin={baseline['margin']}）")
    log("完成。下一步：人工确认后将 ft 模型转 ONNX 并放 FINETUNED 标记开层。")
    write_status(status="done", phase="完成", progress=1.0, step=total_steps,
                 add_eval={"epoch": EPOCHS, **final}, finished_at=now(),
                 onnx_ready=False, saved=True)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        err = traceback.format_exc()
        log("训练异常退出：\n" + err)
        write_status(status="error", phase="异常", err=err[-2000:],
                     finished_at=now())
        sys.exit(1)

# -*- coding: utf-8 -*-
"""
本地向量编码器（vec_encoder）：P2.4 小模型底座。
模型 = BGE-small-zh-v1.5（Xenova ONNX 转换版；scripts/models/bge-small-zh-v1.5/）。
优先 int8 量化版（24MB，CPU 快）；质量不足时切换 fp32（--fp32）。
红线：
  1. 只产候选进人工池（置信 0.50-0.60），绝不自动确认；
  2. 输入预处理剔除品牌与**营销词**，但**保留规格数字**——规格是同品/非同品的判别核心
     （2026-09-09 实测：删规格会让黄金集 AUC 从 0.80 掉到 0.48，等于随机）；
  3. 模型文件缺失 → 全链路静默跳过（is_ready()==False），永不阻塞主流程。
编码缓存：output/_向量缓存/（键=文本指纹，逐批增量写 npz）。

清洗实现统一在 scripts/text_norm.py（与训练数据生成共用同一份代码）。
"""
import json
import os
import re
import sys
import threading

_HERE = os.path.dirname(os.path.abspath(__file__))
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)
import text_norm as TN

BASE = os.path.dirname(_HERE)
MARKER = os.path.join(BASE, "scripts", "models", "FINETUNED")
_MODELS_ROOT = os.path.join(BASE, "scripts", "models")


def _resolve_model_dir():
    """FINETUNED 标记（验收三门槛 + --i-confirm 双闸产物）存在 → 切微调模型。
    标记内 JSON 的 ft_dir 指定模型目录；缺量化版/tokenizer 则回退基座，永不阻塞。"""
    try:
        with open(MARKER, encoding="utf-8") as f:
            ft_dir = (json.load(f) or {}).get("ft_dir") or ""
        if ft_dir:
            cand = os.path.join(_MODELS_ROOT, ft_dir)
            if os.path.exists(os.path.join(cand, "model_quantized.onnx")) and \
               os.path.exists(os.path.join(cand, "tokenizer.json")) and \
               os.path.exists(os.path.join(cand, "config.json")):
                return cand
    except Exception:
        pass
    return os.path.join(_MODELS_ROOT, "bge-small-zh-v1.5")


MODEL_DIR = _resolve_model_dir()
CACHE_DIR = os.path.join(BASE, "output", "_向量缓存")
MAX_LEN = 48

_lock = threading.Lock()
_state = {"sess": None, "tok": None, "dim": None, "q": True, "err": "", "path": None}


def is_ready():
    try:
        _ensure()
        return _state["sess"] is not None
    except Exception:
        return False


def status():
    _ensure()
    return {"ready": _state["sess"] is not None, "quantized": _state["q"],
            "dim": _state["dim"], "err": _state["err"],
            "model": model_id(), "model_dir": os.path.basename(MODEL_DIR),
            "finetuned": os.path.exists(MARKER), "clean_version": TN.VERSION}


def model_id():
    """模型身份指纹（进入向量缓存键）：换模型必须让缓存失效。"""
    _ensure()
    p = _state.get("path")
    if not p:
        return "none"
    try:
        return "%s/%s:%d-%d" % (os.path.basename(os.path.dirname(p)), os.path.basename(p),
                                os.path.getsize(p), int(os.path.getmtime(p)))
    except OSError:
        return os.path.basename(p)


def _ensure():
    with _lock:
        if _state["sess"] is not None or _state["err"]:
            return
        try:
            import onnxruntime as ort
            tok_path = os.path.join(MODEL_DIR, "tokenizer.json")
            q = os.path.join(MODEL_DIR, "model_quantized.onnx")
            f = os.path.join(MODEL_DIR, "model.onnx")
            model_path = q if os.path.exists(q) else f
            _state["q"] = bool(os.path.exists(q))
            if not (os.path.exists(model_path) and os.path.exists(os.path.join(MODEL_DIR, "tokenizer.json"))):
                _state["err"] = "模型文件缺失（scripts/models/bge-small-zh-v1.5/）"
                return
            with open(os.path.join(MODEL_DIR, "config.json"), encoding="utf-8") as fp:
                cfg = json.load(fp)
            _state["dim"] = int(cfg.get("hidden_size") or 384)
            from tokenizers import Tokenizer
            tok = Tokenizer.from_file(os.path.join(MODEL_DIR, "tokenizer.json"))
            tok.enable_padding(pad_id=0, pad_token="[PAD]", length=None)
            tok.enable_truncation(max_length=MAX_LEN)
            opts = ort.SessionOptions()
            opts.intra_op_num_threads = os.cpu_count() or 4
            _state["sess"] = ort.InferenceSession(model_path, opts)
            _state["tok"] = tok
            _state["path"] = model_path
        except Exception as e:
            _state["err"] = f"编码器加载失败：{e}"


_SPEC_RE = None          # 已废弃：规格不再删除，保留在文本中（见 text_norm.clean）
_BRAND_TAIL_RE = None


def clean_text(name, brand=""):
    """品名预处理：剔品牌与营销词，**保留规格数字**。委托 scripts/text_norm.py。"""
    return TN.clean(name, brand)


def encode(texts, batch=64):
    """texts -> (n, dim) float32，L2 归一化。"""
    _ensure()
    if _state["sess"] is None:
        raise RuntimeError(_state["err"] or "编码器未就绪")
    import numpy as np
    out = []
    for i in range(0, len(texts), batch):
        chunk = [clean_text(t) for t in texts[i:i + batch]]
        enc = _state["tok"].encode_batch(chunk)
        ids = [e.ids for e in enc]
        att = [[1 if t != 0 else 0 for t in seq] for seq in ids]
        import numpy as _np
        max_len = max(len(s) for s in ids) if ids else 1
        ids_a = _np.zeros((len(ids), max_len), dtype=_np.int64)
        att_a = _np.zeros((len(ids), max_len), dtype=_np.int64)
        for j, (s, a) in enumerate(zip(ids, att)):
            ids_a[j, :len(s)] = s
            att_a[j, :len(a)] = a[:len(s)]
        feeds = {"input_ids": ids_a, "attention_mask": att_a}
        try:
            outv = _state["sess"].run(None, feeds)[0]
        except Exception:
            feeds["token_type_ids"] = _np.zeros_like(ids_a)
            outv = _state["sess"].run(None, feeds)[0]
        cls = outv[:, 0, :]              # BGE = CLS 池化
        cls = cls / (np.linalg.norm(cls, axis=1, keepdims=True) + 1e-12)
        out.append(cls.astype(_np.float32))
    import numpy as _np
    return _np.vstack(out) if out else _np.zeros((0, _state["dim"] or 384), dtype=_np.float32)


def similarity(a, b):
    import numpy as np
    a, b = np.asarray(a), np.asarray(b)
    return float(np.dot(a, b))


if __name__ == "__main__":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    print("status:", is_ready(), status())
    if is_ready():
        import time
        texts = ["沉头钻尾螺钉 十字沉头螺丝", "自攻螺丝 十字槽沉头", "公牛插座插线板排插",
                 "全牙内六角螺丝M8*20mm", "M10内六角螺丝", "内六角螺丝M8 20mm"]
        t0 = time.time()
        v = encode(texts)
        print(f"编码 {len(texts)} 条 {v.shape}，{time.time()-t0:.2f}s")
        for i in range(len(texts)):
            for j in range(len(texts)):
                if i < j:
                    print(f"  {texts[i][:14]!r} <-> {texts[j][:14]!r} = {similarity(v[i], v[j]):.3f}")

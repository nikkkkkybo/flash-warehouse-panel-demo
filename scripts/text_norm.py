# -*- coding: utf-8 -*-
"""商品名归一化（训练 / 推理唯一事实源）。

铁律（2026-09-09 质检结论）：
  **规格数字必须保留。** 五金件的"同品 / 非同品"往往只差一个规格
  （M8 vs M10、110V vs 220V、2脚 vs 3脚、32口径 vs 40口径）。
  早期版本把规格当噪音删掉，导致底座模型在黄金集上 AUC=0.48（等于随机）。

本模块同时被三处使用，改动必须一起重跑：
  - scripts/build_match_dataset_v3.py   训练数据生成
  - scripts/vec_encoder.py              线上编码（clean_text 委托到此）
  - scripts/align.py                    向量候选层（_vec_clean 委托到此）

VERSION 进入向量缓存键：改动清洗口径 → 缓存自动失效，避免"换了模型/换了口径
但复用旧向量"的静默错误。
"""
import re
import collections

VERSION = "tn1"

# 纯营销 / 标签词（删除它们不影响品级语义）
MARKETING = re.compile(
    r"(规格可选|多型号可选|尺寸可选|颜色随机|颜色可选|包邮|正品|新品|特价|厂家直销|"
    r"品牌随机|优质|高品质|超强|强力|神器|加厚|加大|加长|多功能|高精度|工业级|专业级|"
    r"家用|通用|新款|升级款|网红|爆款|批发)")
# 括号本身去掉，但**括号内的内容保留**（很多规格写在括号里，如"（圆形平垫片M4-M12）"）
TAG_ONLY = re.compile(r"[【】（）()\[\]［］]")


def clean(name, brand=""):
    """归一化商品名：去品牌与营销词，**保留规格数字**。幂等。"""
    t = str(name or "")
    if brand:
        b = str(brand).strip()
        if b:
            t = t.replace(b, " ")
    t = TAG_ONLY.sub(" ", t)
    t = MARKETING.sub(" ", t)
    t = t.replace("／", "/")
    t = re.sub(r"\s+", " ", t).strip().lower()
    return t


# ---------------------------------------------------------------------------
# 规格签名：仅用于"标签判定"（冲突检测），不用于模型输入文本
# ---------------------------------------------------------------------------
_UNITS = [
    ("毫米", "len_mm", 1.0), ("mm", "len_mm", 1.0),
    ("厘米", "len_mm", 10.0), ("公分", "len_mm", 10.0), ("cm", "len_mm", 10.0),
    ("米", "len_mm", 1000.0), ("m", "len_mm", 1000.0),
    ("千克", "wt_g", 1000.0), ("公斤", "wt_g", 1000.0), ("kg", "wt_g", 1000.0),
    ("克", "wt_g", 1.0), ("g", "wt_g", 1.0), ("斤", "wt_g", 500.0),
    ("毫升", "vol_ml", 1.0), ("ml", "vol_ml", 1.0),
    ("升", "vol_ml", 1000.0), ("l", "vol_ml", 1000.0),
    ("伏", "volt_v", 1.0), ("v", "volt_v", 1.0),
    ("瓦", "pow_w", 1.0), ("w", "pow_w", 1.0),
    ("安", "amp_a", 1.0), ("a", "amp_a", 1.0),
    ("寸", "cun", 1.0),
]
_UNIT_ALT = "|".join(re.escape(u) for u, _, _ in sorted(_UNITS, key=lambda x: -len(x[0])))
_SPEC_RE = re.compile(rf"(\d+(?:\.\d+)?)\s*({_UNIT_ALT})", re.I)
_UNIT_LOOKUP = {u.lower(): (d, f) for u, d, f in _UNITS}


def spec_sig(text):
    """抽取 {维度: set(归一值)}，例如 {'len_mm': {80.0}, 'wt_g': {2500.0}}。"""
    sig = collections.defaultdict(set)
    for num, unit in _SPEC_RE.findall(str(text or "")):
        d, f = _UNIT_LOOKUP.get(unit.lower(), (None, None))
        if d is None:
            continue
        try:
            sig[d].add(round(float(num) * f, 2))
        except ValueError:
            pass
    return dict(sig)


def spec_conflict(a, b):
    """返回 a、b 双方都有但取值不同的规格维度（人可读）。"""
    sa, sb = spec_sig(a), spec_sig(b)
    out = []
    for d in sa:
        if d in sb and sa[d] != sb[d]:
            out.append(f"{d}:{sorted(sa[d])}≠{sorted(sb[d])}")
    return out


# ---------------------------------------------------------------------------
MODEL_RE = re.compile(r"[A-Za-z]{1,6}[-]?\d{2,}[A-Za-z0-9-]*")


def model_tokens(text):
    """抽取型号 token（必须同时含字母和数字，长度≥4）。"""
    out = set()
    for m in MODEL_RE.findall(str(text or "")):
        m = m.upper().strip("-")
        if len(m) >= 4 and any(c.isdigit() for c in m) and any(c.isalpha() for c in m):
            out.add(m)
    return out


def norm_bc(v):
    """条码归一化：只留数字/X，长度≥8 才有效。"""
    s = re.sub(r"[^\dXx]", "", str(v or ""))
    return s.upper() if len(s) >= 8 else ""


def nkey(t):
    """用于去重/比对的键：去空白与斜杠后小写。"""
    return re.sub(r"[\s/／]+", "", str(t or "")).lower()

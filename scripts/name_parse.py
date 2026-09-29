# -*- coding: utf-8 -*-
"""L2 品名结构化等价判定（v4 扩量 + 疑似错标复核共用）。

与 text_norm 的分工：
  - text_norm：清洗文本（模型输入）+ 规格签名/型号抽取的**基础件**
  - name_parse：在基础件之上做**结构化等价判定**（同款 / 非同款 / 不确定）

判定哲学（质检方案 §五）：
  - 正对判据：型号全等 且 规格无冲突 → pos（硬规则）
  - 负对判据：规格冲突（同一维度双方都有且**无交集**取值）→ neg（硬规则）
  - 其余 → None（交给 L4 LLM / 人工）

注意：text_norm.spec_conflict 用的是"集合不等"，会把"110v互转220v vs 110v"
误判为冲突；本模块用"无交集"语义，专用于标签判定。
"""
import re

import text_norm as TN


def _sig_conflict(a, b):
    """同一规格维度双方都有且取值无交集（真冲突）。返回人可读冲突描述。"""
    sa, sb = TN.spec_sig(a), TN.spec_sig(b)
    out = []
    for d in sa:
        if d in sb and not (sa[d] & sb[d]):
            out.append(f"{d}:{sorted(sa[d])}≠{sorted(sb[d])}")
    return out


def parse(text, brand=""):
    """商品名 → 结构化字段。"""
    t = TN.clean(text, brand)
    return {
        "clean": t,
        "models": TN.model_tokens(t),
        "spec": TN.spec_sig(t),
        "conflict_with": lambda o: _sig_conflict(t, TN.clean(o, brand)),
    }


def equivalent(a, b, brand_a="", brand_b=""):
    """结构化等价判定。返回 (verdict, reason)：
    verdict ∈ {"pos", "neg", None}。None = 规则覆盖不了，交给 LLM/人工。

    pos 规则：双方型号 token 非空且全等，且无规格冲突。
    neg 规则：规格真冲突（同维度无交集）；或双方型号非空、有交集但
              各自独有的型号相差悬殊且规格也冲突。
    品牌**不**参与判定（跨品牌同品是合法正对，方案要求 ≥30%）。
    """
    ca, cb = TN.clean(a, brand_a), TN.clean(b, brand_b)
    ma, mb = TN.model_tokens(ca), TN.model_tokens(cb)
    conf = _sig_conflict(ca, cb)
    if ma and mb and ma == mb and not conf:
        return "pos", f"型号全等{sorted(ma)[:3]}且规格无冲突"
    if conf:
        return "neg", "规格冲突:" + ";".join(conf[:3])
    if ma and mb:
        inter = ma & mb
        if inter and not conf:
            return "pos", f"共享型号{sorted(inter)[:2]}且规格无冲突"
    return None, ""


def verdict_label(verdict):
    return {"pos": "同款", "neg": "非同款", None: "不确定"}[verdict]

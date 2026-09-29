# -*- coding: utf-8 -*-
"""
关键词搜索工具：空格/分隔符分词，逐词子串 AND 匹配（比整串包含更精确）。

示例：商品名「（十字沉头自攻螺丝M4）高强度十字沉头自攻螺丝...自攻钉18-35mm/颗」
输入「自攻螺丝 M4 18-35mm」→ 拆词为 [自攻螺丝, M4, 18-35mm]，
每词分别 contains，商品名同时包含三者才命中 —— 用 2-3 个特征词精确收窄。
"""
import re

import pandas as pd


def split_keywords(q):
    """查询串 -> 关键词列表（空格/逗号/顿号/分号分隔，去空词）。"""
    if not q:
        return []
    return [w for w in re.split(r"[\s,，、;；]+", str(q).strip()) if w]


def kw_mask(df, q, cols):
    """多关键词 AND 子串匹配。返回 bool Series mask；无关键词/空表返回 None。"""
    words = split_keywords(q)
    if not words or len(df) == 0:
        return None
    mask = pd.Series([True] * len(df), index=df.index)
    for w in words:
        wm = pd.Series([False] * len(df), index=df.index)
        for c in cols:
            if c in df.columns:
                wm = wm | df[c].astype(str).str.contains(re.escape(w), case=False, na=False)
        mask = mask & wm
    return mask


def kw_hit(row_dict, q, keys):
    """对 dict 行做多关键词 AND 子串命中（browse 列表筛选用）。"""
    words = split_keywords(q)
    if not words:
        return True
    for w in words:
        wl = w.lower()
        if not any(wl in str(row_dict.get(k, "")).lower() for k in keys):
            return False
    return True

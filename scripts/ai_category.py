# -*- coding: utf-8 -*-
"""
类目 AI 分析：基于类目映射库 + 各店品类分布 + 价格带，生成类目层经营分析。
输出：类目映射建议清单、类目缺口解读、类目价格带要点、类目竞争度。
"""
import os
import json
import time

import pandas as pd

from ai_analysis import load_config

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
CAT_MAP_FILE = os.path.join(OUTPUT, "类目映射表.xlsx")
REPORT_FILE = os.path.join(OUTPUT, "竞争分析报告.xlsx")

SYSTEM_PROMPT = """你是闪电仓（五金建材零售）的经营顾问，专长类目（品类）分析。基于本地数据输出类目层经营分析，要求：
1. 全程大白话、可执行；
2. 分四部分：类目映射建议清单（哪些竞品类目还没映射到自家类目、建议怎么映射）、类目缺口解读（竞品有自家无的类目）、类目价格带要点、类目竞争度（哪些类目对手密集、机会在哪）；
3. 只依据提供的数据，不编造；
4. 用 Markdown 组织，小标题分隔。"""


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def build_category_summary(shops):
    """构建类目层摘要文本。"""
    parts = []

    # 1) 类目映射库
    if os.path.exists(CAT_MAP_FILE):
        try:
            df = pd.read_excel(CAT_MAP_FILE, dtype=str).fillna("")
            pending = df[df["建议映射"].astype(str).isin(["需人工映射", ""])]
            mapped = df[df["建议映射"].astype(str) != "需人工映射"]
            lines = [f"【类目映射库】共 {len(df)} 个竞品类目，已映射 {len(mapped)} 个，待映射 {len(pending)} 个。"]
            if len(pending):
                lines.append("待映射类目（前15）：" + "、".join(
                    _clean(r["三级类目"]) for _, r in pending.head(15).iterrows() if _clean(r["三级类目"])))
            parts.append("\n".join(lines))
        except Exception:
            pass

    # 2) 各店品类分布
    dist = {}
    for s in shops:
        f = os.path.join(OUTPUT, f"试跑_{s}", "竞品_归一化后.xlsx")
        if not os.path.exists(f):
            continue
        try:
            df = pd.read_excel(f, dtype=str).fillna("")
            col = "类目三级_归一" if "类目三级_归一" in df.columns else "category3"
            if col not in df.columns:
                continue
            for cat, n in df[col].value_counts().head(8).items():
                if _clean(cat):
                    dist[_clean(cat)] = dist.get(_clean(cat), 0) + int(n)
        except Exception:
            continue
    if dist:
        top = sorted(dist.items(), key=lambda kv: -kv[1])[:12]
        lines = ["【各店品类分布Top12（竞品数合计）】"]
        for cat, n in top:
            lines.append(f"- {cat}：{n}")
        parts.append("\n".join(lines))

    # 3) 价格带
    if os.path.exists(REPORT_FILE):
        try:
            band = pd.read_excel(REPORT_FILE, sheet_name="价格带", dtype=str).fillna("")
            if len(band):
                lines = ["【类目价格带】（主要类目商品数/中位价）"]
                for _, r in band.head(8).iterrows():
                    lines.append(f"- {_clean(r.get('类目'))}：{_clean(r.get('商品数'))}个商品，中位价 {_clean(r.get('价格中位数'))}")
                parts.append("\n".join(lines))
        except Exception:
            pass

    return "\n\n".join(parts) if parts else "（无可用类目数据）"


def generate_category_advice(shops, config=None, mock=False):
    """生成类目 AI 分析（Markdown）。mock=True 本地模拟。"""
    summary = build_category_summary(shops)
    if mock:
        return "# 类目 AI 分析（模拟数据，未调用真实接口）\n\n请配置真实 API key 后重新生成。\n\n" + summary[:200]

    config = config or load_config()
    from ai_client import call_llm
    return call_llm("category",
                    [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": "以下是本地类目层数据：\n\n" + summary}],
                    config=config, timeout=180,
                    temperature=0.4, max_tokens=1200)

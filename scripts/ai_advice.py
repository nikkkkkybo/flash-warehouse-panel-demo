# -*- coding: utf-8 -*-
"""
已对齐内容 AI 建议：针对"已确认配对"生成逐品经营建议（定价/差距/销量/防守/毛利）。
  build_aligned_summary(shops)  提炼配对摘要（本地计算价差/月售比/规格差）
  generate_advice(shops)        调 DeepSeek 生成每条建议 + 综述长文
建议仅记录状态（待看/已采纳/忽略），不自动执行任何改动。
"""
import json
import os
import time

import pandas as pd
from datetime import datetime

from ai_analysis import load_config
import sales_text as _ST

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
LIBRARY_FILE = os.path.join(OUTPUT, "对齐汇总库.xlsx")
OWN_FILE = os.path.join(OUTPUT, "自家_归一化后.xlsx")
ADVICE_FILE = os.path.join(OUTPUT, "已对齐建议.xlsx")

ADVICE_COLS = ["时间", "竞店名称", "竞品Sku_id", "竞品商品名", "竞品价", "竞品月售",
               "自家SKUID", "自家商品名", "自家售价", "自家采购价", "自家月销",
               "价差率", "月售比", "规格差", "建议", "状态"]

SYSTEM_PROMPT = """你是闪电仓（五金建材零售）的经营顾问。基于已人工确认的"竞品-自家"对应数据，为每对给出一条具体、可执行的经营建议。

要求：
1. 每条建议 30-60 字，结合给出的数字（价差/月售），不要空话；
2. 只输出 JSON 数组，不要输出其他文字：[{"sku": "竞品Sku_id", "advice": "建议内容"}, ...]；
3. 综合洞察另起一段（以"【综合】"开头，100-200字，给出整体改进方向）。
4. 数据里没有成本、毛利、库存、周转——需要它们才能成立的结论一律写【待核：缺什么】，不得推测；
5. 竞品月售里标了"（意向/收藏，非成交）"的是"多少人在想买"，已按 λ 折成有效销量参与比较：
   不得把它当成交笔数，也不得说它毫无意义——它是弱一档的需求信号。"""


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    try:
        return float(str(v).replace(",", "").strip())
    except (ValueError, TypeError):
        return None


def _shop_comp_lookup(shops):
    """各店竞品表 -> {店: {Sku_id: {价格,月售,规格}}}。"""
    out = {}
    for s in shops:
        f = os.path.join(OUTPUT, f"试跑_{s}", "竞品_归一化后.xlsx")
        if not os.path.exists(f):
            continue
        try:
            df = pd.read_excel(f, dtype=str).fillna("")
            lk = {}
            sku_col = "sku_id" if "sku_id" in df.columns else "Sku_id"
            for _, r in df.iterrows():
                sku = _clean(r.get(sku_col))
                if sku:
                    lk[sku] = {
                        "价格": _clean(r.get("price") or r.get("到手价") or r.get("折后价")),
                        "月售": _clean(r.get("sales") or r.get("月销量") or r.get("月售")),
                        "规格": _clean(r.get("spec") or r.get("规格")),
                    }
            out[s] = lk
        except Exception:
            continue
    return out


def build_aligned_summary(shops):
    """提炼已确认配对摘要。返回 (rows, stats)。
    rows: 每对 {竞店/竞品Sku_id/竞品商品名/竞品价/竞品月售/自家SKUID/自家商品名/自家售价/自家采购价/自家月销/价差率/月售比/规格差/重点}"""
    if not os.path.exists(LIBRARY_FILE):
        return [], {"重点": 0, "普通": 0, "未对应热销": 0}
    sku_df = pd.read_excel(LIBRARY_FILE, sheet_name="SKU匹配层", dtype=str).fillna("")
    sku_df = sku_df[sku_df["竞店名称"].isin(shops)] if shops else sku_df
    if not len(sku_df):
        return [], {"重点": 0, "普通": 0, "未对应热销": 0}

    own_df = pd.read_excel(OWN_FILE, dtype=str).fillna("") if os.path.exists(OWN_FILE) else pd.DataFrame()
    own_info = {}
    if len(own_df):
        sku_col = "SKUID" if "SKUID" in own_df.columns else None
        if sku_col:
            for _, r in own_df.iterrows():
                sku = _clean(r.get(sku_col))
                if sku:
                    own_info[sku] = {
                        "售价": _clean(r.get("美团闪购渠道售价") or r.get("售价")),
                        "采购价": _clean(r.get("采购价")),
                        "月销": _clean(r.get("月销量") or r.get("月售")),
                        "规格": _clean(r.get("规格名称") or r.get("规格")),
                        "有图": _clean(r.get("是否有图片")),
                    }

    comp_lk = _shop_comp_lookup(shops)
    rows = []
    for _, r in sku_df.iterrows():
        shop = _clean(r.get("竞店名称"))
        csku = _clean(r.get("竞品Sku_id"))
        osku = _clean(r.get("自家SKUID"))
        if not csku or not osku:
            continue
        ci = comp_lk.get(shop, {}).get(csku, {})
        oi = own_info.get(osku, {})
        cp, op = _num(ci.get("价格")), _num(oi.get("售价"))
        diff = ""
        if cp is not None and op and op > 0:
            diff = f"{(cp - op) / op * 100:+.0f}%"
        # 月售比用**有效销量 = 成交 + λ×收藏**：收藏照样参与（它是需求信号），只是按权重降级，
        # 不再像旧代码那样把「35人想买」擦掉字样当成 35 笔成交。
        cs_raw = ci.get("月售", "")
        cs = _ST.eff(cs_raw)
        os_ = _num(oi.get("月销"))
        ratio = ""
        if cs and os_ is not None and os_ > 0:
            ratio = f"{cs / os_:.1f}x"
        # 规格差（简版：文本长度差异近似，不作为精确判断）
        spec_diff = ""
        rows.append({
            "竞店名称": shop, "竞品Sku_id": csku, "竞品商品名": _clean(r.get("竞品商品名")),
            "竞品价": _clean(ci.get("价格")),
            "竞品月售": _ST.label(cs_raw) if _ST.is_intent(cs_raw) else _clean(cs_raw),
            "自家SKUID": osku, "自家商品名": _clean(r.get("自家商品名")),
            "自家售价": _clean(oi.get("售价")), "自家采购价": _clean(oi.get("采购价")),
            "自家月销": _clean(oi.get("月销")),
            "价差率": diff, "月售比": ratio, "规格差": spec_diff,
            "重点": bool((diff and abs(float(diff.rstrip("%"))) >= 20) or
                         (ratio and float(ratio.rstrip("x")) >= 2) or
                         _clean(r.get("确认来源")) == "人工确认"),
        })
    stats = {"重点": sum(1 for x in rows if x["重点"]), "普通": len(rows) - sum(1 for x in rows if x["重点"])}
    return rows, stats


def _summary_text(rows, stats):
    """把配对摘要转成发给大模型的文本（重点逐条+普通聚合）。"""
    parts = [f"已确认配对共 {len(rows)} 对（重点 {stats['重点']}，普通 {stats['普通']}）。"]
    key = [r for r in rows if r["重点"]][:30]
    if key:
        parts.append("【重点配对逐条】")
        for r in key:
            parts.append(
                f"- 竞品[{r['竞店名称']}]{r['竞品商品名'][:50]} 价{r['竞品价']}/月售{r['竞品月售']} "
                f"↔ 自家{r['自家商品名'][:40]} 售价{r['自家售价']}/采购{r['自家采购价']}/月销{r['自家月销']} "
                f"| 价差{r['价差率']} 月售比{r['月售比']}")
    others = [r for r in rows if not r["重点"]]
    if others:
        up = sum(1 for r in others if r["价差率"] and float(r["价差率"].rstrip("%")) > 10)
        down = sum(1 for r in others if r["价差率"] and float(r["价差率"].rstrip("%")) < -10)
        parts.append(f"【其余 {len(others)} 对聚合】我们比竞品贵10%以上的有 {up} 对，便宜10%以上的有 {down} 对。")
    return "\n".join(parts)


def generate_advice(shops, config=None, mock=False):
    """生成已对齐建议：返回 (rows(含建议), report_md)。"""
    rows, stats = build_aligned_summary(shops)
    if not rows:
        raise RuntimeError("该范围内暂无已确认配对（请先完成确认）")

    if mock:
        for r in rows[:5]:
            r["建议"] = "（模拟建议）请配置真实 API key 后重新生成。"
        report_md = "# 已对齐建议（模拟）\n\n请配置真实 DeepSeek key 后生成。"
    else:
        config = config or load_config()
        from ai_client import call_llm
        content = call_llm("advice",
                           [{"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": _summary_text(rows, stats)}],
                           config=config, timeout=180,
                           temperature=0.4, max_tokens=1500)
        # 解析 JSON 数组 + 【综合】段
        advice_map = {}
        try:
            s, e = content.find("["), content.rfind("]")
            arr = json.loads(content[s:e + 1])
            for a in arr:
                if isinstance(a, dict) and a.get("sku"):
                    advice_map[str(a["sku"])] = _clean(a.get("advice"))
        except Exception:
            pass
        for r_ in rows:
            r_["建议"] = advice_map.get(r_["竞品Sku_id"], "")
        report_md = content

    # 保存表格（含建议与状态列；保留历史状态）
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    old = {}
    if os.path.exists(ADVICE_FILE):
        try:
            odf = pd.read_excel(ADVICE_FILE, dtype=str).fillna("")
            for _, r_ in odf.iterrows():
                old[(r_["竞店名称"], r_["竞品Sku_id"])] = r_.get("状态", "待看")
        except Exception:
            pass
    recs = []
    for r_ in rows:
        recs.append({**r_, "时间": now,
                     "状态": old.get((r_["竞店名称"], r_["竞品Sku_id"]), "待看")})
    df = pd.DataFrame(recs, columns=ADVICE_COLS)
    df.to_excel(ADVICE_FILE, index=False)
    return rows, report_md, ADVICE_FILE

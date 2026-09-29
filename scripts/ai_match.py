# -*- coding: utf-8 -*-
"""
大模型辅助判断：判断"竞品"与"候选自家品"是否为同一商品（SPU 品级）。
用于粗对齐队列的 单条/批量 辅助按钮（可选使用，点才调用）。
判断结果仅为参考，一律人工确认后才写入对齐结果。
"""
import json
import os

from ai_client import call_llm, load_config

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
JUDGE_CACHE_FILE = os.path.join(BASE, "output", "_ai_judge_cache.json")
LEARN_SAMPLES_FILE = os.path.join(BASE, "output", "学习样本库.xlsx")

SYSTEM_PROMPT = """你是闪电仓（五金建材零售）的商品比对专家。你的任务是判断"竞品商品"与"候选自家商品"是否为同一个品（SPU 级）。

判断口径：
1. 类目、用途、核心参数（功率/容量/尺寸/电压/材质等）一致，即视为同一个品；
2. 规格/颜色/档位差异（如 5W vs 5.5W、白色 vs 金色）不算不同品，仍归同一品；
3. 品牌不同（且都明确）则为不同品；
4. 信息不足时给"拿不准"，不要硬猜。

只输出 JSON（不要输出其他文字）：
{"verdict": "是" 或 "否" 或 "拿不准", "confidence": 0到1的小数, "reason": "一句话理由（不超过30字）", "suggest_own": "建议对应的自家商品名（verdict为"是"或"拿不准"时给出最可能的，否则填空字符串）"}"""


def _parse_verdict(text):
    """解析模型输出为 dict；失败返回 None。"""
    if not text:
        return None
    t = str(text).strip()
    try:
        return json.loads(t)
    except (ValueError, TypeError):
        pass
    # 兜底：提取第一个 { ... }
    try:
        s, e = t.find("{"), t.rfind("}")
        if 0 <= s < e:
            return json.loads(t[s:e + 1])
    except (ValueError, TypeError):
        pass
    return None


def load_examples(limit=20):
    """读取学习样本库（已确认配对），作为 few-shot 示例。文件不存在返回 []。"""
    if not os.path.exists(LEARN_SAMPLES_FILE):
        return []
    try:
        import pandas as pd
        df = pd.read_excel(LEARN_SAMPLES_FILE, dtype=str).fillna("")
        rows = []
        for _, r in df.tail(limit).iterrows():
            comp = str(r.get("竞品商品名", "")).strip()
            own = str(r.get("自家商品名", "")).strip()
            if comp and own:
                rows.append({"comp": comp, "own": own, "verdict": "是"})
        return rows
    except Exception:
        return []


def judge(comp_name, comp_brand="", comp_model="", candidates=None,
          examples=None, config=None, timeout=120, module="match_judge"):
    """判断竞品与候选自家品是否同一品。

    comp_name  : 竞品商品名
    candidates : 候选自家商品名列表（规则粗筛 top3）
    examples   : few-shot 已确认配对 [{comp, own, verdict}]
    module     : 决策日志模块名（prescreen 批量预筛单独记账）
    返回 {"verdict","confidence","reason","suggest_own"}；失败抛 RuntimeError。
    """
    config = config or load_config()

    comp_txt = f"竞品：{comp_name}" + (f"（品牌:{comp_brand}，型号:{comp_model}）" if comp_brand or comp_model else "")
    cand_txt = "候选自家商品：\n" + "\n".join(f"{i + 1}. {c}" for i, c in enumerate(candidates or [])) \
        if candidates else "候选自家商品：无（请只基于竞品名判断，suggest_own 给空）"

    msgs = [{"role": "system", "content": SYSTEM_PROMPT}]
    for ex in (examples or [])[:20]:
        msgs.append({"role": "user",
                     "content": f"竞品：{ex.get('comp')}\n候选自家商品：1. {ex.get('own')}"})
        msgs.append({"role": "assistant", "content": json.dumps(
            {"verdict": ex.get("verdict", "是"), "confidence": 0.95,
             "reason": "已确认样本", "suggest_own": ex.get("own")}, ensure_ascii=False)})
    msgs.append({"role": "user", "content": comp_txt + "\n" + cand_txt})

    content = call_llm(module, msgs, config=config, timeout=timeout,
                       temperature=0.1, max_tokens=300)
    parsed = _parse_verdict(content)
    if parsed and parsed.get("verdict") in ("是", "否", "拿不准"):
        return parsed
    raise RuntimeError("模型输出格式异常: " + str(content)[:80])


def cache_get(shop, sku):
    try:
        with open(JUDGE_CACHE_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)
        return data.get(shop, {}).get(str(sku))
    except Exception:
        return None


def cache_set(shop, sku, result):
    try:
        data = {}
        if os.path.exists(JUDGE_CACHE_FILE):
            with open(JUDGE_CACHE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
        data.setdefault(shop, {})[str(sku)] = result
        os.makedirs(os.path.dirname(JUDGE_CACHE_FILE), exist_ok=True)
        with open(JUDGE_CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

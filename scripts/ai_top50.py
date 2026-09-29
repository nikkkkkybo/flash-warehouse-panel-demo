# -*- coding: utf-8 -*-
"""
AI 建议：竞品销量前50品 匹配结果（2026-08-14，用户方案）
- 已匹配：对比竞品 vs 自家（月售/价格）→ 经营优化建议（定价对标/补规格/标题/防守）
- 无对应：竞品在卖自家没有 → 是否上新建议（建议引进/暂不引进 + 理由 + 参考定价）
- 待匹配：未处理 → 提示先去匹配

build_summary(shop)  构建 top50 匹配摘要（紧凑文本 + 每品信息）
generate_advice(...) 调 DeepSeek 生成逐品建议（JSON），mock 模式不消耗 API
"""
import json
import os
import sys
import time

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(BASE, "scripts"))
OUTPUT = os.path.join(BASE, "output")

from ai_analysis import load_config  # noqa: E402
from top50 import build_top50  # noqa: E402
from spu_semantic import _clean, _num  # noqa: E402

SYSTEM_PROMPT = """你是闪电仓（五金建材零售）经营顾问。以下是竞品店铺「销量前50品」与自家商品的人工匹配结果（输入已含双方 月售/价格/规格 对比数据）：
- [已匹配]：给该品最该做的一件极简建议（动作+标的+数字，如「降价至¥12对标」「补16-38mm规格」「标题加'自攻螺丝'」）
- [无对应]：给「引/不引」+ 参考价（如「引，参考¥0.3-0.5」「不引，量小」）
- [待匹配]：给「待匹配」
硬约束：写「N 人想买（收藏，非成交）」的品，其需求强度已按 λ 折成有效销量参与门槛与排序——
不得当成成交笔数，也不得说它毫无意义；若某品热度全靠收藏支撑、没有成交记录，
建议要落成"小批量试/先补规格看反应"这类低承诺动作，不要按成交体量给备货或跟价建议；
数据里没有成本与毛利，需要它们才成立的判断写【待核】。
只输出 JSON 数组，不要输出任何其他文字；每条建议 ≤20 字：
[{"id":1,"建议":"..."},{"id":2,"建议":"..."},...]"""


def _sales_txt(p):
    """销量口径文字：成交 / 收藏 / 有效销量 三个数一起给模型，省得它自己加权。"""
    import sales_text as _STx
    iv = p.get("意向数") or 0
    sold = 0.0 if p.get("意向") else (p.get("月售") or 0)
    ev = p.get("有效销量")
    if ev in (None, ""):
        ev = sold + _STx.INTENT_WEIGHT * iv
    if iv:
        return ("成交 %g，%g 人想买（收藏，非成交），有效销量 %g"
                % (sold, iv, ev))
    return "月售 %g（成交）" % (sold,)


def build_summary(shop, prods=None):
    """构建 top50 匹配摘要。返回 (摘要文本, 每品信息 list)。

    prods 可由调用方传面板已缓存的榜单（(shop,"top500")）——不传则自己重算，
    而 build_top50 会重读整张归一化表，点一次按钮多花几秒。
    """
    if prods is None:
        prods = build_top50(shop)
    if not prods:
        return "", []
    # 匹配状态：progress + 汇总库 SPU 层
    prog_file = os.path.join(BASE, "scripts", "progress.json")
    progress = {}
    if os.path.exists(prog_file):
        try:
            progress = json.load(open(prog_file, encoding="utf-8"))
        except Exception:
            pass
    from summary_library import load_library
    spu_own = {}
    try:
        spu_lib, _ = load_library()
        for _, r in spu_lib[spu_lib["竞店名称"].astype(str) == shop].iterrows():
            cid = _clean(r.get("竞品品ID"))
            if cid:
                spu_own.setdefault(cid, {})["自家品名"] = _clean(r.get("自家品名"))
                spu_own.setdefault(cid, {})["自家品组ID"] = _clean(r.get("自家品组ID"))
    except Exception:
        pass
    # 自家 SPU 汇总（月销/售价）
    own = pd.read_excel(os.path.join(OUTPUT, "自家_归一化后.xlsx"), dtype=str).fillna("")
    own_spu_info = {}
    for spuid, g in own.groupby("SPUID"):
        g2 = g.copy()
        g2["__s"] = g2["月销量"].apply(_num).fillna(0.0)
        best = g2.loc[g2["__s"].idxmax()]
        prices = [_num(v) for v in g2["美团闪购渠道售价"] if _num(v) is not None]
        own_spu_info[str(spuid)] = {
            "月销": round(float(g2["__s"].sum()), 0),
            "售价": (round(min(prices), 2), round(max(prices), 2)) if prices else (None, None),
        }

    lines, infos = [], []
    for i, p in enumerate(prods, 1):
        pid = p["comp_id"]
        skus = [s["sku_id"] for s in p.get("SKU列表", [])]
        conf = any(progress.get("confirmed", {}).get(shop, {}).get(s) for s in skus)
        nm = any(progress.get("no_match", {}).get(shop, {}).get(s) for s in skus)
        info = {"id": i, "竞品品ID": pid, "品名": p["品名"], "类目": p["类目"],
                "代表Sku_id": (p.get("SKU列表") or [{}])[0].get("sku_id", ""),
                "竞品月售": p["月售"], "销量口径": "意向" if p.get("意向") else "成交",
                "意向数": p.get("意向数", 0),
                "竞品价格": p["价格"], "竞品规格数": p["规格数"],
                "自家品名": "", "自家月销": "", "自家售价": ""}
        st = _sales_txt(p)
        if conf:
            info["状态"] = "已匹配"
            ow = spu_own.get(pid, {})
            info["自家品名"] = ow.get("自家品名", "")
            ospu = own_spu_info.get(ow.get("自家品组ID", ""), {})
            info["自家月销"] = ospu.get("月销", "")
            info["自家售价"] = ospu.get("售价", "")
            lines.append(
                "%d.[已匹配] 竞品：%s（品牌%s，%s，价%s，%d规格）‖ 自家：%s（月销%s，售价%s）"
                % (i, p["品名"][:30], p["品牌"] or "无", st, p["价格"],
                   p["规格数"], str(info["自家品名"])[:30], info["自家月销"], info["自家售价"]))
        elif nm:
            info["状态"] = "无对应"
            lines.append("%d.[无对应] 竞品：%s（品牌%s，类目%s，%s，价%s）→ 自家无"
                         % (i, p["品名"][:30], p["品牌"] or "无", p["类目"], st, p["价格"]))
        else:
            info["状态"] = "待匹配"
            lines.append("%d.[待匹配] 竞品：%s（%s，价%s）→ 尚未处理"
                         % (i, p["品名"][:30], st, p["价格"]))
        infos.append(info)
    return "\n".join(lines), infos


def _call_deepseek(summary_text, config):
    from ai_client import call_llm
    return call_llm("top50_advice",
                    [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": "竞品销量前50品匹配结果：\n" + summary_text}],
                    config=config, timeout=120,
                    temperature=0.3,
                    max_tokens=1500)   # 低 token：输出极简（每条≤20字），结果体现在输入侧


def _advice_row(info, advice=""):
    """建议表一行 = 该品的事实（来自本地数据）+ 模型给的一句话建议。
    字段集中在这里拼，避免 mock / 正常 / 补空 三条路径各写一份而漏列。"""
    return {"id": info["id"], "竞品品ID": info.get("竞品品ID", ""),
            "代表Sku_id": info.get("代表Sku_id", ""),
            "品名": info["品名"], "状态": info["状态"],
            "竞品月售": info["竞品月售"], "销量口径": info.get("销量口径", "成交"),
            "自家品名": info.get("自家品名", ""), "建议": advice}


def generate_advice(shop, config=None, mock=False, prods=None):
    """生成 top50 逐品建议。返回 (建议 list, 每品信息 list)。"""
    summary, infos = build_summary(shop, prods=prods)
    if not infos:
        raise RuntimeError("暂无销量前50数据")
    if mock:
        result = [_advice_row(i, "（mock）示例建议" if i["状态"] != "待匹配" else "先去匹配")
                  for i in infos]
        return result, infos
    config = config or load_config()
    content = _call_deepseek(summary, config)
    try:
        result = json.loads(content)
    except (ValueError, TypeError):
        s, e = content.find("["), content.rfind("]")
        result = json.loads(content[s:e + 1]) if 0 <= s < e else []
    if not isinstance(result, list):
        result = []
    # 校验 id/状态/建议 字段，缺失补默认
    by_id = {i["id"]: i for i in infos}
    clean = []
    for r in result:
        if not isinstance(r, dict):
            continue
        info = by_id.get(r.get("id"))
        if info is None:
            continue
        clean.append(_advice_row(info, str(r.get("建议", ""))[:200]))
    # 未返回的补空
    returned = {c["id"] for c in clean}
    for i in infos:
        if i["id"] not in returned:
            clean.append(_advice_row(i))
    clean.sort(key=lambda x: x["id"])
    return clean, infos


def save_advice_excel(shop, advice):
    """保存建议到 output/top50_Ai建议.xlsx。"""
    import pandas as _pd
    rows = [{
        "排名": a["id"], "竞品品ID": a.get("竞品品ID", ""), "品名": a["品名"], "状态": a["状态"],
        "竞品月售": a["竞品月售"], "销量口径": a.get("销量口径", "成交"),
        "自家品名": a["自家品名"], "AI建议": a["建议"],
    } for a in advice]
    df = _pd.DataFrame(rows)
    path = os.path.join(OUTPUT, f"top50_Ai建议_{shop}.xlsx")
    df.to_excel(path, index=False)
    return path

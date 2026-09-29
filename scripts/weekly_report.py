# -*- coding: utf-8 -*-
"""每周深报（M6·3.2）：把"当周发生了什么 + 建议采纳得怎么样 + 接下来做什么"落成一份可导出的报告。

分工（与洞察台账同一原则）：**规则层算好全部数字，模型只做排序与解读，不新增数字**。
无 DeepSeek key 时照常产出 0 元结构化版（AI 解读位标"未配置"）。

输入（全部本地、当期）：
  1. progress.json 近 N 天新增确认对（按店与来源分布）
  2. 采纳追踪台账（采纳/已做/忽略计数 + 待回看清单）
  3. AI洞察台账：未处理 Top（按 severity）
  4. 各店观测期数/数据截至（口径戳）
输出：结构化 dict（面板/聊天复用）+ markdown（to_markdown，可打印导出）。
"""
import json
import os
import time
from datetime import datetime, timedelta

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def _shop_obs():
    hist = os.path.join(BASE, "output", "竞店标准库", "价格观测历史.jsonl")
    dates = {}
    try:
        with open(hist, encoding="utf-8") as f:
            for line in f:
                try:
                    d = json.loads(line)
                except Exception:
                    continue
                s, t = str(d.get("店") or ""), str(d.get("t") or "")[:10]
                if s and t:
                    dates.setdefault(s, set()).add(t)
    except OSError:
        pass
    return {s: {"last": max(v), "snaps": len(v)} for s, v in dates.items()}


def build(weeks=1):
    """规则层产出当周深报（0 元），返回结构化 dict。"""
    since = (datetime.now() - timedelta(days=7 * max(1, weeks))).strftime("%Y-%m-%d %H:%M:%S")

    # 1) 当周新增确认对（progress.json 按 time 过滤）
    by_shop, by_source = {}, {}
    n_new = 0
    try:
        with open(os.path.join(BASE, "scripts", "progress.json"), encoding="utf-8") as f:
            prog = json.load(f)
    except Exception:
        prog = {}
    for shop, m in (prog.get("confirmed") or {}).items():
        for _sku, rec in (m or {}).items():
            if str(rec.get("time") or "") >= since:
                n_new += 1
                src = str(rec.get("source") or "?")
                by_source[src] = by_source.get(src, 0) + 1
                by_shop[shop] = by_shop.get(shop, 0) + 1

    # 2) 采纳台账
    import adopt_log as AL
    st = AL.stats()
    pend = AL.pending_review()

    # 3) 台账未处理 Top（severity 降序）
    import insight_store as IS
    led = IS.load().get("items", [])
    stamp = IS.caliber_stamp()
    open_items = [it for it in led if it.get("status") == "new" and it.get("present", True)]
    open_items.sort(key=lambda x: -float(x.get("severity") or 0))

    return {
        "generated_at": _now(),
        "weeks": max(1, weeks),
        "since": since,
        "new_confirms": n_new,
        "new_by_shop": dict(sorted(by_shop.items(), key=lambda x: -x[1])),
        "new_by_source": by_source,
        "adopt": {"total": st["total"], "adopted": st["adopted"], "done": st["done"],
                  "ignored": st["ignored"], "pending_review": len(pend)},
        "pending": pend[:5],
        "open_top": [{"title": it.get("title", ""), "layer": it.get("layer", ""),
                      "shop": it.get("scope", ""), "severity": it.get("severity"),
                      "periods": it.get("periods", 1)}
                     for it in open_items[:8]],
        "open_total": len(open_items),
        "caliber": stamp.get("text", ""),
        "obs_snaps": sorted({v["snaps"] for v in _shop_obs().values()} or {0}),
    }


def to_markdown(res, ai_text=""):
    """规则层结果 → markdown。ai_text 为 AI 解读（可选，无则标"未配置"）。"""
    w = res["weeks"]
    lines = [f"# 每周深报（近 {w} 周）", "",
             f"> 生成 {res['generated_at']}｜口径：{res.get('caliber') or '（未取到）'}", ""]
    lines.append("## 1. 当周确认了多少对应")
    if res["new_confirms"]:
        lines.append(f"- 新增人工/自动确认 **{res['new_confirms']} 条**")
        if res["new_by_shop"]:
            top = list(res["new_by_shop"].items())[:5]
            lines.append("- 按店：" + "、".join(f"{k[:12]} {v} 条" for k, v in top))
        if res["new_by_source"]:
            lines.append("- 按来源：" + "、".join(f"{k} {v}" for k, v in res["new_by_source"].items()))
    else:
        lines.append("- 本周无新增确认（对应工作没推进——这是本周最该补的洞）")
    a = res["adopt"]
    lines += ["", "## 2. 建议采纳情况",
              f"- 采纳追踪流水 {a['total']} 条：采纳 {a['adopted']} / 已做 {a['done']} / 忽略 {a['ignored']}",
              f"- 待回看（满 30 天）{a['pending_review']} 条（效果比对需竞品第二期观测）"]
    lines.append("### 待回看清单（前 5）")
    if res["pending"]:
        for p in res["pending"]:
            obj = p.get("obj")
            subj = obj.get("subject", "") if isinstance(obj, dict) else str(obj)
            lines.append(f"- {subj[:36] or p.get('key', '')[:36]}｜已采纳 {p.get('days')} 天｜{p.get('data_note', '')}")
    else:
        lines.append("- 暂无到期待回看条目")
    lines += ["", "## 3. 当前未处理台账 Top（规则层按严重度）"]
    for it in res["open_top"]:
        lines.append(f"- [{str(it['shop'])[:8]}] {str(it['title'])[:56]}（连 {it['periods']} 期未处理）")
    if not res["open_top"]:
        lines.append("- 台账当前没有未处理条目")
    lines += ["", "## 4. AI 解读"]
    lines.append(ai_text if ai_text
                 else "（未配置 DeepSeek key——结构化数字已在上面，配置后点按钮可获得解读）")
    return "\n".join(lines)


def ai_interpret(res, config=None):
    """可选：把结构化深报交给模型做一段解读（约 ¥0.02-0.05）。无 key 返回空串。"""
    from ai_client import call_llm, load_config
    cfg = config or load_config()
    if not str(cfg.get("api_key") or "").strip():
        return ""
    from weekly_report import to_markdown as _md
    # 只喂结构化事实（不带 AI 解读段）
    txt = _md(res, ai_text="")
    sys_prompt = ("你是给前置仓运营写周报的助理。以下是规则层算好的本周深报（数字全部可信）。"
                  "写一段 ≤200 字的解读：本周做了什么、最重要的两件事、下周先做什么。"
                  "只能引用文本里的数字与品名，纯文本无 markdown 符号。")
    try:
        out = call_llm("insight_deep", [{"role": "system", "content": sys_prompt},
                                        {"role": "user", "content": txt[:3000]}],
                       config=cfg, timeout=120, temperature=0.3, max_tokens=500)
        return str(out or "").strip()
    except Exception as e:
        return f"（AI 解读失败：{e}）"


if __name__ == "__main__":
    res = build()
    print(to_markdown(res))

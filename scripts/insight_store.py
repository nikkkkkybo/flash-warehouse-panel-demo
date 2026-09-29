# -*- coding: utf-8 -*-
"""洞察条目仓库（T4 的"落到实处、便于引用"环节）：把一次性闪现的 AI 结论变成可核对的台账。

一条洞察 = 结论 + 数字证据(带算式) + 来源行ID(可点回) + 数据截至/口径 + 处理状态。
没有这些，洞察就还是"AI 说"：老板问"这数几号的"答不上，员工过两天也不知道该不该做。

三条规矩：
  1 数字永远是当期的：同一条目再次命中时，metrics/title/severity 一律用新数据覆盖，
    只有人的标记（status/note/adopted_at）保留——绝不出现"旧数字配新结论"。
  2 同一件事不重复刷屏：按 key 去重，只累加 periods（连续 N 期未处理 = 催办强度）。
  3 每条目必带 data_date 与口径串，导出时写在标题下方。
"""
import json
import os
import time

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STORE = os.path.join(BASE, "output", "AI洞察.json")
KEEP_MAX = 400
STATUS = ("new", "adopted", "ignored", "done")   # 未处理 / 采纳 / 忽略 / 已做


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def caliber_stamp():
    """这页结论"几号的数据、什么口径"——导出与列表页都要显示，避免被当成现价。"""
    import pandas as pd
    out = {"comp_dates": {}, "own_date": "", "manual": 0, "text": ""}
    try:
        hist = os.path.join(BASE, "output", "竞店标准库", "价格观测历史.jsonl")
        last = {}
        with open(hist, encoding="utf-8") as f:
            for line in f:
                try:
                    d = json.loads(line)
                except Exception:
                    continue
                s, t = str(d.get("店") or ""), str(d.get("t") or "")[:10]
                if s and t and (s not in last or t > last[s]):
                    last[s] = t
        out["comp_dates"] = last
    except Exception:
        pass
    try:
        with open(os.path.join(BASE, "output", "own_sales_meta.json"), encoding="utf-8") as f:
            m = json.load(f)
        d = str(m.get("snapshot_date") or "")
        out["own_date"] = f"{d[:4]}-{d[4:6]}-{d[6:]}" if len(d) == 8 else d
        out["own_note"] = m.get("note", "")
    except Exception:
        pass
    try:
        with open(os.path.join(BASE, "scripts", "progress.json"), encoding="utf-8") as f:
            p = json.load(f)
        out["manual"] = sum(len(v) for v in (p.get("confirmed") or {}).values())
    except Exception:
        pass
    ds = sorted(v for v in out["comp_dates"].values() if v)
    parts = []
    if ds:
        parts.append(f"竞品数据截至 {ds[-1]}（最旧 {ds[0]}）")
    if out.get("own_date"):
        parts.append(f"自家销量快照 {out['own_date']}（{out.get('own_note', '')}）")
    parts.append(f"人工确认 {out['manual']} 条")
    out["text"] = " ｜ ".join(parts)
    out["data_date"] = ds[-1] if ds else time.strftime("%Y-%m-%d")
    return out


def load():
    try:
        with open(STORE, encoding="utf-8") as f:
            doc = json.load(f)
        if isinstance(doc, dict) and isinstance(doc.get("items"), list):
            return doc
    except Exception:
        pass
    return {"items": [], "updated": ""}


def save(doc):
    doc["updated"] = _now()
    os.makedirs(os.path.dirname(STORE), exist_ok=True)
    tmp = STORE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=1))
    os.replace(tmp, STORE)
    return doc


def record_scan(coverage=None, run=None):
    """把这一轮扫描的自述落盘（探测覆盖、发现/新增/更新条数、跑的时间）。

    这些信息过去只存在进程内存里：面板一重启，抽屉首屏就显示"没有任何探测被卡住"，
    把"没跑成"伪装成"一切正常"。落盘后重启也还能看见上一轮的真实状态。
    """
    doc = load()
    run = run or {}
    doc["scan"] = {"at": _now(),
                   "coverage": coverage or {},
                   "found": run.get("found"), "added": run.get("added"),
                   "updated": run.get("updated"), "total": run.get("total")}
    return save(doc)["scan"]


def scan_info():
    """上一轮扫描的自述（没有则空 dict）。"""
    return load().get("scan") or {}


def merge(found, stamp=None):
    """把本轮探测结果并进台账。数字用最新的，人的标记不覆盖。"""
    stamp = stamp or caliber_stamp()
    dd = stamp.get("data_date") or time.strftime("%Y-%m-%d")
    doc = load()
    idx = {it["key"]: it for it in doc["items"]}
    added = updated = 0
    seen = set()
    for f in found:
        key = f["key"]
        seen.add(key)
        cur = idx.get(key)
        if cur:
            same_period = cur.get("last_seen") == dd
            cur.update({"title": f["title"], "metrics": f["metrics"], "refs": f["refs"],
                        "severity": f["severity"], "hint": f.get("hint", ""),
                        "layer": f["layer"], "detector": f["detector"],
                        "scope": f["scope"], "subject": f["subject"],
                        "present": True})
            if not same_period:
                cur["last_seen"] = dd
                cur["periods"] = int(cur.get("periods", 1)) + 1
            updated += 1
        else:
            rec = dict(f)
            rec.update({"first_seen": dd, "last_seen": dd, "periods": 1,
                        "status": "new", "note": "", "present": True,
                        "created_at": _now(), "ai_note": "", "ai_at": ""})
            # M6·3.5（决策3）：数据体检类自动归档——数据核对线索不是行动建议，不计催办；
            # 人工可在台账改回任何状态（此后人的标记不受影响）。
            if rec.get("layer") == "数据体检":
                rec["status"] = "ignored"
                rec["note"] = "自动归档（数据体检类，不计催办）"
                rec["auto_note"] = True
            doc["items"].append(rec)
            idx[key] = rec
            added += 1
    for key, it in idx.items():
        if key not in seen:
            it["present"] = False        # 本期数据里不再出现（处理掉了或数据变了）
    doc["items"] = sorted(doc["items"], key=lambda x: (0 if x.get("present") else 1,
                                                       -float(x.get("severity") or 0)))[:KEEP_MAX]
    doc["caliber"] = stamp.get("text", "")
    doc["data_date"] = dd
    save(doc)
    return {"added": added, "updated": updated, "total": len(doc["items"]), "found": len(found)}


def list_items(layer="", status="", shop="", present_only=True):
    doc = load()
    out = []
    for it in doc["items"]:
        if present_only and not it.get("present", True):
            continue
        if layer and it.get("layer") != layer:
            continue
        if status and it.get("status") != status:
            continue
        if shop and it.get("scope") not in (shop, "跨店", "全部"):
            continue
        out.append(it)
    return {"items": out, "caliber": doc.get("caliber", ""),
            "data_date": doc.get("data_date", ""), "updated": doc.get("updated", "")}


def set_status(key, status, note=""):
    if status not in STATUS:
        return False
    doc = load()
    hit = None
    for it in doc["items"]:
        if it["key"] == key:
            it["status"] = status
            if note:
                it["note"] = note
            it["status_at"] = _now()
            if status in ("adopted", "done") and not it.get("adopted_at"):
                it["adopted_at"] = _now()      # 首次采纳时间（30 天回看的起算点）
            hit = it
    if hit:
        save(doc)
    return bool(hit)


def set_ai_note(keys, notes):
    """写入模型的解释（只覆盖 ai_note，绝不动 metrics——数字只能来自探测层）。"""
    doc = load()
    idx = {it["key"]: it for it in doc["items"]}
    n = 0
    for k in keys:
        rec = idx.get(k.get("key"))
        if rec and k.get("key") in notes:
            rec["ai_note"] = notes[k["key"]].get("explain", "")[:600]
            rec["ai_action"] = notes[k["key"]].get("action", "")[:200]
            rec["ai_at"] = _now()
            n += 1
    if n:
        save(doc)
    return n


# --------------------------------------------------------------------------
# 两个出口
# --------------------------------------------------------------------------
def export_brief(top=12):
    """给老板看的一页摘要（markdown）。只讲结论与动作，数字与口径照抄台账。"""
    d = list_items()
    live = [i for i in d["items"] if i.get("status") != "ignored"]
    # 每层各取前几条：只按 severity 排会让"细节"把"全局"挤掉，而老板要的是两样都有
    per = max(2, top // 2)
    items = [i for i in live if i.get("layer") == "全局"][:per]         + [i for i in live if i.get("layer") == "细节"][:top - per]
    L = [f"# 选品洞察摘要 · 一句话版", "",
         f"> {d.get('caliber') or ''}　｜　生成 {time.strftime('%Y-%m-%d %H:%M')}", ""]
    if not items:
        L.append("本期规则层未发现值得马上动手的问题（宁缺毋滥，不硬凑）。")
    for title, layer in (("🌐 全局问题（人工逐行看不出的）", "全局"),
                         ("🔍 细节可疑（容易忽略的坑）", "细节")):
        sub = [i for i in items if i.get("layer") == layer]
        if not sub:
            continue
        L += [f"## {title}", ""]
        for i in sub:
            ev = "；".join(f"{k}={v}" for k, v in list((i.get("metrics") or {}).items())[:5]
                           if k != "算式")
            L.append(f"- **{i['title']}**（{i.get('scope', '')}）")
            L.append(f"  - 依据：{ev}")
            L.append(f"  - 动作：{i.get('ai_action') or i.get('hint') or ''}")
            if i.get("periods", 1) > 1:
                L.append(f"  - 已连续 {i['periods']} 期未处理")
            L.append("")
    L += ["---", "", "口径提醒：竞品月售为平台观测值（含收藏/意向），自家月销为补货参考快照实测值，"
          "两者只比量级趋势，不做直接相减。"]
    return "\n".join(L)


def export_xlsx(path=None):
    """导出台账（Excel 一行一条，含凭据列），给员工自己排期用。"""
    import pandas as pd
    d = list_items(present_only=False)
    rows = []
    for i in d["items"]:
        rows.append({
            "层级": i.get("layer"), "探测器": i.get("detector"), "范围": i.get("scope"),
            "对象": i.get("subject"), "结论": i.get("title"),
            "数字证据": json.dumps(i.get("metrics") or {}, ensure_ascii=False),
            "算式": (i.get("metrics") or {}).get("算式", ""),
            "建议动作": i.get("ai_action") or i.get("hint", ""),
            "AI解读": i.get("ai_note", ""), "状态": i.get("status"),
            "连续期数": i.get("periods"), "首期": i.get("first_seen"), "最近": i.get("last_seen"),
            "本期是否仍在": "是" if i.get("present", True) else "否",
            "来源凭据": json.dumps(i.get("refs") or [], ensure_ascii=False),
            "数据截至": d.get("data_date", ""), "口径": d.get("caliber", ""),
        })
    df = pd.DataFrame(rows)
    out = path or os.path.join(BASE, "output", "AI洞察台账.xlsx")
    with pd.ExcelWriter(out) as w:
        (df if len(df) else pd.DataFrame(columns=["说明"])).to_excel(w, sheet_name="洞察台账", index=False)
    return out, len(df)


if __name__ == "__main__":
    print(json.dumps(caliber_stamp(), ensure_ascii=False, indent=1))

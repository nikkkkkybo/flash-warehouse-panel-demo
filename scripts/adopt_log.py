# -*- coding: utf-8 -*-
"""采纳追踪台账（M6·3.3）：把"建议 → 人做了什么 → 30 天后有没有用"变成可核对的记录。

一条记录 = {ts, type, key, shop, obj, action, note, data_date, source}
  type   : insight(台账条目) / chat(聊天建议) / action(确认/撤销动作) / review(回看结论)
  action : 登记 / 采纳 / 已做 / 忽略 / 撤销 / 回看
文件   : output/采纳记录.jsonl —— append-only，永不改写（审计留痕；台账条目的
         最新状态仍以 AI洞察.json 为准，本文件是追踪与回看的流水事实源）。

三条规矩：
  1 只追加不覆盖——人的任何标记都是事实，回看不改写历史，只追加新行。
  2 mock/演示一律不落盘（防测试污染），由调用方把关。
  3 回看结论必须有数据截至戳：没到 30 天或竞品只有 1 期观测 → "数据未到"，
    绝不拿一期数据冒充效果。
"""
import json
import os
import time

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FILE = os.path.join(BASE, "output", "采纳记录.jsonl")

TYPE_INSIGHT = "insight"
TYPE_CHAT = "chat"
TYPE_ADVICE = "advice"
TYPE_ACTION = "action"
TYPE_REVIEW = "review"
ACTION_DONE = "已做"
ACTION_ADOPT = "采纳"
ACTION_IGNORE = "忽略"
REVIEW_WIN = "有效"
REVIEW_NO_DATA = "数据未到"


def _now():
    return time.strftime("%Y-%m-%d %H:%M:%S")


def record(otype, key, obj=None, action="", note="", data_date="", source="",
           shop=""):
    """追加一条记录（append-only）。返回写入的 dict（失败返回 None，不抛异常）。"""
    if not otype:
        return None
    ent = {"ts": _now(), "type": str(otype), "key": str(key or ""),
           "shop": str(shop or ""), "obj": obj if isinstance(obj, dict) else (str(obj or "")),
           "action": str(action or ""), "note": str(note or "")[:400],
           "data_date": str(data_date or ""), "source": str(source or "")}
    try:
        os.makedirs(os.path.dirname(FILE), exist_ok=True)
        with open(FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(ent, ensure_ascii=False) + "\n")
        return ent
    except Exception:
        return None


def load_entries():
    try:
        with open(FILE, encoding="utf-8") as f:
            return [json.loads(l) for l in f if l.strip()]
    except (FileNotFoundError, json.JSONDecodeError):
        return []
    except Exception:
        return []


def stats():
    """按动作计数 + 类型计数（台账页/体检用）。"""
    ents = load_entries()
    st = {"total": len(ents), "by_action": {}, "by_type": {},
          "adopted": 0, "done": 0, "ignored": 0}
    for e in ents:
        a = str(e.get("action") or "-")
        st["by_action"][a] = st["by_action"].get(a, 0) + 1
        st["by_type"][str(e.get("type") or "-")] = \
            st["by_type"].get(str(e.get("type") or "-"), 0) + 1
    st["adopted"] = st["by_action"].get(ACTION_ADOPT, 0)
    st["done"] = st["by_action"].get(ACTION_DONE, 0)
    st["ignored"] = st["by_action"].get(ACTION_IGNORE, 0)
    return st


def latest_by_key():
    """每条 key 取最新动作（台账页按钮切换后的最终状态）。"""
    out = {}
    for e in load_entries():
        k = str(e.get("key") or "")
        if k:
            out[k] = e
    return out


def pending_review(days=30):
    """采纳超过 days 天且之后没有回看结论的条目。

    返回 [{key, shop, ts, days, type, obj, note}] —— 指标比对在 web_panel 侧做
    （那里有自家表/观测历史），本层只负责"谁到期待回看"。"""
    seen_outcome = set()
    for e in load_entries():
        if str(e.get("type")) == TYPE_REVIEW:
            seen_outcome.add(str(e.get("key") or ""))
    latest = latest_by_key()
    now = time.time()
    out = []
    for k, e in latest.items():
        if str(e.get("action")) not in (ACTION_ADOPT, ACTION_DONE):
            continue
        if k in seen_outcome:
            continue
        ts = str(e.get("ts") or "")
        try:
            t0 = time.mktime(time.strptime(ts[:19], "%Y-%m-%d %H:%M:%S"))
        except (ValueError, TypeError):
            continue
        days_elapsed = (now - t0) / 86400.0
        if days_elapsed >= days:
            out.append({"key": k, "shop": e.get("shop", ""), "ts": ts,
                        "days": int(days_elapsed), "type": e.get("type", ""),
                        "obj": e.get("obj", ""), "note": e.get("note", "")})
    out.sort(key=lambda x: -x["days"])
    return out


def has_shop_obs2(shop):
    """该店竞品观测期数是否 ≥2（回看的前置数据条件；缺 → 回看只能标"数据未到"）。"""
    try:
        hist = os.path.join(BASE, "output", "竞店标准库", "价格观测历史.jsonl")
        dates = set()
        with open(hist, encoding="utf-8") as f:
            for line in f:
                try:
                    d = json.loads(line)
                except Exception:
                    continue
                if str(d.get("店") or "") == str(shop or ""):
                    dates.add(str(d.get("t") or "")[:10])
        return len(dates) >= 2
    except Exception:
        return False

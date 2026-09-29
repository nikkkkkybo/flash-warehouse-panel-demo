# -*- coding: utf-8 -*-
"""
对话记忆（chat_memory）：聊天助手的服务端会话存储，可回溯。

设计：
- 每个会话一个 append-only jsonl：output/ai_chat/<sid>.jsonl（逐条追加，永不改写历史）。
- 索引 output/ai_chat/_sessions.json：{sid: {title, shop, created, updated, count}}，
  原子替换写；title=首条用户消息截断（会话列表可读）。
- 每条记录带 **当时的范围(shop) 与数据日期(data_date)**——回溯时能知道
  "这条答案是基于哪天的数据说的"，这是回溯的关键，不只是存文字。
- 所有异常尽量不抛（存储失败不影响当次回答返回），sid 严格白名单字符。
"""
import json
import os
import re
import time

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CHAT_DIR = os.path.join(BASE, "output", "ai_chat")
INDEX_FILE = os.path.join(CHAT_DIR, "_sessions.json")

_SID_RE = re.compile(r"[^A-Za-z0-9_-]")


def _safe_sid(sid):
    return _SID_RE.sub("", str(sid or ""))[:64]


def _index():
    try:
        with open(INDEX_FILE, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _write_index(d):
    try:
        os.makedirs(CHAT_DIR, exist_ok=True)
        tmp = INDEX_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=1)
        os.replace(tmp, INDEX_FILE)
    except Exception:
        pass


def list_sessions(limit=30):
    d = _index()
    out = []
    for sid, m in d.items():
        if not os.path.exists(os.path.join(CHAT_DIR, sid + ".jsonl")):
            continue
        out.append({"sid": sid, **(m if isinstance(m, dict) else {})})
    out.sort(key=lambda x: -(x.get("updated") or 0))
    return out[:limit]


def create_session(shop=""):
    os.makedirs(CHAT_DIR, exist_ok=True)
    sid = time.strftime("%Y%m%d_%H%M%S_") + os.urandom(2).hex()
    d = _index()
    d[sid] = {"title": "新会话", "shop": shop or "", "created": time.time(),
              "updated": time.time(), "count": 0}
    _write_index(d)
    return {"sid": sid, **d[sid]}


def append(sid, record):
    """追加一条记录（role/text/trace/links/focus/shop/data_date）。
    返回该记录的 0 基行号（供单条导出定位）；存储失败返回 None（不影响回答）。"""
    sid = _safe_sid(sid)
    if not sid:
        return None
    try:
        os.makedirs(CHAT_DIR, exist_ok=True)
        rec = {"t": time.strftime("%Y-%m-%d %H:%M:%S"),
               "role": str(record.get("role") or "user"),
               "text": str(record.get("text") or "")[:4000],
               "shop": str(record.get("shop") or ""),
               "data_date": str(record.get("data_date") or "")}
        if record.get("trace"):
            rec["trace"] = record["trace"][:8]
        if record.get("links"):
            rec["links"] = record["links"][:6]
        if record.get("focus"):
            rec["focus"] = record["focus"][:24]
        d = _index()
        m = d.get(sid) or {}
        line_no = int(m.get("count") or 0)
        with open(os.path.join(CHAT_DIR, sid + ".jsonl"), "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        m.setdefault("created", time.time())
        m["updated"] = time.time()
        m["count"] = line_no + 1
        if rec["role"] == "user" and (m.get("title") in ("", "新会话", None)):
            m["title"] = rec["text"][:24] or "新会话"
        if rec.get("shop"):
            m["shop"] = rec["shop"]
        d[sid] = m
        _write_index(d)
        return line_no
    except Exception:
        return None


def load_messages(sid, limit=100):
    sid = _safe_sid(sid)
    if not sid:
        return []
    out = []
    try:
        with open(os.path.join(CHAT_DIR, sid + ".jsonl"), encoding="utf-8") as f:
            for line in f:
                try:
                    out.append(json.loads(line))
                except (ValueError, TypeError):
                    continue
    except OSError:
        return []
    return out[-limit:]


def rename_session(sid, title):
    sid = _safe_sid(sid)
    if not sid:
        return False
    d = _index()
    if sid not in d:
        return False
    d[sid]["title"] = str(title or "").strip()[:24] or d[sid].get("title") or "会话"
    _write_index(d)
    return True


def delete_session(sid):
    sid = _safe_sid(sid)
    d = _index()
    if sid not in d:
        return False
    d.pop(sid, None)
    _write_index(d)
    try:
        os.remove(os.path.join(CHAT_DIR, sid + ".jsonl"))
    except OSError:
        pass
    return True

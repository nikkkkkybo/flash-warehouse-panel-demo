# -*- coding: utf-8 -*-
"""店↔仓对应表（config/store_warehouse_map.yaml）的唯一读写口。

存在理由：美团的两套编码互不相通（店编码≠仓编码），补货参考只给「仓」维度销量，
店对店比较必须由人指认「哪家店用哪个仓的数」。旧实现靠"仓名关键词包含匹配店名"，
匹配不上就静默回退多仓合计 = 拿别店销量冒充本店销量。本模块把对应关系变成显式数据，
并提供四种状态供界面区分：mapped / pending / unmapped / no_data。

source 语义：manual=人确认过（机器不得覆盖）；pending=系统按名称推定（界面挂待确认）。
"""
import os
import re
import time

import yaml

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MAP_FILE = os.path.join(BASE, "config", "store_warehouse_map.yaml")

_HEADER = """# 店↔仓对应表：补货参考快照里的「仓」 ↔ 自家「门店」
#
# 为什么需要这张表：美团的两套编码体系互不相通——店编码 ≠ 仓编码，
# 补货参考表只给仓维度销量，店对店比较必须靠人指认「这家店用哪个仓的数」。
# source: manual=人已确认（机器不得覆盖）; pending=系统按名称推定（界面挂待确认徽章）
# 未指认的店：页面显示告警，且不出现任何店对店销量结论（禁止静默回退合计）。
"""

_cache = {"mtime": None, "data": None}


def _norm(v):
    return str(v if v is not None else "").strip()


def load(force=False):
    """→ {仓编码: {wh_name, keyword, store_code, store_name, source, note}}；文件缺失 → {}。"""
    try:
        mt = os.path.getmtime(MAP_FILE)
    except OSError:
        _cache["mtime"], _cache["data"] = None, {}
        return {}
    if not force and _cache["data"] is not None and _cache["mtime"] == mt:
        return _cache["data"]
    try:
        with open(MAP_FILE, encoding="utf-8") as f:
            doc = yaml.safe_load(f) or {}
        raw = doc.get("warehouses") or {}
    except Exception as e:
        print("[店仓映射] 读取失败，按未配置处理:", e)
        raw = {}
    out = {}
    for code, v in (raw or {}).items():
        code = _norm(code)
        if not code or not isinstance(v, dict):
            continue
        out[code] = {"wh_name": _norm(v.get("wh_name")),
                     "keyword": _norm(v.get("keyword")),
                     "store_code": _norm(v.get("store_code")),
                     "store_name": _norm(v.get("store_name")),
                     "source": _norm(v.get("source")) or "pending",
                     "note": _norm(v.get("note"))}
    _cache["mtime"], _cache["data"] = mt, out
    return out


def save(m):
    """原子写；固定注释头由本模块持有，避免每次保存丢注释。"""
    os.makedirs(os.path.dirname(MAP_FILE), exist_ok=True)
    body = {"updated": time.strftime("%Y-%m-%d"),
            "warehouses": {str(k): v for k, v in sorted((m or {}).items())}}
    txt = _HEADER + yaml.safe_dump(body, allow_unicode=True, sort_keys=False,
                                   default_flow_style=False)
    tmp = MAP_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        f.write(txt)
    os.replace(tmp, MAP_FILE)
    _cache["mtime"], _cache["data"] = None, None
    return load(force=True)


def keyword_of(wh_name):
    """仓名 → 关键词（示例共享仓1→ 示例片区A；示例片区B仓 → 示例片区B）。仅用于推定提示。"""
    wn = _norm(wh_name)
    kw = wn
    mm = re.search(r"[（(](.+?)[)）]", wn)
    if mm:
        kw = mm.group(1)
    return kw[:-1] if kw.endswith("仓") else kw


def guess(store_names, wh_name, wh_code=""):
    """按关键词在自家店清单里推定；命中唯一才给建议。→ (store_name, store_code) 或 None。"""
    kw = keyword_of(wh_name)
    if not kw:
        return None
    hits = [s for s in (store_names or []) if kw in _norm(s.get("name"))]
    return (hits[0]["name"], _norm(hits[0].get("code"))) if len(hits) == 1 else None


def set_mapping(wh_code, wh_name, store_name, store_code="", source="manual", note=""):
    """指认一个仓 → 一家店。source=manual 时覆盖 pending（人工确认不可被机器覆盖）。"""
    m = load()
    code = _norm(wh_code) or _norm(wh_name)
    m[code] = {"wh_name": _norm(wh_name), "keyword": keyword_of(wh_name),
               "store_code": _norm(store_code), "store_name": _norm(store_name),
               "source": source, "note": note}
    return save(m)


def clear_mapping(wh_code):
    m = load()
    if _norm(wh_code) in m:
        m.pop(_norm(wh_code))
        save(m)
    return m


def warehouse_for_store(store_name):
    """店名 → 已登记的仓 {code, wh_name, keyword, source}；没有 → None。"""
    sn = _norm(store_name)
    if not sn:
        return None
    for code, v in load().items():
        if v.get("store_name") == sn:
            return {"code": code, "wh_name": v["wh_name"],
                    "keyword": v["keyword"] or keyword_of(v["wh_name"]),
                    "source": v["source"]}
    return None


def status(store_name, present_codes=()):
    """店对店销量口径的四种状态，界面据此决定「显示数值 / 挂待确认 / 出告警」。

    present_codes = 当期分仓快照里真实有数据的仓编码集合。
      mapped   已人工指认且当期有数据      → 用仓级口径
      pending  系统推定未确认且当期有数据  → 用仓级口径但必须挂「待确认」
      no_data  已指认但当期快照没这个仓    → 不得显示该店销量，提示补数据
      unmapped 从未指认                    → 不得显示该店销量，指向认界面
    """
    w = warehouse_for_store(store_name)
    if not w:
        return {"state": "unmapped", "usable": False, "store": _norm(store_name),
                "wh_name": "", "code": "", "kw": "", "source": ""}
    if str(w["code"]) not in set(str(c) for c in present_codes):
        return {"state": "no_data", "usable": False, "store": _norm(store_name),
                "wh_name": w["wh_name"], "code": w["code"],
                "kw": w["keyword"], "source": w["source"]}
    state = w["source"] if w["source"] in ("manual", "pending") else "manual"
    return {"state": state, "usable": True, "store": _norm(store_name),
            "wh_name": w["wh_name"], "code": w["code"],
            "kw": w["keyword"], "source": w["source"]}


def summary():
    """工具页/体检用：全部条目 + 缺数据的仓。"""
    return load()


if __name__ == "__main__":
    import json
    print(json.dumps(summary(), ensure_ascii=False, indent=1))
    for s in ("示例自家店1", "示例自家店3"):
        print(s, "→", status(s, ["1000001", "1000002"])["state"])

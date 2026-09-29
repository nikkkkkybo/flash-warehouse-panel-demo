# -*- coding: utf-8 -*-
"""
引进候选池（③ 候选池交互化）
====================================
两级结构：总库（全部条目）→ 区视图（② 分区落地后按区过滤；当前按店铺分组展示）。
- 条目 = (竞店, 竞品品ID)。入池时快照品信息（品名/品牌/类目/价格/月售/图片/SKU条码），
  之后竞店重导表不影响池内展示（池是决策工作台，不是竞品镜像）。
- 绑定组 = 跨店同品组：多个店的条目绑成一组，可再关联一个自家品组ID。
  作用：① 同品跨店全面分析（价格/销量/规格横比）② ②区热销榜的同品合并地基。

绑定推荐（竞品×竞品，六层级联，便宜层先跑）：
  L1 条码交集        → 0.95 自动
  L2 品牌+型号+类目   → 0.85 自动
  L3 清洗名称相等     → 0.80 自动
  L4 名称 bigram Jaccard → 0.45~0.75 自动（越像分越高）
  L5 embedding 余弦  → 配置了 embeddings 才启用（未配置返回 None 自动退化）
  L6 LLM 裁决        → 按需点按钮（结果缓存复用），口径同 ai_match：品牌明确不同=不同品

入池即自动推荐（本地零成本）：add_entry 时对新条目与池内其他店条目双向落盘"推荐"，
score≥0.80 的对进入"高置信待确认"清单（pending_pairs），池页一键批量绑定；
建组清互指、出池清引用、忽略（suggest_dismiss）双向移除。

存储：output/候选池.json（append/update 由本模块统一收口；文件锁走 data_lock 由调用方持有）
"""
import json
import os
import re
import threading
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
POOL_FILE = os.path.join(BASE, "output", "候选池.json")
LLM_CACHE_FILE = os.path.join(BASE, "output", "_pool_llm_cache.json")

_POOL_LOCK = threading.Lock()


# ---------------------------------------------------------------------------
# 存储层
# ---------------------------------------------------------------------------

def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def load_pool():
    """读候选池；文件不存在/坏文件 → 空池（永不抛异常）。"""
    with _POOL_LOCK:
        if not os.path.exists(POOL_FILE):
            return {"entries": [], "groups": []}
        try:
            with open(POOL_FILE, encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                return {"entries": [], "groups": []}
            data.setdefault("entries", [])
            data.setdefault("groups", [])
            return data
        except (ValueError, OSError):
            return {"entries": [], "groups": []}


def save_pool(pool):
    with _POOL_LOCK:
        os.makedirs(os.path.dirname(POOL_FILE), exist_ok=True)
        tmp = POOL_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(pool, f, ensure_ascii=False, indent=2)
        os.replace(tmp, POOL_FILE)


def _next_id(prefix, existing):
    n = 0
    for i in range(1, 100000):
        cid = f"{prefix}{i:04d}"
        if cid not in existing:
            return cid
        n = i
    return f"{prefix}{n + 1:04d}"


# ---------------------------------------------------------------------------
# 条目 CRUD
# ---------------------------------------------------------------------------

def find_entry(pool, shop, comp_id):
    """按 (店, 品ID) 找池内条目（防重复入池）。"""
    for e in pool["entries"]:
        if e.get("shop") == shop and str(e.get("comp_id")) == str(comp_id):
            return e
    return None


def add_entry(shop, prod, source="手动入池"):
    """入池一条：prod = top50 品级 dict（品名/品牌/类目/价格/月售/图片/SKU列表）。
    已在池 → 返回 (entry, False)；成功 → (entry, True)。"""
    pool = load_pool()
    exist = find_entry(pool, shop, prod.get("comp_id"))
    if exist is not None:
        return exist, False
    skus = []
    for s in (prod.get("SKU列表") or [])[:20]:
        skus.append({
            "sku_id": str(s.get("sku_id") or ""),
            "规格": str(s.get("规格") or ""),
            "价格": str(s.get("价格") or ""),
            "月售": s.get("月售", 0),
            "条码": str(s.get("条码") or ""),
            "图片": str(s.get("图片") or ""),
        })
    eid = _next_id("E", {e["eid"] for e in pool["entries"]})
    entry = {
        "eid": eid,
        "shop": shop,
        "comp_id": str(prod.get("comp_id") or ""),
        "品名": str(prod.get("品名") or ""),
        "品牌": str(prod.get("品牌") or ""),
        "类目": str(prod.get("类目") or ""),
        "价格": str(prod.get("价格") or ""),
        "原价": str(prod.get("原价") or ""),
        "月售": prod.get("月售", 0),
        "图片": str(prod.get("图片") or ""),
        "规格数": prod.get("规格数", len(skus)),
        "SKU列表": skus,
        "入池时间": _now(),
        "来源": source,
        "状态": "池内",
        "组": "",
    }
    pool["entries"].append(entry)

    # 入池自动推荐（本地零成本：不走 embedding / LLM；高置信对进待确认清单）
    sug = _suggest_in(entry, pool, use_embed=False)
    entry["推荐"] = [dict(s) for s in sug if s.get("score", 0) > 0][:8]
    smap = {s["eid"]: s for s in sug if s.get("score", 0) > 0}
    for other in pool["entries"]:
        s = smap.get(other["eid"])
        if s is None:
            continue
        rec = {"eid": entry["eid"], "shop": entry["shop"], "品名": entry["品名"],
               "图片": entry["图片"], "价格": entry["价格"], "月售": entry["月售"],
               "品牌": entry["品牌"], "score": s["score"], "layer": s.get("layer", ""),
               "reason": s.get("reason", ""), "组": ""}
        lst = [x for x in (other.get("推荐") or []) if x.get("eid") != entry["eid"]]
        lst.append(rec)
        lst.sort(key=lambda x: -float(x.get("score") or 0))
        other["推荐"] = lst[:12]

    save_pool(pool)
    return entry, True


def remove_entry(eid):
    """出池；若在绑定组中则先移出（组空自动删）。"""
    pool = load_pool()
    e = next((x for x in pool["entries"] if x["eid"] == eid), None)
    if e is None:
        return False
    gid = e.get("组") or ""
    pool["entries"] = [x for x in pool["entries"] if x["eid"] != eid]
    if gid:
        _group_remove_raw(pool, gid, eid)
    # 出池清引用：其他条目的推荐里移除该 eid
    for x in pool["entries"]:
        if x.get("推荐"):
            x["推荐"] = [s for s in x["推荐"] if s.get("eid") != eid]
    save_pool(pool)
    return True


# ---------------------------------------------------------------------------
# 绑定组 CRUD
# ---------------------------------------------------------------------------

def _group_remove_raw(pool, gid, eid):
    """组内移除（不落盘）；组空删除组。"""
    g = next((x for x in pool["groups"] if x["gid"] == gid), None)
    if g is None:
        return
    g["eids"] = [i for i in g.get("eids", []) if i != eid]
    if not g["eids"]:
        pool["groups"] = [x for x in pool["groups"] if x["gid"] != gid]


def _strip_group_sugs(pool, eids):
    """组内条目互为同品，不再是候选：清除组内互指的推荐（不落盘）。"""
    eset = {e for e in eids}
    for x in pool["entries"]:
        if x["eid"] in eset and x.get("推荐"):
            x["推荐"] = [s for s in x["推荐"] if s.get("eid") not in eset]


def create_group(eids, name=""):
    """新建绑定组：条目不能重复入组（先移出原组）。返回组 dict。"""
    pool = load_pool()
    valid = []
    for eid in eids:
        e = next((x for x in pool["entries"] if x["eid"] == eid), None)
        if e is None:
            continue
        if e.get("组"):
            _group_remove_raw(pool, e["组"], eid)
        valid.append(eid)
    if not valid:
        return None
    gid = _next_id("G", {g["gid"] for g in pool["groups"]})
    if not name:
        first = next((x for x in pool["entries"] if x["eid"] == valid[0]), {})
        name = (first.get("品名") or "绑定组")[:30]
    pool["groups"].append({
        "gid": gid, "名称": name, "eids": valid,
        "own_gid": "", "own_name": "", "创建时间": _now(), "备注": "",
    })
    for eid in valid:
        e = next((x for x in pool["entries"] if x["eid"] == eid), None)
        if e is not None:
            e["组"] = gid
    _strip_group_sugs(pool, valid)   # 组内互指的候选清除
    save_pool(pool)
    return next(g for g in pool["groups"] if g["gid"] == gid)


def group_add(gid, eid):
    """条目加入组（先移出原组；组不存在返回 None）。"""
    pool = load_pool()
    g = next((x for x in pool["groups"] if x["gid"] == gid), None)
    e = next((x for x in pool["entries"] if x["eid"] == eid), None)
    if g is None or e is None:
        return None
    if e.get("组") and e["组"] != gid:
        _group_remove_raw(pool, e["组"], eid)
    if eid not in g["eids"]:
        g["eids"].append(eid)
    e["组"] = gid
    _strip_group_sugs(pool, [eid] + list(g.get("eids", [])))
    save_pool(pool)
    return g


def group_remove(gid, eid):
    """移出组（组空删除组）。"""
    pool = load_pool()
    e = next((x for x in pool["entries"] if x["eid"] == eid), None)
    if e is not None:
        e["组"] = ""
    _group_remove_raw(pool, gid, eid)
    save_pool(pool)
    return True


def group_rename(gid, name):
    pool = load_pool()
    g = next((x for x in pool["groups"] if x["gid"] == gid), None)
    if g is None:
        return None
    g["名称"] = (name or g["名称"])[:60]
    save_pool(pool)
    return g


def group_delete(gid):
    """解散组：条目全部回到未绑定。"""
    pool = load_pool()
    g = next((x for x in pool["groups"] if x["gid"] == gid), None)
    if g is None:
        return False
    for eid in g.get("eids", []):
        e = next((x for x in pool["entries"] if x["eid"] == eid), None)
        if e is not None:
            e["组"] = ""
    pool["groups"] = [x for x in pool["groups"] if x["gid"] != gid]
    save_pool(pool)
    return True


def group_link_own(gid, own_gid, own_name):
    """绑定组关联自家品组（辅助商品对应：一组的跨店竞品同时挂到自家品）。"""
    pool = load_pool()
    g = next((x for x in pool["groups"] if x["gid"] == gid), None)
    if g is None:
        return None
    g["own_gid"] = str(own_gid or "")
    g["own_name"] = str(own_name or "")
    save_pool(pool)
    return g


# ---------------------------------------------------------------------------
# 绑定推荐：六层级联（L1-L5 本地自动；L6 LLM 按需）
# ---------------------------------------------------------------------------

def _norm_barcode(b):
    s = re.sub(r"\D", "", str(b or ""))
    return s.lstrip("0") if s else ""


def _name_tokens(name):
    """品名 → bigram 集合（中文友好：'公牛转换器' → {'公牛','牛转','转换','换器'}）。
    英文/数字连续段整体作为一个 token。"""
    s = str(name or "").strip().lower()
    if not s:
        return set()
    s = re.sub(r"[^\w\u4e00-\u9fff]+", " ", s)
    toks = set()
    for part in s.split():
        if re.fullmatch(r"[a-z0-9]+", part):
            toks.add(part)
        else:
            for i in range(len(part) - 1):
                toks.add(part[i:i + 2])
            if len(part) == 1:
                toks.add(part)
    return toks


def _jaccard(a, b):
    if not a or not b:
        return 0.0
    inter = len(a & b)
    return inter / (len(a) + len(b) - inter)


def _brand_conflict(b1, b2):
    """品牌冲突：两边都明确且不同 → True（跨品牌不可能是同品）。"""
    b1, b2 = str(b1 or "").strip(), str(b2 or "").strip()
    if not b1 or not b2 or b1 in ("无", "无品牌", "-") or b2 in ("无", "无品牌", "-"):
        return False
    return b1 != b2


def _embed_scores(texts, cache_key):
    """L5 embedding：未配置/失败返回 None（null-object 退化）。结果进程内缓存。"""
    global _EMBED_CACHE
    try:
        from ai_client import load_config, embeddings_config, call_embedding
        cfg = load_config()
        if not embeddings_config(cfg):
            return None
        key = tuple(texts) + (cache_key,)
        if key in _EMBED_CACHE:
            return _EMBED_CACHE[key]
        vecs = call_embedding(texts, config=cfg)
        if not vecs or len(vecs) != len(texts):
            return None
        out = vecs
        _EMBED_CACHE[key] = out
        return out
    except Exception:
        return None


_EMBED_CACHE = {}


def _suggest_in(me, pool, use_embed=True):
    """对池内其他店条目打分（内存池，不落盘）。返回全部候选按分降序（含0分）。"""
    my_barcodes = {_norm_barcode(s.get("条码")) for s in me.get("SKU列表", [])}
    my_barcodes.discard("")
    my_toks = _name_tokens(me.get("品名"))

    # 跨店 + 品牌冲突排除（省 LLM 钱 & 防误绑）
    cands = [e for e in pool["entries"]
             if e["eid"] != me["eid"] and e.get("shop") != me.get("shop")
             and not _brand_conflict(me.get("品牌"), e.get("品牌"))]
    if not cands:
        return []

    # L5 embedding（一次性批量；自动路径 use_embed=False 保持零成本）
    emb_vecs = None
    if use_embed and cands:
        texts = [me.get("品名") or ""] + [e.get("品名") or "" for e in cands]
        emb_vecs = _embed_scores(texts, "pool_bind")

    out = []
    for idx, e in enumerate(cands):
        score, layer, reason = 0.0, "", ""
        # L1 条码
        barcodes = {_norm_barcode(s.get("条码")) for s in e.get("SKU列表", [])}
        barcodes.discard("")
        if my_barcodes and barcodes and (my_barcodes & barcodes):
            score, layer, reason = 0.95, "L1 条码", "SKU条码一致"
        # L2 品牌+型号口径近似：品牌一致（都明确）+ 类目一致 + 名称Jaccard高
        else:
            toks = _name_tokens(e.get("品名"))
            j = _jaccard(my_toks, toks)
            same_brand = (str(me.get("品牌") or "").strip()
                          and str(me.get("品牌") or "").strip()
                          == str(e.get("品牌") or "").strip())
            same_cat = (str(me.get("类目") or "").strip()
                        == str(e.get("类目") or "").strip() and bool(me.get("类目")))
            if j >= 0.60 and same_brand and same_cat:
                score, layer, reason = max(0.85, 0.5 + j * 0.45), "L2 品牌+名称", f"同品牌同类目，名称相似{j:.0%}"
            elif j >= 0.60:
                score, layer, reason = 0.80, "L3 名称相似", f"名称相似度 {j:.0%}"
            elif j >= 0.35:
                score, layer, reason = 0.5 + j * 0.4, "L4 名称部分相似", f"名称相似度 {j:.0%}"
            elif same_cat and same_brand:
                score, layer, reason = 0.45, "L4 同类目同品牌", "类目品牌一致，名称差异大"
            # L5 embedding 加权（仅提高置信，不单独成层）
            if emb_vecs and score > 0 and score < 0.95:
                try:
                    import math
                    a, b = emb_vecs[0], emb_vecs[idx + 1]
                    cos = sum(x * y for x, y in zip(a, b)) / (
                        math.sqrt(sum(x * x for x in a)) * math.sqrt(sum(x * x for x in b)) + 1e-9)
                    if cos >= 0.86:
                        score = max(score, min(0.88, 0.5 + cos * 0.4))
                        layer = (layer + "+语义") if layer else "L5 语义"
                except Exception:
                    pass
        out.append({
            "eid": e["eid"], "shop": e.get("shop", ""), "品名": e.get("品名", ""),
            "图片": e.get("图片", ""), "价格": e.get("价格", ""), "月售": e.get("月售", 0),
            "品牌": e.get("品牌", ""), "score": round(score, 2), "layer": layer,
            "reason": reason, "组": e.get("组", ""),
        })
    out.sort(key=lambda x: -x["score"])
    return out


def bind_suggest(eid, top_n=8):
    """给一条目找池内其他店的同品候选（六层前五层本地打分；LLM 层用户点按钮）。
    返回 [{eid, shop, 品名, 图片, 价格, 月售, score, layer, reason}] 按分降序。"""
    pool = load_pool()
    me = next((x for x in pool["entries"] if x["eid"] == eid), None)
    if me is None:
        return []
    return _suggest_in(me, pool)[:top_n]


def refresh_suggest(eid, top_n=8):
    """重算一条目的推荐并落盘（打开推荐弹窗时调用；自愈历史/过期数据）。"""
    pool = load_pool()
    me = next((x for x in pool["entries"] if x["eid"] == eid), None)
    if me is None:
        return []
    sug = _suggest_in(me, pool)
    me["推荐"] = [dict(s) for s in sug if s.get("score", 0) > 0][:8]
    save_pool(pool)
    return sug[:top_n]


def suggest_dismiss(eid, eid2):
    """忽略一条候选（双向移除；重算 refresh_suggest 会恢复——忽略只影响存储态）。"""
    pool = load_pool()
    changed = False
    pair = {eid, eid2}
    for e in pool["entries"]:
        if e["eid"] in pair and e.get("推荐"):
            before = len(e["推荐"])
            e["推荐"] = [s for s in e["推荐"] if s.get("eid") not in pair]
            if len(e["推荐"]) != before:
                changed = True
    if changed:
        save_pool(pool)
    return changed


def _sug_summary(e):
    return {"eid": e["eid"], "shop": e.get("shop", ""), "品名": e.get("品名", ""),
            "图片": e.get("图片", ""), "价格": e.get("价格", ""), "原价": e.get("原价", ""),
            "月售": e.get("月售", 0), "品牌": e.get("品牌", "")}


def pending_pairs(threshold=0.80):
    """高置信待确认对：存储推荐中 score≥阈值、双方都未绑定、按对去重。降序。"""
    pool = load_pool()
    eidx = {e["eid"]: e for e in pool["entries"]}
    seen, out = set(), []
    for e in pool["entries"]:
        if e.get("组"):
            continue
        for s in (e.get("推荐") or []):
            score = float(s.get("score") or 0)
            t = eidx.get(s.get("eid"))
            if score < threshold or t is None or t.get("组"):
                continue
            key = frozenset((e["eid"], t["eid"]))
            if key in seen:
                continue
            seen.add(key)
            out.append({"a": _sug_summary(e), "b": _sug_summary(t),
                        "score": score, "layer": s.get("layer", ""),
                        "reason": s.get("reason", "")})
    out.sort(key=lambda x: -x["score"])
    return out


def backfill_all():
    """全量重算所有条目的推荐（历史数据补齐/修复，本地零成本）。返回条目数。"""
    pool = load_pool()
    for e in pool["entries"]:
        sug = _suggest_in(e, pool, use_embed=False)
        e["推荐"] = [dict(s) for s in sug if s.get("score", 0) > 0][:8]
    save_pool(pool)
    return len(pool["entries"])


# ---------------------------------------------------------------------------
# L6 LLM 同品裁决（缓存复用，点按钮才花钱）
# ---------------------------------------------------------------------------

_LLM_SYSTEM = """你是闪电仓（五金建材零售）的商品比对专家。判断两家不同竞品店的"商品A"与"商品B"是否为同一个品（SPU级）。

判断口径：
1. 类目、用途、核心参数（功率/容量/尺寸/电压/材质等）一致，视为同一个品；
2. 规格/颜色/档位差异（如 5米 vs 3米、白色 vs 金色）不算不同品，仍归同一品；
3. 品牌不同（且都明确）则为不同品；
4. 信息不足给"拿不准"，不要硬猜。

只输出 JSON（不要输出其他文字）：
{"verdict": "是" 或 "否" 或 "拿不准", "confidence": 0到1的小数, "reason": "一句话理由（不超过30字）"}"""


def _llm_cache_load():
    if not os.path.exists(LLM_CACHE_FILE):
        return {}
    try:
        with open(LLM_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (ValueError, OSError):
        return {}


def _llm_cache_save(data):
    tmp = LLM_CACHE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, LLM_CACHE_FILE)


def llm_verify(eid1, eid2):
    """LLM 裁决两条目是否同品（结果按 eid 对缓存；fake 模式返回拿不准）。
    返回 {"verdict","confidence","reason","cached"}；失败抛 RuntimeError。"""
    pool = load_pool()
    e1 = next((x for x in pool["entries"] if x["eid"] == eid1), None)
    e2 = next((x for x in pool["entries"] if x["eid"] == eid2), None)
    if e1 is None or e2 is None:
        raise RuntimeError("条目不存在")

    key = "|".join(sorted([eid1, eid2]))
    cache = _llm_cache_load()
    if key in cache:
        out = dict(cache[key])
        out["cached"] = True
        return out

    def _sku_txt(e):
        specs = "；".join(filter(None, [s.get("规格") for s in (e.get("SKU列表") or [])[:5]]))
        return (f"商品：{e.get('品名')}（品牌:{e.get('品牌') or '未知'}，类目:{e.get('类目') or '未知'}，"
                f"规格示例:{specs or '无'}）")

    from ai_client import call_llm, load_config
    msgs = [
        {"role": "system", "content": _LLM_SYSTEM},
        {"role": "user", "content": f"【{e1.get('shop')}】{_sku_txt(e1)}\n【{e2.get('shop')}】{_sku_txt(e2)}"},
    ]
    content = call_llm("pool_bind_llm", msgs, config=load_config(),
                       timeout=90, temperature=0.1, max_tokens=200, json_mode=True)
    try:
        parsed = json.loads(content)
    except (ValueError, TypeError):
        s, t = content.find("{"), content.rfind("}")
        parsed = json.loads(content[s:t + 1]) if 0 <= s < t else {}
    verdict = parsed.get("verdict")
    if verdict not in ("是", "否", "拿不准"):
        raise RuntimeError("模型输出格式异常: " + str(content)[:80])
    out = {"verdict": verdict,
           "confidence": round(float(parsed.get("confidence") or 0), 2),
           "reason": str(parsed.get("reason") or "")[:60]}
    cache[key] = out
    _llm_cache_save(cache)
    out["cached"] = False
    return out


# ---------------------------------------------------------------------------
# 概览（池页渲染）
# ---------------------------------------------------------------------------

def pool_overview():
    """池页数据：条目（按入池时间倒序，含活跃候选数）+ 绑定组（含成员展开）
    + 高置信待确认对 + 统计。"""
    pool = load_pool()
    entries = sorted(pool["entries"], key=lambda e: e.get("入池时间", ""), reverse=True)
    eidx = {e["eid"]: e for e in pool["entries"]}
    # 活跃候选数：对方仍在池且未绑定（对方已绑定的推荐不算活跃）
    for e in entries:
        n = 0
        for s in (e.get("推荐") or []):
            t = eidx.get(s.get("eid"))
            if t is not None and not t.get("组") and float(s.get("score") or 0) > 0:
                n += 1
        e["候选数"] = n
    groups = []
    for g in pool["groups"]:
        members = []
        for eid in g.get("eids", []):
            e = eidx.get(eid)
            if e is not None:
                members.append({k: e.get(k, "") for k in
                                ("eid", "shop", "品名", "品牌", "价格", "原价", "月售", "图片", "规格数", "SKU列表")})
        groups.append({
            "gid": g["gid"], "名称": g.get("名称", ""), "own_gid": g.get("own_gid", ""),
            "own_name": g.get("own_name", ""), "创建时间": g.get("创建时间", ""),
            "备注": g.get("备注", ""), "members": members,
        })
    pending = pending_pairs()
    stats = {
        "total": len(pool["entries"]),
        "shops": len({e.get("shop") for e in pool["entries"] if e.get("shop")}),
        "groups": len(pool["groups"]),
        "ungrouped": sum(1 for e in pool["entries"] if not e.get("组")),
        "linked_own": sum(1 for g in pool["groups"] if g.get("own_gid")),
        "pending": len(pending),
    }
    return {"entries": entries, "groups": groups, "stats": stats, "pending": pending}

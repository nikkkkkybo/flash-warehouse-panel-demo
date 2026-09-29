# -*- coding: utf-8 -*-
"""
② 店铺分区 + 区热销榜
====================================
区 = 视图配置（不复制数据）：{区名, 自家店[], 竞店[]}。
区热销榜 = 该区勾选竞店的品级销量榜（top500 缓存），
跨店同品合并（P0.2 自动化）优先级：人工绑定组 > 同条码 > 品名归一（并查集），
未绑定的品靠条码/归一名自动合并，绑定组仅作人工纠偏。

口径（回应设计确认的三个坑）：
- 销量不可比 → 双指标：绝对月售 + 店内分位（该品在本店 top500 中的排名百分位）；
- 同品多规格 → 品级聚合（top500 本身就是 comp_id 品级）；
- 各店数据时间不同步 → 每店快照时间明示（取价格观测历史该店最新快照 t）；
- 价格打架 → 展示区间：最低价-最高价，逐店 chip 明细。

存储：output/店铺分区.json
"""
import json
import os
import threading
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZONE_FILE = os.path.join(BASE, "output", "店铺分区.json")

_LOCK = threading.RLock()   # 可重入：save_zone/delete_zone 持锁内还会调 load_zones


def _now():
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def load_zones():
    """[{name, own_stores[], comp_shops[], created}]；文件不存在/坏 → []。"""
    with _LOCK:
        if not os.path.exists(ZONE_FILE):
            return []
        try:
            with open(ZONE_FILE, encoding="utf-8") as f:
                data = json.load(f)
            return data.get("zones", []) if isinstance(data, dict) else []
        except (ValueError, OSError):
            return []


def save_zone(name, own_stores, comp_shops):
    """新增/更新区（按 name 幂等）。name 空/竞店空 → None。

    独占归属：一个竞店只能属于一个区（加入本区自动从其他区移出）。
    自家店不收纳（own_stores 仅作区内默认选中，可多区共享）。
    """
    name = (name or "").strip()
    comp_shops = [s for s in (comp_shops or []) if s]
    own_stores = [s for s in (own_stores or []) if s]
    if not name or not comp_shops:
        return None
    with _LOCK:
        zones = load_zones()
        z = next((x for x in zones if x.get("name") == name), None)
        if z is None:
            z = {"name": name, "created": _now()}
            zones.append(z)
        z["own_stores"] = own_stores
        z["comp_shops"] = comp_shops
        z["updated"] = _now()
        # 独占：本区竞店从其他区移出（搬家语义）
        for other in zones:
            if other is z:
                continue
            taken = [s for s in other.get("comp_shops", []) if s not in comp_shops]
            if taken != other.get("comp_shops", []):
                other["comp_shops"] = taken
                other["updated"] = _now()
        # 空竞店的区清空后仍保留（避免误删），由用户显式删区
        os.makedirs(os.path.dirname(ZONE_FILE), exist_ok=True)
        tmp = ZONE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"zones": zones}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, ZONE_FILE)
        return z


def delete_zone(name):
    """删区 = 释放：区内竞店自动回到首页单店列表（未分组）。"""
    with _LOCK:
        zones = load_zones()
        rest = [z for z in zones if z.get("name") != name]
        if len(rest) == len(zones):
            return False
        tmp = ZONE_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"zones": rest}, f, ensure_ascii=False, indent=2)
        os.replace(tmp, ZONE_FILE)
        return True


# ---------------------------------------------------------------------------
# 区热销榜（P0.2：跨店同品自动合并 = 绑定组 > 条码 > 品名归一，并查集）
# ---------------------------------------------------------------------------

def _num(v):
    """数值解析：委托全站唯一口径 `sales_text.num`（本函数原先是 float(去逗号)，
    遇到 "月售35" / "1.5k" 会静默返回 0，只是上游目前给的是数字才没出事）。"""
    import sales_text
    return sales_text.num(v)


def _norm_merge_name(s):
    """跨店同品合并键：品名归一。
    去【】营销块、尾部打包单位（/个 /卷）、空白与标点、转小写；
    **保留数字规格**（防 5号/7号电池、M8/M10螺丝 这类不同规格品被过合并）。"""
    import re
    s = str(s or "").strip()
    s = re.sub(r"【[^】]*】", "", s)          # 【规格可选】【品牌随机】等营销块
    s = re.sub(r"/[^/]{1,6}$", "", s)        # 尾部打包单位：五金配件/个 → 五金配件
    s = re.sub(r"[\s（）()\[\]【】,，。.、;；:：\-—_*#|！!？?～~\"']+", "", s)
    return s.lower()


def _norm_barcode(v):
    """条码归一：纯数字 6-14 位才参与合并（13 位 EAN / 短码店内码均可试碰）。"""
    import re
    s = str(v or "").strip()
    return s if re.fullmatch(r"\d{6,14}", s) else ""


def build_zone_rank(zone, comp_rank_getter, include_shops=None, top_n=200):
    """构建区热销榜。

    zone            : 区 dict（comp_shops 参与店；name="__all__"=全部竞店虚拟区）
    comp_rank_getter: callable(shop) -> list[品级dict]（web_panel 传入 top500 缓存读取器，
                      品级 dict 需含 comp_id/品名/品牌/类目/月售/价格/原价/图片/规格数/SKU列表)
    include_shops   : 页面临时勾选的店（None=区全部店）
    返回 {"items": [...], "shops": [{shop, 品数}]}

    跨店同品合并（P0.2）：人工候选池绑定组（最高优先）> 同条码 > 品名归一。
    合并根一致才聚合为同一行；绑定组仅作人工纠偏覆盖，未绑定的靠自动合并。
    """
    import candidate_pool as cp

    shops = [s for s in (include_shops or zone.get("comp_shops", []))]
    if not shops:
        return {"items": [], "shops": []}

    # ---- 并查集 ----
    parent = {}

    def _find(x):
        parent.setdefault(x, x)
        r = x
        while parent[r] != r:
            r = parent[r]
        while parent[x] != r:            # 路径压缩
            parent[x], x = r, parent[x]
        return r

    def _union(a, b):
        ra, rb = _find(a), _find(b)
        if ra != rb:
            parent[rb] = ra

    # 1) 人工绑定组（最高优先）
    gid_map = {}      # (shop, comp_id) -> gid
    gid_name = {}     # gid / (gid,"own"/"own_name") -> 值
    try:
        pool = cp.load_pool()
        for g in pool.get("groups", []):
            if len(g.get("eids", [])) < 2:
                continue
            gid = g["gid"]
            gid_name[gid] = g.get("名称", "")
            gid_name[(gid, "own")] = g.get("own_gid", "")
            gid_name[(gid, "own_name")] = g.get("own_name", "")
            bg_root = ("bg", gid)
            for eid in g.get("eids", []):
                e = next((x for x in pool["entries"] if x["eid"] == eid), None)
                if e is None:
                    continue
                node = (e.get("shop"), str(e.get("comp_id")))
                gid_map[node] = gid
                _union(bg_root, node)
    except Exception:
        pass

    # 2) 收集各店品级行
    all_prods = []
    for s in shops:
        prods = comp_rank_getter(s) or []
        if prods:
            all_prods.append((s, prods))

    # 3) 自动合并并查：同条码（任一 SKU 共享）+ 品名归一
    barcode_node = {}   # 条码 -> 首见 node
    name_node = {}      # 归一品名 -> 首见 node
    for s, prods in all_prods:
        for p in prods:
            node = (s, str(p.get("comp_id")))
            nm = _norm_merge_name(p.get("品名"))
            if nm:
                if nm in name_node:
                    _union(name_node[nm], node)
                else:
                    name_node[nm] = node
            for sk in (p.get("SKU列表") or []):
                bc = _norm_barcode(sk.get("条码"))
                if bc:
                    if bc in barcode_node:
                        _union(barcode_node[bc], node)
                    else:
                        barcode_node[bc] = node

    # 4) 聚合：按合并根分组建 items（店铺明细按 (店,品) 逐条保留）
    items = {}
    root_gid = {}       # 合并根 -> 该组内的人工绑定 gid（若有）
    shop_meta = []
    for s, prods in all_prods:
        total = len(prods)
        for rank_i, p in enumerate(prods):
            node = (s, str(p.get("comp_id")))
            gid = gid_map.get(node, "")
            root = _find(node)
            if gid:
                root_gid.setdefault(root, gid)
            it = items.get(root)
            if it is None:
                it = {
                    "key": root, "品名": p.get("品名", ""), "品牌": p.get("品牌", ""),
                    "类目": p.get("类目", ""), "图片": p.get("图片", ""),
                    "绑定组": "", "总月售": 0.0, "收藏合计": 0.0, "有效月售": 0.0,
                    "最低价": None, "最高价": None,
                    "店铺明细": [], "规格数": p.get("规格数", 0),
                }
                items[root] = it
            price = _num(p.get("价格"))
            # 成交记 总月售、收藏记 收藏合计（两个口径分开摆着看），
            # 排序用 有效月售——收藏照样参与，只是按 λ 折算后计入（见 sales_text）。
            if p.get("意向"):
                it["收藏合计"] += _num(p.get("月售"))
            else:
                it["总月售"] += _num(p.get("月售"))
            ev = p.get("有效销量")
            it["有效月售"] += _num(ev) if ev not in (None, "") else _num(p.get("月售"))
            if price > 0:
                it["最低价"] = price if it["最低价"] is None else min(it["最低价"], price)
                it["最高价"] = price if it["最高价"] is None else max(it["最高价"], price)
            it["店铺明细"].append({
                "shop": s, "月售": p.get("月售", 0), "价格": p.get("价格", ""),
                "意向": bool(p.get("意向")),
                "有效销量": p.get("有效销量", p.get("月售", 0)),
                "意向数": p.get("意向数", 0),
                "分位": round(rank_i / total * 100),  # 店内分位：0=本店最畅销
                "comp_id": str(p.get("comp_id") or ""),   # 供"跳到该店视图"链接
                "sku_id": str((p.get("SKU列表") or [{}])[0].get("sku_id") or ""),
            })
        shop_meta.append({"shop": s, "品数": total})

    # 汇总条目按有效月售降序（缺该字段的旧数据退回总月售）
    out = sorted(items.values(),
                 key=lambda x: -(x.get("有效月售") or x.get("总月售") or 0))[:top_n]
    # 绑定组挂自家品（该组含人工绑定时）
    for it in out:
        it["在售店数"] = len(it.get("店铺明细") or [])
        gid = root_gid.get(it.get("key"))
        if gid:
            it["绑定组"] = gid
            it["自家品"] = gid_name.get((gid, "own_name"), "")
            it["自家品组ID"] = gid_name.get((gid, "own"), "")
        else:
            it["绑定组"] = ""
            it["自家品"] = ""
            it["自家品组ID"] = ""
    return {"items": out, "shops": shop_meta}

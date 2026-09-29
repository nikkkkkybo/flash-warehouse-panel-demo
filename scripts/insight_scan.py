# -*- coding: utf-8 -*-
"""洞察探测层（T4 的"发现"环节）：纯规则 + 统计筛选，零大模型成本。

为什么不让大模型直接读表找问题：
  · 模型读的是截断后的切片（前 22-30 行），"全局"根本看不见；
  · 数字由模型复述，就有编造风险（违反"落到实处"）；
  · 每轮都花钱，还不能复算核对。
所以分工是：**本层负责发现并给出可复算凭据，大模型只负责解释与排序**（见 interpret()）。

两类探测：
  layer=全局：必须跨品/跨类目/跨店聚合才看得见（人工逐行看不出）
  layer=细节：藏在单行里的可疑（规格不可比、条码两侧都有却没配上、意向虚高）

每条洞察都必须带 refs（来源行 ID）与 metrics（数字证据 + 计算式），
前端点开能回到那些行——不然后面只能再说一遍"AI 说"。

统计口径说明：各类目/各品的分布高度右偏（少数品占大量月售），所以用
中位数 + MAD（绝对中位差）算稳健 z 分数，而不是均值 ± 标准差；
z = 0.6745 × (x − median) / MAD，|z| ≥ 3 才算离群（约等于正态 3σ）。
"""
import os
import re
import statistics

import sales_text   # 全站销量口径（含 λ：收藏按多少权重折进有效销量）

# 探测阈值集中在这里，改口径不用翻代码（也便于回归脚本断言）
TH = {
    "z_min": 3.0,             # 稳健 z 离群门槛
    "cat_min_items": 8,       # 类目级判定最少品数（样本太小不谈规律）
    "opp_share_min": 0.35,    # 机会品占该类目比例下限
    "multi_shop_min": 3,      # 跨店共卖：至少几家竞店同时在卖
    "intent_ratio_p": 95,     # 意向/月售 比值的分位门槛
    "intent_sales_min": 20,   # 且月售本身要有体量，避免个位数噪声
    "undercut_pct": 0.15,     # 有界价差 ≥15% = 被打价
    "price_floor": 1.0,       # 单价低于 1 元多为按个卖，不做价格比较
    "magnitude_skip": 5.0,    # 两侧价格数量级差 ≥5 倍 → 判计量不可比，不算价格战
    "undercut_sales": 10,     # 且竞品月售有体量
    "diverge_comp": 30,       # 对应可疑：竞品月售 ≥
    "diverge_own": 0,         # 而自家月销 =
    "per_detector": 3,        # 每个探测最多报几条（宁缺毋滥）
}


# --------------------------------------------------------------------------
# 小工具
# --------------------------------------------------------------------------
def _num(v):
    """销量/价格数值：委托给全站唯一口径 `sales_text.num`（认 万/千/k 单位）。

    这里原本自己写 `re.search(r"-?\\d+(\\.\\d+)?")`，不认单位——同一串"1.5k人想买"
    在 top50 里算 1500、在这里算 1.5，两套口径迟早对不上账。
    成交/意向的区分不在这里做，用 `sales_text.is_intent` 与 `real_sales_ge`。
    """
    import sales_text
    return sales_text.num(v)


def _cat_eff(r):
    """一行（类目/品级聚合）里的有效销量。

    优先用生成器写好的 `竞品有效月售`；旧产物没这一列时按 总月售−收藏+λ×收藏 现算，
    再缺就退回原值——探测层不该因为产物版本而算不出数。
    """
    import sales_text as _STx
    v = r.get("竞品有效月售")
    if v not in (None, ""):
        return _num(v)
    total = _num(r.get("竞品总月售"))
    want = _num(r.get("竞品收藏量"))
    return (total - want) + _STx.INTENT_WEIGHT * want


def _robust_z(vals):
    """稳健 z 分数列表（中位数 + MAD）；MAD=0 时退化为符号距离。"""
    if len(vals) < 3:
        return [0.0] * len(vals)
    med = statistics.median(vals)
    mad = statistics.median([abs(v - med) for v in vals])
    if mad <= 0:
        return [0.0 if v == med else (1e9 if v > med else -1e9) for v in vals]
    return [0.6745 * (v - med) / mad for v in vals]


def _pct(vals, p):
    if not vals:
        return 0.0
    s = sorted(vals)
    k = min(len(s) - 1, max(0, int(round((p / 100.0) * (len(s) - 1)))))
    return s[k]


def _mk(key, layer, detector, scope, subject, title, metrics, refs, severity, hint=""):
    return {"key": key, "layer": layer, "detector": detector, "scope": scope,
            "subject": subject, "title": title, "metrics": metrics, "refs": refs,
            "severity": round(float(severity), 2), "hint": hint}


# --------------------------------------------------------------------------
# 全局层：必须聚合才看得见
# --------------------------------------------------------------------------
def d_cat_suppress(shop, cdf):
    """① 结构性落后类目：竞品合计月售相对自家极端离群 + 机会品扎堆（不是单品的锅）。"""
    out = []
    rows = [r for r in cdf.to_dict("records") if _num(r.get("品数")) >= TH["cat_min_items"]]
    if len(rows) < 3:
        return out
    ratios = [_cat_eff(r) / (1.0 + _num(r.get("自家月销合计"))) for r in rows]
    zs = _robust_z(ratios)
    order = sorted(range(len(rows)), key=lambda i: -zs[i])
    for i in order[:TH["per_detector"]]:
        r, z, ratio = rows[i], zs[i], ratios[i]
        share = _num(r.get("机会品")) / max(1.0, _num(r.get("品数")))
        if z < TH["z_min"] or share < TH["opp_share_min"]:
            continue
        out.append(_mk(
            f"cat_suppress::{shop}::{r.get('类目')}", "全局", "结构性落后类目", shop,
            str(r.get("类目")),
            f"「{r.get('类目')}」被整类目压制，不是个别品的问题",
            {"该类目竞品有效月售": round(_cat_eff(r), 1),
             "该类目竞品总月售": _num(r.get("竞品总月售")),
             "其中收藏量": _num(r.get("竞品收藏量")),
             "收藏折算λ": sales_text.INTENT_WEIGHT,
             "自家月销合计": _num(r.get("自家月销合计")),
             "压制倍数": round(ratio, 1), "稳健z": round(z, 1), "品数": _num(r.get("品数")),
             "机会品占比": f"{share * 100:.0f}%",
             "算式": "有效销量=成交+λ×收藏；倍数=有效销量/(1+自家月销合计)；"
                    "z=0.6745×(倍数−类目中位)/MAD"},
            [{"kind": "类目", "ids": [str(r.get("类目"))], "shop": shop}],
            z, "整个类目补品比逐个补品省人：先定该类目要铺的规格档与价格带，再逐品对应"))
    return out


def d_price_band_gap(shop, pdf):
    """② 价位档缺口：某类目机会品里大多数都是"竞品有成交带、我家没有对应价"。"""
    out = []
    by_cat = {}
    for r in pdf.to_dict("records"):
        if str(r.get("结论")) != "机会品":
            continue
        cat = str(r.get("类目") or "未分类")
        band = r.get("竞品价格带")
        own = r.get("自家售价区间")
        comp_mid = 0.0
        if isinstance(band, (tuple, list)) and len(band) == 2 and band[0] is not None:
            comp_mid = (_num(band[0]) + _num(band[1])) / 2.0
        # "我家在这个价位带没有货"的信号：既没有对应规格、月销也为 0
        # （只看 自家售价区间 会漏——机会品也可能有售价区间但实际零销量）
        own_sold = str(own).strip().lower() not in ("", "nan", "none", "nat")
        own_has = (own_sold and _num(r.get("自家月销")) > 0) or _num(r.get("自家规格数")) > 0
        by_cat.setdefault(cat, []).append((comp_mid, own_has, r))
    cand = []
    for cat, items in by_cat.items():
        if len(items) < 5:
            continue
        priced = [it for it in items if it[0] > 0]
        if not priced:
            continue
        no_own = sum(1 for it in items if not it[1])
        share = no_own / len(items)
        if share < 0.6:
            continue
        lo = min(p[0] for p in priced)
        hi = max(p[0] for p in priced)
        cand.append((share * len(items), cat, len(items), no_own, lo, hi, items))
    cand.sort(key=lambda x: -x[0])
    for _score, cat, n, no_own, lo, hi, items in cand[:TH["per_detector"]]:
        ids = [str(it[2].get("品组ID")) for it in items if not it[1]][:12]
        names = "、".join(str(it[2].get("品名"))[:14] for it in items[:3])
        out.append(_mk(
            f"band_gap::{shop}::{cat}", "全局", "价位档缺口", shop, cat,
            f"「{cat}」{no_own}/{n} 个机会品，我家在该价位带完全没有货",
            {"机会品数": n, "我家无对应价数": no_own, "竞品价带跨度": f"{lo:.1f}~{hi:.1f} 元",
             "占比": f"{no_own / n * 100:.0f}%",
             "算式": "占比=我家无对应价机会品/该类目机会品；价带取各品竞品价带中位的最小/最大"},
            [{"kind": "品组", "ids": ids, "shop": shop}],
            no_own * 1.2, f"例：{names}——按这个价带补档，比逐个找品更快"))
    return out


def d_intent_inflated(shop, items):
    """④ 意向虚高：收藏/意向相对成交异常高，说明热度是"想要"不是"买到"——别照它的月售定价/选品。

    意向数量来自 build_top50 的 `意向数`（原始销量列里的「N人想买」）。以前读的是 `意向`
    标志位，_num(True)=1，"意向是月售的几倍"这个判断根本没有数量可用。
    """
    recs = []
    n_intent_items = 0
    for r in (items or []):
        intent = _num(r.get("意向数"))
        if intent <= 0:
            continue
        n_intent_items += 1
        # 整条记录都是意向时，`月售` 里装的就是收藏数，不能当成交基线
        sales = 0.0 if r.get("意向") else _num(r.get("月售"))
        recs.append((r, sales, intent))
    recs = [x for x in recs if x[1] >= TH["intent_sales_min"] and x[2] > 0]
    if not recs:
        _why("意向虚高", f"{shop}：{n_intent_items} 个品有收藏量，但没有一个有成交基线"
                        f"（月售≥{TH['intent_sales_min']}）——没有可比的分母，不比倍数")
        return []
    if len(recs) < 5:
        _why("意向虚高", f"{shop}：既有收藏量又有成交基线的品只有 {len(recs)} 个，"
                        f"样本太小（<5）不谈规律")
        return []
    ratios = [x[2] / (x[1] + 1.0) for x in recs]
    cut = _pct(ratios, TH["intent_ratio_p"])
    zs = _robust_z(ratios)
    ranked = sorted(range(len(recs)), key=lambda i: -zs[i])
    out = []
    for i in ranked[:TH["per_detector"]]:
        r, sales, intent = recs[i]
        ratio = ratios[i]
        if ratio < cut or zs[i] < TH["z_min"]:
            continue
        out.append(_mk(
            f"intent::{shop}::{r.get('comp_id')}", "细节", "意向虚高（热度不可当成交看）", shop,
            str(r.get("品名")),
            f"「{str(r.get('品名'))[:18]}」收藏 {intent:.0f} 人想买，是成交 {sales:.0f} 的 {ratio:.0f} 倍",
            {"成交月售": sales, "意向数(人想买)": intent, "意向/月售": round(ratio, 1),
             "稳健z": round(zs[i], 1), "该店该比值P95": round(cut, 1),
             "算式": "比值=意向数/(成交月售+1)；门槛=本店 top500 该比值的 P95 且稳健z≥3；"
                    "意向数取自销量列的「N人想买」"},
            [{"kind": "竞品品", "ids": [str(r.get("comp_id"))], "shop": shop}],
            zs[i], "该品的热度是「想要」不是「买到」：引进与定价看成交月售，别看收藏数"))
    return out


def d_align_diverge(shop, det):
    """⑤ 对应可疑·销量背离：竞品月售很大而我家同品销量为 0。

    两种解释：①对应配错了对象（弱语义层配上的），②真的没在卖。前者危害大——它会同时
    制造假"机会品"和假"领先品"。所以按「自家常品」聚合（同一个自家 SKU 只报一条，
    把配到它身上的多个竞品并成凭据），避免同一件事刷屏。

    销量口径：`N人想买`（收藏）不剔除，按 λ 折算成有效销量参与判定——
    它照样是需求信号，只是权重低于成交，要过同样的门槛就得比成交多 1/λ 倍。
    """
    import sales_text as _STx
    groups = {}
    n_from_intent = 0
    for r in det.to_dict("records"):
        raw = r.get("竞品月售")
        e = _STx.eff(raw)
        if e < TH["diverge_comp"] or _num(r.get("自家月销")) > TH["diverge_own"]:
            continue
        if _STx.is_intent(raw):
            n_from_intent += 1
        osk = str(r.get("自家SKUID") or "").strip() or str(r.get("竞品Sku_id"))
        g = groups.setdefault(osk, {"own_name": str(r.get("自家商品名") or ""), "rows": []})
        g["rows"].append((e, raw, r))
    out = []
    for osk, g in groups.items():
        g["rows"].sort(key=lambda x: -x[0])
        sold, want, total = _STx.split_blend([x[1] for x in g["rows"]])
        layers = "、".join(sorted({str(x[2].get("语义层级") or "") for x in g["rows"][:6]} - {""}))
        confs = [x[2].get("置信度") for x in g["rows"][:6] if x[2].get("置信度")]
        e0, _raw0, r0 = g["rows"][0]
        out.append(_mk(
            f"diverge::{shop}::{osk}", "细节", "对应可疑·销量背离", shop,
            g["own_name"] or str(r0.get("竞品商品名")),
            f"常品「{g['own_name'][:18]}」被 {len(g['rows'])} 个竞品SKU 当成本品"
            f"（有效销量 {total:.0f}：成交 {sold:.0f} + 收藏 {want:.0f}×{_STx.INTENT_WEIGHT:g}），"
            f"自家销量却是 0——先怀疑配错对象",
            {"对应竞品SKU数": len(g["rows"]), "有效销量合计": round(total, 1),
             "成交销量合计": sold, "收藏量合计": want, "收藏折算λ": _STx.INTENT_WEIGHT,
             "最高单条有效销量": round(e0, 1),
             "自家月销": 0, "匹配层": layers, "最低置信度": min(confs) if confs else "",
             "跨品牌": "、".join(sorted({str(x[2].get("跨品牌") or "") for x in g["rows"][:6]} - {""})),
             "本店靠收藏折算过门槛的行": n_from_intent,
             "算式": f"有效销量=成交+λ×收藏（λ={_STx.INTENT_WEIGHT:g}，待第二期数据回测校准）；"
                    f"筛 有效销量≥{TH['diverge_comp']} 且 自家月销≤{TH['diverge_own']}，按自家SKUID 聚合成一条"},
            [{"kind": "竞品SKU", "ids": [str(x[2].get("竞品Sku_id")) for x in g["rows"][:10]], "shop": shop},
             {"kind": "自家SKU", "ids": [osk], "shop": shop}],
            total,
            f"点开逐条核规格：{str(r0.get('竞品商品名'))[:16]} 规格「{str(r0.get('竞品规格'))[:20]}」 vs "
            f"「{str(r0.get('自家规格'))[:20]}」；配错的对应要撤销，否则机会/领先判断全跟着错"))
    out.sort(key=lambda x: -x["severity"])
    return out[:TH["per_detector"]]


def _price_gap(comp, own):
    """有界价差：(自家 − 竞品)/max(自家, 竞品) ∈ [−1, 1]，正数 = 我家更贵（被打价）。

    为什么不直接用现成的 价差率 列：它是 (竞品−自家)/自家，分母是自家价——五金里螺丝扎带
    按个卖（0.01~0.3 元）时比值能到上百，一条按包装卖的竞品会被算成"我家贵 200 倍"。
    有界差值不会被极小分母炸掉，跨品才可比。
    """
    comp, own = float(comp), float(own)
    m = max(comp, own)
    return 0.0 if m <= 0 else (own - comp) / m


def d_undercut(shop, det):
    """⑥ 被打价未反应：竞品到手价比我方便宜 15% 以上，且双方都是可比量级、对方有销量体量。"""
    out = []
    rows = []
    n_from_intent = 0
    import sales_text as _STx
    for r in det.to_dict("records"):
        cp, op = _num(r.get("竞品到手价")), _num(r.get("自家售价"))
        lo = min(cp, op)
        if lo <= 0 or cp < TH["price_floor"] or op < TH["price_floor"]:
            continue                       # 单价低到 1 元以下多为"按个卖"，与按包装卖不可比
        if max(cp, op) / min(cp, op) > TH["magnitude_skip"]:
            continue                       # 数量级差太大 → 交给 ⑦ 判计量，不算价格竞争

        gap = _price_gap(cp, op)
        if gap < TH["undercut_pct"]:
            continue
        raw = r.get("竞品月售")
        if _STx.eff(raw) < TH["undercut_sales"]:
            continue                       # 折算后仍没体量 → 不构成要应对的价格战
        if _STx.is_intent(raw):
            n_from_intent += 1
        rows.append((gap, cp, op, r))
    rows.sort(key=lambda x: -(x[0] * (1 + _STx.eff(x[3].get("竞品月售")))))
    for gap, cp, op, r in rows[:TH["per_detector"] * 2]:
        sales = _STx.eff(r.get("竞品月售"))
        raw = r.get("竞品月售")
        sold0, want0 = (_num(raw), 0.0) if not _STx.is_intent(raw) else (0.0, _num(raw))
        out.append(_mk(
            f"undercut::{shop}::{r.get('竞品Sku_id')}", "细节", "被打价未反应", shop,
            str(r.get("竞品商品名")),
            f"「{str(r.get('竞品商品名'))[:18]}」竞品 {cp:.2f} 比我家 {op:.2f} 低 {gap * 100:.0f}%，"
            f"对方有效销量 {sales:.0f}（成交 {sold0:.0f}/收藏 {want0:.0f}）",
            {"竞品到手价": cp, "自家售价": op, "我家贵出": f"{gap * 100:.0f}%",
             "竞品有效销量": round(sales, 1), "竞品成交月售": sold0, "竞品收藏量": want0,
             "收藏折算λ": _STx.INTENT_WEIGHT,
             "自家月销": _num(r.get("自家月销")), "自家采购价": _num(r.get("自家采购价")),
             "本店靠收藏折算过门槛的行": n_from_intent,
             "算式": "有界价差=(自家−竞品)/max(自家,竞品)；要求 ≥15% 且双方单价≥1元 且量级差≤5倍；"
                    f"体量看有效销量=成交+λ×收藏（λ={_STx.INTENT_WEIGHT:g}）≥{TH['undercut_sales']}"
                    f"，本轮 {n_from_intent} 行靠收藏折算过门槛"},
            [{"kind": "竞品SKU", "ids": [str(r.get("竞品Sku_id"))], "shop": shop},
             {"kind": "自家SKU", "ids": [str(r.get("自家SKUID"))], "shop": shop}],
            gap * (1 + sales),
            "先核规格是否可比；可比再定跟价或调价格锚点，不可比则对齐包装规格"))
    out.sort(key=lambda x: -x["severity"])
    return out[:TH["per_detector"]]


def d_incomparable_price(shop, det):
    """⑦ 价差是假的·计量或数据不可比：两侧价差额数量级差 ≥5 倍。

    这类行会伪装成"被打价"或"我家领先"，把人带偏——真实含义多半是两边不是同一个
    计量单位（按个 vs 按包），或者是脏数据。规格文本里的数量单位能对上就顺带列出作证据。
    M6·3.5 瘦身（决策3"数据异常类自动归档"）：
      - **按品聚合**：同品多 SKU 归一条（键=店+品名归一），refs 带全部 SKU；
      - 层改为「数据体检」——它是数据核对线索，不是行动建议；台账创建时自动归档
        （状态=ignored，不计催办，人工可恢复）。
    """
    out = []
    rows = []
    for r in det.to_dict("records"):
        cp, op = _num(r.get("竞品到手价")), _num(r.get("自家售价"))
        lo, hi = min(cp, op), max(cp, op)
        if lo <= 0 or hi / lo < TH["magnitude_skip"]:
            continue
        rows.append((hi / lo, cp, op, r))

    def _pkey(name):
        return re.sub(r"[\s/／]+", "", str(name or "")).lower()[:48]

    agg = {}
    for ratio, cp, op, r in rows:
        k = _pkey(str(r.get("竞品商品名") or ""))
        e = agg.get(k)
        if e is None:
            agg[k] = {"ratio": ratio, "cp": cp, "op": op, "r": r,
                      "skus": [str(r.get("竞品Sku_id") or "")]}
        else:
            e["skus"].append(str(r.get("竞品Sku_id") or ""))
            if ratio > e["ratio"]:
                e.update({"ratio": ratio, "cp": cp, "op": op, "r": r})

    for k, e in agg.items():
        ratio, cp, op, r = e["ratio"], e["cp"], e["op"], e["r"]
        cs, os_ = str(r.get("竞品规格") or ""), str(r.get("自家规格") or "")
        units = lambda t: set(re.findall(
            r"\d+(?:\.\d+)?\s*(?:只|个|片|粒|支|根|米|寸|型|A|V|W|kg|KG|克|g|mm|cm|包|盒|袋|套|抽|层)", t))
        a, b = units(cs), units(os_)
        who = "竞品" if cp > op else "自家"
        skus = e["skus"]
        ref_ids = skus[:5] + (["…共%d个SKU" % len(skus)] if len(skus) > 5 else [])
        out.append(_mk(
            f"incomparable::{shop}::{k}", "数据体检", "价差是假的·计量不可比", shop,
            str(r.get("竞品商品名")),
            f"「{str(r.get('竞品商品名'))[:18]}」两边价差 {ratio:.0f} 倍（{len(skus)} 个SKU归并），"
            f"{who}高得离谱——先怀疑计量单位或脏数据，不是价格战",
            {"竞品到手价": cp, "自家售价": op, "倍数差": round(ratio, 1),
             "同品SKU数": len(skus),
             "竞品规格": cs[:40], "自家规格": os_[:40],
             "规格单元交集": "无" if (a and b and not (a & b)) else "有/未标注",
             "价差率原值": str(r.get("价差率") or ""),
             "算式": "max/min ≥ 5 判不可比（价差率=(竞品−自家)/自家 在按个卖的商品上会炸到上百）"},
            [{"kind": "竞品SKU", "ids": ref_ids, "shop": shop},
             {"kind": "自家SKU", "ids": [str(r.get("自家SKUID"))], "shop": shop}],
            min(ratio, 30) * 20,
            "换算成单件价再比；自动归档为数据体检（不计催办），确需处理可在台账恢复"))
    out.sort(key=lambda x: -x["severity"])
    return out[:TH["per_detector"]]


def d_barcode_missed(shop, det, comp_lookup, own_df):
    """⑧ 两侧都有条码却没配上：最该自动命中的漏网（条码是最硬的身份），说明数据或链路有问题。"""
    out = []
    try:
        own_bc = {}
        for rec in own_df.to_dict("records"):
            bc = re.sub(r"\D", "", str(rec.get("商品条码") or ""))
            if len(bc) >= 13:
                own_bc.setdefault(bc[:13], str(rec.get("SKUID") or ""))
        if not own_bc:
            return out
        have = {str(r.get("竞品Sku_id")) for r in det.to_dict("records")}
        for sku, row in (comp_lookup or {}).items():
            if sku in have:
                continue
            bc = re.sub(r"\D", "", str(row.get("barcode") or row.get("商品条码") or row.get("UPC") or ""))
            if len(bc) >= 13 and bc[:13] in own_bc:
                name = str(row.get("name") or row.get("商品名字") or "")[:18]
                out.append(_mk(
                    f"bcmiss::{shop}::{sku}", "细节", "同条码却没配上", shop, name,
                    f"条码 {bc[:13]} 两侧都有，却没进对应关系——这是最硬的身份，先查链路不是先怀疑数据",
                    {"条码": bc[:13], "竞品Sku_id": str(sku), "自家SKUID": own_bc[bc[:13]],
                     "算式": "竞品条码[:13] ∈ 自家条码集合 且 不在已对应明细内"},
                    [{"kind": "竞品SKU", "ids": [str(sku)], "shop": shop},
                     {"kind": "自家SKU", "ids": [own_bc[bc[:13]]], "shop": shop}],
                    100, "手动补上这条对应；若能批量出现，说明条码层匹配有漏"))
            if len(out) >= TH["per_detector"]:
                break
    except Exception:
        return out
    return out


# --------------------------------------------------------------------------
# 汇总入口
# --------------------------------------------------------------------------
_coverage = {}


def coverage():
    """最近一次扫描的探测器覆盖情况——静默返回空 = 摆设，必须说清为什么。"""
    return dict(_coverage)


def _why(label, text):
    """探测器自己登记"这轮为什么没出结论"。

    scan() 的 ready 判据只能表达"数据没到位"；还有一种情况是数据到位了但不足以
    支撑这条结论（样本太小、没有成交基线），不说出来就跟"没跑"看起来一模一样。
    """
    cov = _coverage.setdefault(label, {"fired": 0, "shops": 0, "blocked": []})
    if text and text not in cov["blocked"]:
        cov["blocked"].append(text)


def scan(get_frames, shops):
    """跑全部探测。get_frames(shop) → (pdf, cdf, det, top500)，由调用方负责预热与缓存。

    返回按 severity 降序的洞察列表（全局层在前优先展示由前端按 layer 分组）。
    """
    frames = {}
    for s in shops:
        try:
            frames[s] = get_frames(s)
        except Exception as e:
            print("[洞察] 取帧失败", s, e)
    out = []
    global _coverage
    _coverage = {}
    probes = [("结构性落后类目", d_cat_suppress, lambda f: f[1],
               lambda fr: len(fr[1]) >= 3, "类目统计数据不足"),
              ("价位档缺口", d_price_band_gap, lambda f: f[0],
               lambda fr: (fr[0]["结论"] == "机会品").any(), "本店无机会品"),
              ("意向虚高", d_intent_inflated, lambda f: f[3],
               lambda fr: any(_num(r.get("意向数")) > 0 for r in (fr[3] or [])),
               "榜单里没有「N人想买」型销量数据（若榜单是旧快照，重扫一次即带 意向数）"),
              ("对应可疑·销量背离", d_align_diverge, lambda f: f[2],
               lambda fr: len(fr[2]) > 0, "无已对应明细"),
              ("被打价未反应", d_undercut, lambda f: f[2],
               lambda fr: len(fr[2]) > 0, "无已对应明细"),
              ("价差是假的·计量不可比", d_incomparable_price, lambda f: f[2],
               lambda fr: len(fr[2]) > 0, "无已对应明细")]
    for s, fr in frames.items():
        # 每个探测器单独兜异常：一个数据脏了不能把后面几个一起吞掉
        for label, fn, pick, ready, why in probes:
            cov = _coverage.setdefault(label, {"fired": 0, "shops": 0, "blocked": []})
            cov["shops"] += 1
            try:
                if not ready(fr):
                    if why not in cov["blocked"]:
                        cov["blocked"].append(why)
                    continue
                got = fn(s, pick(fr))
                cov["fired"] += len(got)
                out += got
            except Exception as e:
                print(f"[洞察] 探测 {label} 在 {s} 失败: {e}")
                if "异常" not in str(_coverage[label]["blocked"]):
                    _coverage[label]["blocked"].append("异常")
    # ③ 跨店共卖缺口（要同时看多家店的榜单，人眼更不可能看到）
    try:
        cov = _coverage.setdefault("多店共卖·我家全无", {"fired": 0, "shops": len(frames), "blocked": []})
        gaps = _cross_shop_gap(frames)
        cov["fired"] = len(gaps)
        if not gaps:
            cov["blocked"].append("无 ≥3 家竞店同名共卖且我家无对应的品（或同名归一太严）")
        out += gaps
    except Exception as e:
        print("[洞察] 跨店缺口失败:", e)
    cov = _coverage.setdefault("同条码却没配上", {"fired": 0, "shops": len(frames), "blocked": []})
    for s, (pdf, cdf, det, p500) in frames.items():
        try:
            got = d_barcode_missed(s, det, _comp_lookup_of(s), _own_df())
            cov["fired"] += len(got)
        except Exception as e:
            print("[洞察] 条码漏配探测失败", s, e)
    if not cov["fired"]:
        cov["blocked"].append("两侧都有 13 位条码且未配上的品：本期没有（多数自家条码是内码非EAN）")
    out.sort(key=lambda x: -x["severity"])
    return out


def _cross_shop_gap(frames):
    """③ 跨店共卖而我家无对应（全局层）：同一品被 ≥3 家竞店同时卖进榜，
    需求已被反复证实，比单品机会榜的确定性更高——但必须跨店聚合同名才看得见。
    """
    import zone_lib as _zl
    import sales_text as _STx
    buckets = {}
    for shop, (pdf, cdf, det, p500) in frames.items():
        for it in p500 or []:
            iv = bool(it.get("意向"))
            raw = it.get("月售")
            sold, want = (0.0, _num(raw)) if iv else (_num(raw), 0.0)
            try:
                nm = _zl._norm_merge_name(str(it.get("品名") or ""))
            except Exception:
                nm = re.sub(r"\s+", "", str(it.get("品名") or ""))
            k = nm[:18]
            if len(k) < 4:
                continue
            b = buckets.setdefault(k, {"name": str(it.get("品名")), "shops": {}})
            b["shops"][shop] = (sold, want, _STx.eff_of(raw, iv), iv,
                                _num(it.get("价格")), str(it.get("comp_id")))
    # 我家已对应的品组名（任何一家店有对应就不算"全无"）
    owned = set()
    for shop, (pdf, cdf, det, p500) in frames.items():
        try:
            for nm in det["自家商品名"].tolist():
                s = re.sub(r"\s+", "", str(nm or ""))
                if s and s.lower() != "nan":
                    owned.add(s[:8])
        except Exception:
            pass
    out = []
    for k, b in buckets.items():
        if len(b["shops"]) < TH["multi_shop_min"]:
            continue
        vs = list(b["shops"].values())
        eff_total = sum(v[2] for v in vs)
        sold_total = sum(v[0] for v in vs)
        want_total = sum(v[1] for v in vs)
        only_want = sum(1 for v in vs if v[3])
        if k[:8] in owned:
            continue
        out.append(_mk(
            f"multi::跨店::{k}", "全局", "多店共卖·我家全无", "跨店", b["name"],
            f"「{b['name'][:20]}」{len(b['shops'])} 家竞店同时上榜（有效销量 {eff_total:.0f}："
            f"成交 {sold_total:.0f} + 收藏 {want_total:.0f}×{_STx.INTENT_WEIGHT:g}），我家一件没有",
            {"在售竞店数": len(b["shops"]), "其中仅收藏无成交的店数": only_want,
             "合计有效销量": round(eff_total, 1), "合计成交": sold_total, "合计收藏": want_total,
             "收藏折算λ": _STx.INTENT_WEIGHT,
             "各店有效销量": {s[:8]: round(v[2], 1) for s, v in b["shops"].items()},
             "算式": f"同名归一后出现在 ≥{TH['multi_shop_min']} 家竞店 top500 且我家无对应；"
                    f"有效销量=成交+λ×收藏（λ={_STx.INTENT_WEIGHT:g}），收藏店不剔除但排序靠后"},
            [{"kind": "竞品品", "ids": [v[5] for v in b["shops"].values()], "shop": s}
             for s in b["shops"]],
            eff_total * len(b["shops"]),
            "多店共卖=需求已被反复证实，这条引进的确定性比单品机会榜高；"
            "若只有收藏没有成交，先小批量试，别按成交体量备货"))
    out.sort(key=lambda x: -x["severity"])
    return out[:TH["per_detector"] * 2]


def _comp_lookup_of(shop):
    import web_panel as W
    return W.competitor_lookup.get(shop) or {}


def _own_df():
    import web_panel as W
    return W.own_df

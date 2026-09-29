# -*- coding: utf-8 -*-
"""销量文本的唯一解析口径（全站共用）。

为什么要有这个文件：竞品侧的销量列是**一列两种量纲**的文本——
`月售35`（真实成交）与 `35人想买`（收藏/意向）混排在同一列里（实测某店 3089 行中 1634 行是意向）。
各处自己写解析就会分叉：`float()` 直接抛异常把整行静默丢掉（页面因此报"体检干净"）；
`re` 抠数字的版本又分不清成交与意向（把收藏当成交参与比价），且对 `1.5k`/`2万` 的处理不一致。

规则：数值一律用 `num()`（带单位），语义一律用 `is_intent()`；
参与任何"比大小"之前先折算成 **有效销量 = 成交 + λ × 收藏**（`eff` / `blend`），
并且把 成交 / 收藏 / 有效 三个数一起写进凭据，让人看得见权重。

λ 的含义与出处：收藏证明"有人想要"，不证明"有人付钱"，所以它该算、但算得轻。
0.2 是保守先验（当前只有一期观测，无法从数据里估出"收藏→成交"的转化率）；
等同一批品有第二期数据后，用 上期收藏在本期变成成交 的比例回测替换这个常数。
可用环境变量 INTENT_WEIGHT 覆盖，便于 A/B 不同权重对结论条数的影响。
"""
import os
import re

INTENT_WEIGHT = float(os.environ.get("INTENT_WEIGHT", "0.2") or 0.2)

# "1,285" / "¥12.5" / "月售2" / "8人想买" / "30天 12" / "1.5k" / "2万" / "100+" / 空/nan/-
# 模式与现网用得最广的 spu_semantic._num 保持一致，便于三处副本收敛成这一处。
_NUM_PAT = re.compile(r"(\d[\d,]*(?:\.\d+)?)\s*([万亿千kKmM])?")
_INTENT_KW = ("想要", "想买", "人想", "意向", "收藏")
_EMPTY = ("", "nan", "none", "nat", "-")
_UNIT_MULT = {"万": 10000.0, "w": 10000.0, "W": 10000.0, "m": 10000.0, "M": 10000.0,
              "千": 1000.0, "k": 1000.0, "K": 1000.0}


def is_empty(v):
    return str(v if v is not None else "").strip().lower() in _EMPTY


def num(v):
    """抠出销量/数量的数值，单位（万/千/k）就地换算。解析不出数字 → 0.0。"""
    if is_empty(v):
        return 0.0
    s = str(v).replace(",", "").strip()
    m = _NUM_PAT.search(s)
    if not m:
        return 0.0
    try:
        f = float(m.group(1).replace(",", ""))
    except (TypeError, ValueError):
        return 0.0
    unit = (m.group(2) or "").strip()
    if unit:
        f *= _UNIT_MULT.get(unit, 1.0)
    return f if f == f else 0.0


def has_num(v):
    """文本里到底有没有数字（`num()` 对 "0" 和 "无" 都返回 0，这里能区分）。"""
    if is_empty(v):
        return False
    return bool(_NUM_PAT.search(str(v)))


def is_intent(v):
    """这条销量是不是「意向/收藏」而非真实成交。"""
    s = str(v if v is not None else "")
    return any(kw in s for kw in _INTENT_KW)


def parse(v):
    """→ (数值, 是否意向)。所有需要同时知道"多少"和"是什么"的地方用这个。"""
    return num(v), is_intent(v)


def label(v):
    """给人看的形态：成交不带尾巴，意向明确标注，避免在页面/提示词里被当成销量。"""
    if is_empty(v):
        return "无销量记录"
    n, intent = parse(v)
    return f"{n:.0f}（意向/收藏，非成交）" if intent else f"{n:.0f}"


def eff(v, weight=None):
    """这一行值多少"有效销量"：成交全额，收藏按 λ 折算（不是不算，是算得轻）。"""
    w = INTENT_WEIGHT if weight is None else weight
    n = num(v)
    return n * w if is_intent(v) else n


def eff_of(value, intent, weight=None):
    """数值已解析、语义存在另一个标志位时的折算（如 top500 榜单：收藏数放在 月售 里，
    性质记在 意向 字段）。与 `eff` 同一套权重。"""
    w = INTENT_WEIGHT if weight is None else weight
    n = num(value)
    return n * w if intent else n


def blend(vs, weight=None):
    """一批销量文本 → 有效销量合计。"""
    return sum(eff(v, weight) for v in (vs or []))


def split_blend(vs, weight=None):
    """一批销量文本 → (成交合计, 收藏合计, 有效合计)。凭据里要三个都给。"""
    sold = want = 0.0
    for v in (vs or []):
        n = num(v)
        if is_intent(v):
            want += n
        else:
            sold += n
    w = INTENT_WEIGHT if weight is None else weight
    return sold, want, sold + want * w


def eff_sales_ge(v, threshold, weight=None):
    """折算后是否达门槛。收藏要过同样的门槛，就得比成交多 λ 倍的量。"""
    return eff(v, weight) >= threshold


def real_sales_ge(v, threshold):
    """是否有**真实成交**达门槛——只用于"非成交不可"的场合（如④意向虚高需要成交基线）。
    一般的比大小判定请用 `eff_sales_ge`。"""
    if is_intent(v):
        return False
    return num(v) >= threshold

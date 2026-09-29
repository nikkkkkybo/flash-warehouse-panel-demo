# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 竞品优点分析与建议模块

把 差异打分_汇总.xlsx 的打分结果转成"竞品优点清单 + 对应建议"两栏对照，
优点描述具体到可执行（含数值/关键词/具体动作）。

输入:
  output/差异打分_汇总.xlsx   主输入（30 行 42 列，7 维度打分）
  output/竞品_归一化后.xlsx   辅助：取 图片url / 标题关键词字段 / 规格档位（缺失则降级）
  output/自家_归一化后.xlsx   辅助：取 标题关键词字段 / 规格档位（缺失则降级）
输出:
  output/竞品优点分析_建议.xlsx

列: 匹配组ID / 竞品品ID / 竞品商品名 / 自家SKUID / 自家商品名 /
    优点维度 / 竞品优点描述 / 对应建议 / 建议优先级 / 备注

规则:
  - 每个匹配组、每维度至多一条优点；无优点的维度不输出
    （如价格自家更便宜 -> 不输出价格优点）
  - 优先级: 高(价格/图) / 中(标题/规格) / 低(活动/评价)
  - 一个匹配组可有多条优点（多维度），故输出行数 > 30
"""
import re
import sys
from pathlib import Path

import pandas as pd

# ============================================================
# 路径常量
# ============================================================
ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
OUTPUT_DIR = ROOT / "output"
SCORE_FILE = OUTPUT_DIR / "差异打分_汇总.xlsx"
COMP_FILE = OUTPUT_DIR / "竞品_归一化后.xlsx"
OWN_FILE = OUTPUT_DIR / "自家_归一化后.xlsx"
OUT_FILE = OUTPUT_DIR / "竞品优点分析_建议.xlsx"

# 标题覆盖用到的 5 个关键词字段（与 score_diff.py 保持一致）
TITLE_FIELDS = ["品牌（抽取）", "型号", "功率", "容量", "尺寸"]

# 维度 -> 建议优先级
DIM_PRIORITY = {
    "价格": "高",
    "图": "高",
    "标题": "中",
    "规格": "中",
    "活动": "低",
    "评价": "低",
}

OUTPUT_COLUMNS = [
    "匹配组ID", "竞品品ID", "竞品商品名", "自家SKUID", "自家商品名",
    "优点维度", "竞品优点描述", "对应建议", "建议优先级", "备注",
]


# ============================================================
# 工具函数
# ============================================================
def _clean(v):
    """None/NaN/'nan'/'None' -> ''，其余 strip 后转 str。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    """转 float；空/无法解析返回 None。"""
    if v is None:
        return None
    try:
        if isinstance(v, str):
            s = v.strip()
            if s == "" or s.lower() in ("nan", "none"):
                return None
            return float(s)
        f = float(v)
        return None if pd.isna(f) else f
    except (ValueError, TypeError):
        return None


def _fmt_money(x):
    """金额格式化：None -> '未知'；整数去 .0；否则保留 2 位。"""
    if x is None:
        return "未知"
    if abs(x - round(x)) < 0.001:
        return str(int(round(x)))
    return f"{x:.2f}"


def _spec_tokens(s):
    """拆分规格字符串为 token 集合（长度 >= 2，过滤单字噪声）。"""
    s = _clean(s)
    if not s:
        return set()
    parts = re.split(r"[/／~～\-—+＋\s,，、;；:：()（）\[\]【】\"']+", s)
    return {p for p in parts if len(p) >= 2}


# ============================================================
# 数据加载
# ============================================================
def load_summary():
    if not SCORE_FILE.exists():
        print(f"[错误] 找不到打分汇总: {SCORE_FILE}", file=sys.stderr)
        sys.exit(1)
    df = pd.read_excel(
        SCORE_FILE,
        dtype={"竞品Sku_id": str, "自家SKUID": str, "竞品商品ID": str},
    )
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].fillna("")
    print(f"[加载] 差异打分_汇总: {df.shape}")
    return df


def _build_variant_sets(df, spec_col):
    """(品牌, 型号) -> set(规格原始值)。用于规格档位差的具体列举。"""
    out = {}
    if df is None or spec_col not in df.columns:
        return out
    for _, r in df.iterrows():
        b = _clean(r.get("品牌（抽取）"))
        m = _clean(r.get("型号"))
        if b == "":
            continue
        spec = _clean(r.get(spec_col))
        out.setdefault((b, m), set()).add(spec)
    return out


def load_aux():
    """加载归一化表用于 enrichment（图片url / 标题关键词 / 规格档位）。

    缺失时返回 None，下游会自动降级（缺图建议不带 url、规格建议用计数兜底）。
    返回: (comp_index, own_index, comp_vars, own_vars)
    """
    comp = own = None
    comp_vars = own_vars = {}

    if COMP_FILE.exists():
        c = pd.read_excel(COMP_FILE, dtype={"Sku_id": str, "商品ID": str})
        for col in c.columns:
            if c[col].dtype == object:
                c[col] = c[col].fillna("")
        comp_vars = _build_variant_sets(c, "规格")
        comp = c.drop_duplicates("Sku_id", keep="first").set_index("Sku_id")
        print(f"[加载] 竞品_归一化后: {c.shape}（去重后 {len(comp)} 行）")
    else:
        print(f"[警告] 竞品归一化表不存在: {COMP_FILE}（图片url/规格档位将降级）")

    if OWN_FILE.exists():
        o = pd.read_excel(OWN_FILE, dtype={"SKUID": str, "SPUID": str})
        for col in o.columns:
            if o[col].dtype == object:
                o[col] = o[col].fillna("")
        own_vars = _build_variant_sets(o, "规格名称")
        own = o.drop_duplicates("SKUID", keep="first").set_index("SKUID")
        print(f"[加载] 自家_归一化后: {o.shape}（去重后 {len(own)} 行）")
    else:
        print(f"[警告] 自家归一化表不存在: {OWN_FILE}（标题关键词/规格档位将降级）")

    return comp, own, comp_vars, own_vars


def _lookup(df_index, key):
    """按主键取一行；命中多行取首行；未命中返回 None。"""
    if df_index is None or not key or key not in df_index.index:
        return None
    row = df_index.loc[key]
    if isinstance(row, pd.DataFrame):
        row = row.iloc[0]
    return row


# ============================================================
# 各维度优点提取（返回 dict 或 None；None 表示该维度无竞品优点）
# ============================================================
def adv_price(row):
    """价格维度：竞品到手价 < 自家售价 才输出。"""
    comp_price = _num(row.get("竞品到手价"))
    own_price = _num(row.get("自家售价"))
    if comp_price is None or own_price is None:
        return None
    if comp_price >= own_price:
        # 自家更便宜或持平 -> 无竞品价格优点
        return None

    src = _clean(row.get("竞品到手价来源")) or "到手价"
    src_label = {"神价": "神价", "折后价": "折后价", "原价": "原价"}.get(src, "到手价")
    diff = own_price - comp_price
    # "比你售价低 Y%" —— 以自家售价为基准（与任务示例 15.9/19.9/20% 一致）
    rate = diff / own_price * 100 if own_price != 0 else 0.0
    upper = comp_price * 1.05

    adv = (f"竞品{src_label}{_fmt_money(comp_price)}元比你售价"
           f"{_fmt_money(own_price)}元低{rate:.1f}%（{_fmt_money(diff)}元）")
    sug = (f"调价至{_fmt_money(comp_price)}~{_fmt_money(upper)}元区间，"
           f"或匹配竞品活动力度")
    return {"优点": adv, "建议": sug, "备注": f"竞品到手价来源={src}"}


def adv_image(row, comp_row):
    """图维度：竞品有图即输出（双方有图标注待人工，自家缺图标注缺图风险）。"""
    status = _clean(row.get("图_状态"))
    comp_has = _clean(row.get("图_竞品有图")) == "是"
    own_has = _clean(row.get("图_自家有图")) == "是"
    img_url = ""
    if comp_row is not None:
        img_url = _clean(comp_row.get("图片url"))

    if comp_has and not own_has:
        url_part = f"，参考竞品图url：{img_url}" if img_url else ""
        return {
            "优点": "竞品有主图，你缺图",
            "建议": f"补主图+详情图{url_part}",
            "备注": "缺图风险",
        }
    if comp_has and own_has:
        url_part = f"，参考竞品图url：{img_url}" if img_url else ""
        return {
            "优点": "双方均有图，图内容对比需人工",
            "建议": f"人工核对主图/详情图差异{url_part}",
            "备注": "待人工看图",
        }
    # 竞品无图（自家有图 / 双方无图）-> 无竞品优点
    return None


def adv_title(row, comp_row, own_row):
    """标题维度：竞品关键词覆盖度 > 自家（标题_覆盖差距 < 0）才输出。

    覆盖差距 = 自家覆盖数 - 竞品覆盖数（负=竞品更全）。
    具体关键词 = 竞品覆盖、自家未覆盖的字段值（品牌/型号/功率/容量/尺寸）。
    """
    gap = _num(row.get("标题_覆盖差距"))
    if gap is None or gap >= 0:
        return None  # 竞品未严格更全 -> 无优点

    missing = []
    if comp_row is not None and own_row is not None:
        c_title = _clean(comp_row.get("商品名字")).lower()
        o_title = _clean(own_row.get("商品名称")).lower()
        for f in TITLE_FIELDS:
            c_val = _clean(comp_row.get(f))
            o_val = _clean(own_row.get(f))
            if c_val and c_val.lower() in c_title:
                # 竞品标题覆盖了该字段
                if not o_val or o_val.lower() not in o_title:
                    missing.append(c_val)

    if not missing:
        return None  # 覆盖差距为负但定位不到具体词（数据缺失），不强行输出

    kw = "、".join(missing[:6])
    adv = (f"竞品标题关键词覆盖度更高（比你多{int(-gap)}个维度），"
           f"你漏了：{kw}")
    sug = f"标题补入：{kw}"
    return {"优点": adv, "建议": sug, "备注": ""}


def adv_spec(row, comp_row, own_row, comp_vars, own_vars):
    """规格维度：竞品SKU档位数 > 自家（规格_档位差 > 0）才输出。

    列出竞品有、自家没有的规格档位 token。
    """
    diff = _num(row.get("规格_档位差"))  # 竞品 - 自家
    if diff is None or diff <= 0:
        return None

    c_n = _num(row.get("规格_竞品SKU数")) or 0
    o_n = _num(row.get("规格_自家SKU数")) or 0
    adv = (f"竞品有{int(c_n)}个规格档位，你只有{int(o_n)}个"
           f"（差{int(diff)}）")

    # 列出竞品有自家没有的档位 token
    comp_tok = set()
    own_tok = set()
    if comp_row is not None:
        cb = _clean(comp_row.get("品牌（抽取）"))
        cm = _clean(comp_row.get("型号"))
        if cb:
            for s in comp_vars.get((cb, cm), set()):
                comp_tok |= _spec_tokens(s)
        # 当前匹配 SKU 自身的规格也纳入参考
        comp_tok |= _spec_tokens(comp_row.get("规格"))
    if own_row is not None:
        ob = _clean(own_row.get("品牌（抽取）"))
        om = _clean(own_row.get("型号"))
        if ob:
            for s in own_vars.get((ob, om), set()):
                own_tok |= _spec_tokens(s)

    missing_tok = {t for t in (comp_tok - own_tok) if t}
    if missing_tok:
        sug = f"考虑补入规格档位：{'、'.join(sorted(missing_tok)[:6])}"
    else:
        sug = "考虑补入规格档位（参考竞品同品牌同型号 SKU 分布）"

    note = "型号缺失" if _clean(row.get("规格_型号缺失")) == "是" else ""
    return {"优点": adv, "建议": sug, "备注": note}


def adv_activity(row):
    """活动维度：竞品有神价/折扣、自家无活动数据 才输出。"""
    shen = _num(row.get("活动_竞品神价"))
    zhe = _clean(row.get("活动_竞品折扣"))

    if shen is None and zhe == "":
        return None  # 竞品也无活动 -> 无优点

    parts = []
    if shen is not None:
        parts.append(f"神价活动（{_fmt_money(shen)}元）")
    if zhe:
        parts.append(f"折扣{zhe}折")
    adv = f"竞品有{'、'.join(parts)}，你无活动"

    if shen is not None:
        sug = f"考虑参加神价活动或设满减券（参考神价{_fmt_money(shen)}元）"
    else:
        sug = f"设折扣{zhe}折或满减券"
    return {"优点": adv, "建议": sug, "备注": ""}


def adv_review(row):
    """评价维度：竞品月销 > 自家月销 才输出（自家缺失则跳过）。"""
    c_val = _num(row.get("评价_竞品月售值"))
    o_val = _num(row.get("评价_自家月销量"))
    if c_val is None:
        return None
    if o_val is None:
        # 自家月销量缺失，无法对比 -> 不输出
        return None
    if c_val <= o_val:
        return None  # 自家销量更高或持平 -> 无优点

    diff = c_val - o_val
    adv = (f"竞品月销{int(c_val)}，你月销{int(o_val)}，"
           f"差距{int(diff)}（竞品高{diff / o_val * 100:.0f}%）" if o_val else
           f"竞品月销{int(c_val)}，你月销{int(o_val)}，差距{int(diff)}")
    sug = "排查是否价格/图/标题导致转化低"
    return {"优点": adv, "建议": sug, "备注": ""}


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 70)
    print("竞品优点分析与建议")
    print("=" * 70)

    sdf = load_summary()
    comp, own, comp_vars, own_vars = load_aux()

    out_rows = []
    for _, r in sdf.iterrows():
        gid = _clean(r.get("匹配组ID"))
        comp_pid = _clean(r.get("竞品商品ID"))
        comp_name = _clean(r.get("竞品商品名"))
        own_sid = _clean(r.get("自家SKUID"))
        own_name = _clean(r.get("自家商品名"))
        sid = _clean(r.get("竞品Sku_id"))

        comp_row = _lookup(comp, sid)
        own_row = _lookup(own, own_sid)

        extractors = [
            ("价格",   lambda: adv_price(r)),
            ("图",     lambda: adv_image(r, comp_row)),
            ("标题",   lambda: adv_title(r, comp_row, own_row)),
            ("规格",   lambda: adv_spec(r, comp_row, own_row, comp_vars, own_vars)),
            ("活动",   lambda: adv_activity(r)),
            ("评价",   lambda: adv_review(r)),
        ]
        for dim, fn in extractors:
            res = fn()
            if not res:
                continue
            out_rows.append({
                "匹配组ID": gid,
                "竞品品ID": comp_pid,
                "竞品商品名": comp_name,
                "自家SKUID": own_sid,
                "自家商品名": own_name,
                "优点维度": dim,
                "竞品优点描述": res["优点"],
                "对应建议": res["建议"],
                "建议优先级": DIM_PRIORITY[dim],
                "备注": res.get("备注", ""),
            })

    out = pd.DataFrame(out_rows, columns=OUTPUT_COLUMNS)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUT_FILE, index=False)
    print(f"\n[导出] {OUT_FILE}")
    print(f"      共 {len(out)} 行 {out.shape[1]} 列  "
          f"（来源匹配组 {sdf.shape[0]} 个）")

    # ---- 维度分布 ----
    print("\n" + "=" * 70)
    print("优点维度分布")
    print("=" * 70)
    dist = out["优点维度"].value_counts()
    for dim, cnt in dist.items():
        print(f"  {dim:<4} ({DIM_PRIORITY[dim]}): {cnt} 条")
    print(f"  合计: {len(out)} 条")
    print("\n优先级分布:")
    for p, cnt in out["建议优先级"].value_counts().items():
        print(f"  {p}: {cnt} 条")

    # ---- 抽检 3 条（从不同维度各取 1 条，展示具体性）----
    print("\n" + "=" * 70)
    print("抽检 3 条（优点描述 / 对应建议）")
    print("=" * 70)
    sample_dims = ["价格", "规格", "活动"]
    picked = 0
    for dim in sample_dims:
        hits = out.index[out["优点维度"] == dim].tolist()
        if not hits:
            continue
        idx = hits[0]
        rr = out.iloc[idx]
        picked += 1
        print(f"\n[{picked}] {rr['匹配组ID']} | 维度={rr['优点维度']} | "
              f"优先级={rr['建议优先级']}")
        print(f"    竞品: {rr['竞品商品名'][:34]}")
        print(f"    自家: {rr['自家商品名'][:34]}")
        print(f"    优点: {rr['竞品优点描述']}")
        print(f"    建议: {rr['对应建议']}")
        if rr["备注"]:
            print(f"    备注: {rr['备注']}")
        if picked >= 3:
            break
    # 若上面不足 3 条，用"图"维度补
    if picked < 3:
        for idx in out.index[out["优点维度"] == "图"].tolist()[: 3 - picked]:
            rr = out.iloc[idx]
            picked += 1
            print(f"\n[{picked}] {rr['匹配组ID']} | 维度={rr['优点维度']} | "
                  f"优先级={rr['建议优先级']}")
            print(f"    竞品: {rr['竞品商品名'][:34]}")
            print(f"    自家: {rr['自家商品名'][:34]}")
            print(f"    优点: {rr['竞品优点描述']}")
            print(f"    建议: {rr['对应建议']}")

    print("\n完成。")


if __name__ == "__main__":
    main()

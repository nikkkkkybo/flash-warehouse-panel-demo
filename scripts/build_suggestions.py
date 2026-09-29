# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 阶段5：选品决策建议清单

产出两类建议，均需人工确认后才能导出：
  Task 5.1 自家缺失品引入建议
      取竞品热销品清单（Top100 + 按类目Top20 合并去重）
      排除已在"对齐_汇总.xlsx"（已确认匹配）中的竞品品
      （保守版：仅排除对齐_汇总；确认池中"确认结果=确认"的可选排除，默认开启）
  Task 5.2 同款竞争力弱优化建议
      读"差异打分_汇总.xlsx"，筛选自家劣势项：
        - 价格_竞争力 = 劣势
        - 或 图_状态 含 "缺图风险"
        - 或 规格_档位差 > 0  （按 score_diff.py 定义：竞品SKU数-自家SKU数，正=竞品规格更丰富=自家劣势）
        - 或 标题_覆盖差距 < 0 （自家标题覆盖数少于竞品）
      注意：任务描述写"规格_档位差<0"，但 score_diff.py 中该列 = 竞品SKU数 - 自家SKU数，
      正值才代表自家规格劣势。本脚本按实际数据语义（>0 为劣势）筛选，与差异打分表一致。
  Task 5.3 合并输出到 output/选品建议清单.xlsx，含确认/修改后值/人工备注空白列

输入:
  output/热销品_全表Top100.xlsx
  output/热销品_按类目Top20.xlsx
  output/对齐_汇总.xlsx
  output/对齐_人工确认池.xlsx
  output/竞品_归一化后.xlsx
  output/差异打分_汇总.xlsx
输出:
  output/选品建议清单.xlsx

用法:
  python scripts/build_suggestions.py
"""
import os
import sys
from pathlib import Path

import pandas as pd

# 复用阶段4 的工具函数（_clean/_num）和阶段3 的月售解析
sys.path.insert(0, str(Path(__file__).resolve().parent))
from score_diff import _clean, _num  # noqa: E402


# ============================================================
# 路径常量
# ============================================================
ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
OUTPUT_DIR = ROOT / "output"
HOT_OVERALL_FILE = OUTPUT_DIR / "热销品_全表Top100.xlsx"
HOT_PER_CAT_FILE = OUTPUT_DIR / "热销品_按类目Top20.xlsx"
ALIGN_FILE = OUTPUT_DIR / "对齐_汇总.xlsx"
ALIGN_POOL_FILE = OUTPUT_DIR / "对齐_人工确认池.xlsx"
COMP_FILE = OUTPUT_DIR / "竞品_归一化后.xlsx"
DIFF_FILE = OUTPUT_DIR / "差异打分_汇总.xlsx"
OUT_FILE = OUTPUT_DIR / "选品建议清单.xlsx"

# 统一输出的列顺序（5.1 + 5.2 合并到一张表）
OUTPUT_COLUMNS = [
    "建议ID", "建议类型", "品身份证",
    "竞品商品ID", "竞品商品名", "竞品类目",
    "自家SKUID", "自家商品名", "匹配组ID",
    "劣势维度", "价格建议_建议售价", "图建议_补图",
    "规格建议_扩规格", "标题建议_建议标题",
    "建议引入价区间", "预估毛利",
    "竞品月售原值", "竞品热度分", "竞品到手价", "竞品到手价来源",
    "待办清单", "建议理由", "状态",
    "确认结果", "修改后值", "人工备注",
]

STATUS_PENDING = "待人工确认"


# ============================================================
# 数据加载
# ============================================================
def load_data():
    """读取各输入表，ID 列强制 str。"""
    for f in (HOT_OVERALL_FILE, HOT_PER_CAT_FILE,
              COMP_FILE, DIFF_FILE):
        if not f.exists():
            print(f"[错误] 输入文件不存在: {f}", file=sys.stderr)
            sys.exit(1)

    hot_overall = pd.read_excel(HOT_OVERALL_FILE, dtype={"商品ID": str})
    hot_per_cat = pd.read_excel(HOT_PER_CAT_FILE, dtype={"商品ID": str})
    # 已匹配集合优先读两层汇总库 SKU层（机器+人工确认行均计入）；
    # 汇总库不存在/为空时回退旧 对齐_汇总.xlsx（历史快照）
    align = None
    try:
        sys.path.insert(0, str(ROOT / "scripts"))
        from summary_library import load_library
        _, align = load_library()
        align = align.fillna("")
    except Exception:
        align = None
    if align is None or len(align) == 0:
        if not ALIGN_FILE.exists():
            print(f"[错误] 汇总库为空且输入文件不存在: {ALIGN_FILE}", file=sys.stderr)
            sys.exit(1)
        align = pd.read_excel(ALIGN_FILE, dtype={"竞品品ID": str, "竞品Sku_id": str,
                                                 "自家SKUID": str})
    comp = pd.read_excel(COMP_FILE, dtype={"商品ID": str, "Sku_id": str})
    diff = pd.read_excel(DIFF_FILE, dtype={"竞品商品ID": str, "自家SKUID": str,
                                           "竞品Sku_id": str})

    # 人工确认池（可选，不存在不报错）
    pool = None
    if ALIGN_POOL_FILE.exists():
        pool = pd.read_excel(ALIGN_POOL_FILE,
                             dtype={"竞品品ID": str, "竞品Sku_id": str,
                                    "自家SKUID": str})

    # 空值统一
    for df in (hot_overall, hot_per_cat, align, comp, diff):
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].fillna("")
    if pool is not None:
        for c in pool.columns:
            if pool[c].dtype == object:
                pool[c] = pool[c].fillna("")

    print(f"[加载] 热销Top100: {hot_overall.shape}  按类目Top20: {hot_per_cat.shape}")
    print(f"      汇总库(SKU层): {align.shape}  竞品: {comp.shape}  差异打分: {diff.shape}")
    if pool is not None:
        print(f"      人工确认池: {pool.shape}")
    return hot_overall, hot_per_cat, align, pool, comp, diff


# ============================================================
# Task 5.1 自家缺失品建议生成
# ============================================================
def build_missing_suggestions(hot_overall, hot_per_cat, align, pool, comp):
    """生成自家缺失品引入建议。

    步骤：
      1. Top100 + 按类目Top20 合并，按商品ID去重（保留首行作代表）
      2. 排除已在 对齐_汇总 中的竞品品ID
      3. 排除已在 对齐_人工确认池 中且"确认结果=确认"的竞品品ID（可选）
      4. 从竞品表关联：竞品类目、月售原值、到手价（神价优先否则折后价，取该商品ID下最低）
      5. 生成建议条目
    返回 list[dict]。
    """
    # 1. 合并去重
    hot = pd.concat([hot_overall, hot_per_cat], ignore_index=True)
    hot = hot.drop_duplicates(subset=["商品ID"], keep="first").reset_index(drop=True)
    print(f"\n[5.1] 热销品合并去重后: {len(hot)} 个商品ID")

    # 2. 已匹配的竞品品ID集合（对齐_汇总）
    matched_ids = set(align["竞品品ID"].astype(str).unique())
    print(f"      对齐_汇总中已匹配竞品品ID: {len(matched_ids)} 个")

    # 3. 确认池中已确认的（可选排除）
    confirmed_pool_ids = set()
    if pool is not None and "确认结果" in pool.columns:
        confirmed_mask = pool["确认结果"].astype(str).str.strip().isin(
            {"确认", "confirm", "yes", "y", "是"}
        )
        confirmed_pool_ids = set(pool.loc[confirmed_mask, "竞品品ID"].astype(str).unique())
        print(f"      确认池中已确认竞品品ID: {len(confirmed_pool_ids)} 个（将排除）")

    # 4. 缺失品 = 热销品 - 已匹配 - 已确认
    hot_ids = set(hot["商品ID"].astype(str))
    missing_ids = hot_ids - matched_ids - confirmed_pool_ids
    print(f"      自家缺失的热销品: {len(missing_ids)} 个")

    # 5. 从竞品表关联价格/类目/月售（按商品ID聚合，取最低到手价）
    comp["__到手价"] = comp.apply(_comp_final_price, axis=1)
    # 同一商品ID多个SKU：取到手价最低、月售原值首个非空、类目首个非空
    comp_agg = comp.groupby("商品ID").agg(
        竞品类目=("类目三级_归一", lambda s: _first_nonempty(s)),
        月售原值=("月售", lambda s: _first_nonempty(s)),
        到手价_min=("__到手价", "min"),
        到手价来源=("__到手价来源", lambda s: _first_nonempty(s)),
    ).reset_index()

    comp_idx = comp_agg.set_index("商品ID")

    # 热销品表已有 热度分/月售原值/类目，优先用热销品表自带
    hot_idx = hot.set_index("商品ID")

    rows = []
    for i, pid in enumerate(sorted(missing_ids), start=1):
        sid = f"INTRO-{i:04d}"
        hrow = hot_idx.loc[pid] if pid in hot_idx.index else pd.Series(dtype=object)
        if isinstance(hrow, pd.DataFrame):
            hrow = hrow.iloc[0]

        crow = comp_idx.loc[pid] if pid in comp_idx.index else pd.Series(dtype=object)
        if isinstance(crow, pd.DataFrame):
            crow = crow.iloc[0]

        comp_name = _clean(hrow.get("商品名字"))
        comp_cat = _clean(hrow.get("美团后台三级类")) or _clean(crow.get("竞品类目"))
        month_raw = _clean(hrow.get("月售原值")) or _clean(crow.get("月售原值"))
        hot_score = hrow.get("热度分")
        try:
            hot_score = round(float(hot_score), 4) if pd.notna(hot_score) else None
        except (ValueError, TypeError):
            hot_score = None

        final_price = _num(crow.get("到手价_min"))
        price_src = _clean(crow.get("到手价来源"))

        # 建议引入价区间：竞品价 * 1.0 ~ 1.1
        if final_price is not None and final_price > 0:
            intro_low = round(final_price * 1.0, 2)
            intro_high = round(final_price * 1.1, 2)
            intro_range = f"{intro_low}~{intro_high}"
        else:
            intro_range = ""

        # 品身份证：缺失品无自家SKUID，仅竞品商品ID
        pid_card = f"{pid}-"

        reason_parts = [f"竞品热销品（热度分={hot_score}）"]
        if month_raw:
            reason_parts.append(f"月售={month_raw}")
        if final_price is not None:
            reason_parts.append(f"到手价={final_price}({price_src})")
        reason_parts.append("自家尚未匹配到对应商品，建议引入")
        reason = "；".join(reason_parts)

        row = {
            "建议ID": sid,
            "建议类型": "引入新品",
            "品身份证": pid_card,
            "竞品商品ID": pid,
            "竞品商品名": comp_name,
            "竞品类目": comp_cat,
            "自家SKUID": "",
            "自家商品名": "",
            "匹配组ID": "",
            "劣势维度": "",
            "价格建议_建议售价": "",
            "图建议_补图": "",
            "规格建议_扩规格": "",
            "标题建议_建议标题": "",
            "建议引入价区间": intro_range,
            "预估毛利": "",
            "竞品月售原值": month_raw,
            "竞品热度分": hot_score,
            "竞品到手价": final_price,
            "竞品到手价来源": price_src,
            "待办清单": "□找货源 □找图 □拟标题 □定规格",
            "建议理由": reason,
            "状态": STATUS_PENDING,
            "确认结果": "",
            "修改后值": "",
            "人工备注": "",
        }
        rows.append(row)

    print(f"      生成引入新品建议: {len(rows)} 条")
    return rows


def _comp_final_price(row):
    """竞品行到手价：神价优先否则折后价否则原价。返回价格值。"""
    shen = _num(row.get("神价"))
    if shen is not None:
        return shen
    zhe = _num(row.get("折后价"))
    if zhe is not None:
        return zhe
    return _num(row.get("原价"))


def _comp_final_price_src(row):
    """返回到手价来源标记。"""
    shen = _num(row.get("神价"))
    if shen is not None:
        return "神价"
    zhe = _num(row.get("折后价"))
    if zhe is not None:
        return "折后价"
    yuan = _num(row.get("原价"))
    if yuan is not None:
        return "原价"
    return ""


# 把来源计算挂到 comp 上（build_missing_suggestions 里用到 __到手价来源）
def _annotate_comp_price_src(comp):
    comp["__到手价来源"] = comp.apply(_comp_final_price_src, axis=1)
    return comp


def _first_nonempty(series):
    """取序列中第一个非空字符串。"""
    for v in series:
        s = _clean(v)
        if s:
            return s
    return ""


# ============================================================
# Task 5.2 同款竞争力弱建议生成
# ============================================================
def build_weak_suggestions(diff, comp):
    """生成同款竞争力弱优化建议。

    筛选自家劣势项：
      - 价格_竞争力 = 劣势
      - 或 图_状态 含 "缺图风险"
      - 或 规格_档位差 > 0  （竞品SKU数 > 自家SKU数，自家规格少）
      - 或 标题_覆盖差距 < 0 （自家标题覆盖数 < 竞品）
    返回 list[dict]。
    """
    # 竞品表索引：按 Sku_id 取品牌/型号/功率/容量/尺寸（用于拼建议标题）
    comp_attr = comp.set_index("Sku_id")

    # 筛选劣势项
    def _is_weak(r):
        if _clean(r.get("价格_竞争力")) == "劣势":
            return True
        if "缺图风险" in _clean(r.get("图_状态")):
            return True
        sd = _num(r.get("规格_档位差"))
        if sd is not None and sd > 0:
            return True
        tc = _num(r.get("标题_覆盖差距"))
        if tc is not None and tc < 0:
            return True
        return False

    weak = diff[diff.apply(_is_weak, axis=1)].reset_index(drop=True)
    print(f"\n[5.2] 差异打分表劣势项筛选: {len(weak)}/{len(diff)} 行")

    # 统计各维度命中数
    n_price = int((_clean_col(weak, "价格_竞争力") == "劣势").sum())
    n_img = int(weak["图_状态"].astype(str).str.contains("缺图风险", na=False).sum())
    n_spec = int(weak["规格_档位差"].apply(lambda x: _num(x) is not None and _num(x) > 0).sum())
    n_title = int(weak["标题_覆盖差距"].apply(lambda x: _num(x) is not None and _num(x) < 0).sum())
    print(f"      价格劣势: {n_price}  缺图风险: {n_img}  规格少: {n_spec}  标题弱: {n_title}")

    rows = []
    for i, r in weak.iterrows():
        sid = f"OPT-{i + 1:04d}"
        gid = _clean(r.get("匹配组ID"))
        comp_pid = _clean(r.get("竞品商品ID"))
        comp_name = _clean(r.get("竞品商品名"))
        own_sku = _clean(r.get("自家SKUID"))
        own_name = _clean(r.get("自家商品名"))

        pid_card = f"{comp_pid}-{own_sku}"

        weak_dims = []
        price_sug = ""
        img_sug = ""
        spec_sug = ""
        title_sug = ""

        # 价格劣势 → 调价建议
        if _clean(r.get("价格_竞争力")) == "劣势":
            weak_dims.append("价格劣势")
            comp_price = _num(r.get("竞品到手价"))
            if comp_price is not None and comp_price > 0:
                low = round(comp_price * 1.05, 2)
                high = round(comp_price * 1.1, 2)
                price_sug = f"{low}~{high}（竞品到手价*1.05~1.1，保留5-10%毛利）"

        # 缺图风险 → 补图建议
        if "缺图风险" in _clean(r.get("图_状态")):
            weak_dims.append("缺图风险")
            img_sug = "需补主图+详情图"

        # 规格少 → 扩规格建议
        sd = _num(r.get("规格_档位差"))
        if sd is not None and sd > 0:
            weak_dims.append("规格少")
            c_n = _num(r.get("规格_竞品SKU数"))
            o_n = _num(r.get("规格_自家SKU数"))
            spec_sug = (f"竞品规格档位更丰富（竞品{c_n}个SKU vs 自家{o_n}个SKU），"
                        f"建议补齐竞品有但自家无的规格档位")

        # 标题弱 → 换标题建议
        tc = _num(r.get("标题_覆盖差距"))
        if tc is not None and tc < 0:
            weak_dims.append("标题弱")
            # 关联竞品表取品牌/型号/核心参数拼建议标题
            comp_sid = _clean(r.get("竞品Sku_id"))
            crow = comp_attr.loc[comp_sid] if comp_sid in comp_attr.index else pd.Series(dtype=object)
            if isinstance(crow, pd.DataFrame):
                crow = crow.iloc[0]
            brand = _clean(crow.get("品牌（抽取）"))
            model = _clean(crow.get("型号"))
            power = _clean(crow.get("功率"))
            capacity = _clean(crow.get("容量"))
            size = _clean(crow.get("尺寸"))
            parts = [p for p in [brand, model, power, capacity, size] if p]
            if parts:
                title_sug = "建议标题：" + " ".join(parts) + "（基于品牌+型号+核心参数）"
            else:
                title_sug = "建议基于品牌+型号+核心参数重写标题（竞品表无结构化参数，需人工补）"

        reason = f"匹配组{gid} 自家劣势维度：{'、'.join(weak_dims)}"

        row = {
            "建议ID": sid,
            "建议类型": "优化现有",
            "品身份证": pid_card,
            "竞品商品ID": comp_pid,
            "竞品商品名": comp_name,
            "竞品类目": "",
            "自家SKUID": own_sku,
            "自家商品名": own_name,
            "匹配组ID": gid,
            "劣势维度": "、".join(weak_dims),
            "价格建议_建议售价": price_sug,
            "图建议_补图": img_sug,
            "规格建议_扩规格": spec_sug,
            "标题建议_建议标题": title_sug,
            "建议引入价区间": "",
            "预估毛利": "",
            "竞品月售原值": _clean(r.get("评价_竞品月售原值")),
            "竞品热度分": "",
            "竞品到手价": _num(r.get("竞品到手价")),
            "竞品到手价来源": _clean(r.get("竞品到手价来源")),
            "待办清单": "",
            "建议理由": reason,
            "状态": STATUS_PENDING,
            "确认结果": "",
            "修改后值": "",
            "人工备注": "",
        }
        rows.append(row)

    print(f"      生成优化现有建议: {len(rows)} 条")
    return rows


def _clean_col(df, col):
    """安全地把某列转成 cleaned str 序列。"""
    return df[col].apply(lambda v: _clean(v))


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 70)
    print("阶段5 选品决策建议清单生成")
    print("=" * 70)

    hot_overall, hot_per_cat, align, pool, comp, diff = load_data()

    # 给竞品表标注到手价来源列（5.1 聚合用）
    comp = _annotate_comp_price_src(comp)

    # Task 5.1 自家缺失品建议
    missing_rows = build_missing_suggestions(
        hot_overall, hot_per_cat, align, pool, comp
    )

    # Task 5.2 同款竞争力弱建议
    weak_rows = build_weak_suggestions(diff, comp)

    # Task 5.3 合并输出
    all_rows = missing_rows + weak_rows
    out = pd.DataFrame(all_rows, columns=OUTPUT_COLUMNS)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUT_FILE, index=False)
    print(f"\n[导出] {OUT_FILE}  共 {len(out)} 行 {out.shape[1]} 列")
    print(f"       引入新品建议: {len(missing_rows)} 条")
    print(f"       优化现有建议: {len(weak_rows)} 条")

    # ---- 统计 ----
    print("\n" + "=" * 70)
    print("建议统计")
    print("=" * 70)
    print(f"建议类型分布: {out['建议类型'].value_counts().to_dict()}")
    print(f"状态分布: {out['状态'].value_counts().to_dict()}")

    # ---- 抽检：缺失品 3 条 ----
    print("\n" + "=" * 70)
    print("抽检 - 自家缺失品建议（前 3 条）")
    print("=" * 70)
    miss_df = out[out["建议类型"] == "引入新品"].head(3)
    for _, r in miss_df.iterrows():
        print(f"\n[{r['建议ID']}] 品身份证={r['品身份证']}")
        print(f"    竞品: {r['竞品商品名'][:36]}  类目={r['竞品类目']}")
        print(f"    月售={r['竞品月售原值']}  热度分={r['竞品热度分']}  "
              f"到手价={r['竞品到手价']}({r['竞品到手价来源']})")
        print(f"    建议引入价区间={r['建议引入价区间']}  待办={r['待办清单']}")
        print(f"    理由={r['建议理由'][:60]}  状态={r['状态']}")

    # ---- 抽检：竞争力弱 3 条 ----
    print("\n" + "=" * 70)
    print("抽检 - 同款竞争力弱建议（前 3 条）")
    print("=" * 70)
    weak_df = out[out["建议类型"] == "优化现有"].head(3)
    for _, r in weak_df.iterrows():
        print(f"\n[{r['建议ID']}] 匹配组={r['匹配组ID']} 品身份证={r['品身份证']}")
        print(f"    竞品: {r['竞品商品名'][:32]}  自家: {r['自家商品名'][:32]}")
        print(f"    劣势维度={r['劣势维度']}")
        if r["价格建议_建议售价"]:
            print(f"    调价建议: {r['价格建议_建议售价']}")
        if r["图建议_补图"]:
            print(f"    补图建议: {r['图建议_补图']}")
        if r["规格建议_扩规格"]:
            print(f"    扩规格建议: {r['规格建议_扩规格']}")
        if r["标题建议_建议标题"]:
            print(f"    换标题建议: {r['标题建议_建议标题']}")
        print(f"    状态={r['状态']}")

    print("\n完成。")


if __name__ == "__main__":
    main()

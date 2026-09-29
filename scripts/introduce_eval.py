# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 商品引进评估阶段

把"找不到对应品"(无对应品清单) 和 "缺失品建议"(选品建议清单中的引入新品)
合并，一次性产出引入评估清单，为每条候选给出入库前评估信息：
  - 基础字段（竞品商品ID/商品名/类目/品牌/型号/规格/月售/热度分/竞品到手价）
  - 拟标题建议（品牌+型号+核心参数：功率/容量/尺寸）
  - 换规格建议（同商品ID下所有SKU的规格档位）
  - 找图建议（标注"待人工找图"，当前仅能判有无，无法自动抓图）
  - 找货源建议（标注"待1688询价"，无法自动抓取）
  - 利润测算（标注"待货源询价后联动测算"，依赖采购价，当前无）
  - 建议引入价区间（竞品到手价*1.0~1.1）
  - 状态：待评估

输入:
  output/无对应品清单.xlsx   (Web面板产出，可选；不存在则仅用选品建议清单)
  output/选品建议清单.xlsx   (build_suggestions 产出，含"引入新品"建议)
  output/竞品_归一化后.xlsx  (关联 SKU 规格/品牌/型号/核心参数)

输出:
  output/引进评估清单.xlsx

用法:
  python scripts/introduce_eval.py
"""
import os
import sys
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
OUTPUT_DIR = ROOT / "output"
NO_MATCH_FILE = OUTPUT_DIR / "无对应品清单.xlsx"        # Web面板产出（可选）
SUGGEST_FILE = OUTPUT_DIR / "选品建议清单.xlsx"         # build_suggestions 产出
COMP_FILE = OUTPUT_DIR / "竞品_归一化后.xlsx"           # 关联 SKU 规格/品牌/型号
OUT_FILE = OUTPUT_DIR / "引进评估清单.xlsx"

OUTPUT_COLUMNS = [
    "建议ID", "来源", "竞品商品ID", "竞品商品名", "类目", "品牌", "型号", "规格",
    "月售原值", "热度分", "竞品到手价", "建议引入价区间",
    "拟标题建议", "可引入规格档位", "找图建议", "找货源建议", "利润测算",
    "状态", "备注",
]

STATUS_PENDING = "待评估"


# ============================================================
# 小工具
# ============================================================
def _clean(v):
    """空值/NaN 归一为空字符串。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    if s.lower() in ("nan", "none", ""):
        return ""
    return s


def _num(v):
    """转 float，失败/空返回 None。"""
    s = _clean(v)
    if not s:
        return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _first_nonempty(series):
    """取序列中第一个非空（清洗后）值。"""
    for v in series.tolist():
        s = _clean(v)
        if s:
            return s
    return ""


# ============================================================
# 数据加载
# ============================================================
def load_no_match():
    """读取无对应品清单（Web面板产出）。

    兼容多种列名：竞品商品ID / 竞品品ID / 商品ID / 品身份证。
    返回 set[竞品商品ID]（字符串）。文件不存在返回空集。
    品身份证形如 "123456-"，取 "-" 前部分作为竞品商品ID。
    """
    if not NO_MATCH_FILE.exists():
        print(f"[提示] 无对应品清单不存在: {NO_MATCH_FILE}")
        print("       将仅基于选品建议清单生成评估。")
        print("       （如需合并Web面板'找不到对应'标记，请先用Web面板产出该文件后重跑）")
        return set()

    try:
        df = pd.read_excel(NO_MATCH_FILE, dtype=str)
    except Exception as e:
        print(f"[警告] 读取无对应品清单失败: {e}，跳过。")
        return set()

    id_cols = ["竞品商品ID", "竞品品ID", "商品ID", "品身份证"]
    id_col = next((c for c in id_cols if c in df.columns), None)
    if id_col is None:
        print(f"[警告] 无对应品清单未找到ID列(期望其一: {id_cols})，跳过。")
        return set()

    ids = set()
    for v in df[id_col].tolist():
        s = _clean(v)
        if s:
            ids.add(s.split("-")[0])
    print(f"[加载] 无对应品清单: {NO_MATCH_FILE.name}  {len(df)} 行  "
          f"提取竞品商品ID {len(ids)} 个")
    return ids


def load_suggestions():
    """读取选品建议清单，筛选'引入新品'建议。返回 DataFrame。"""
    if not SUGGEST_FILE.exists():
        print(f"[错误] 选品建议清单不存在: {SUGGEST_FILE}", file=sys.stderr)
        print("       请先运行: python scripts/build_suggestions.py", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(SUGGEST_FILE, dtype=str)
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].fillna("")

    if "建议类型" not in df.columns:
        print(f"[错误] 选品建议清单缺少'建议类型'列: {SUGGEST_FILE}", file=sys.stderr)
        sys.exit(1)

    intro = df[df["建议类型"].astype(str).str.strip() == "引入新品"].reset_index(drop=True)
    print(f"[加载] 选品建议清单: {SUGGEST_FILE.name}  共 {len(df)} 行  "
          f"其中'引入新品' {len(intro)} 条")
    return intro


def load_comp():
    """读取竞品归一化表，按商品ID聚合。

    返回 dict[竞品商品ID -> info]，info 含:
      name/brand/model/spec/power/capacity/size/cat/month_raw/final_price/specs(list)
    代表行取首个非空；specs 收集所有SKU规格去重；
    final_price = 神价优先 否则折后价 否则原价，取该商品ID下最低。
    """
    if not COMP_FILE.exists():
        print(f"[错误] 竞品归一化表不存在: {COMP_FILE}", file=sys.stderr)
        print("       请先运行: python scripts/normalize.py", file=sys.stderr)
        sys.exit(1)

    df = pd.read_excel(COMP_FILE, dtype={"商品ID": str, "Sku_id": str})
    for c in df.columns:
        if df[c].dtype == object:
            df[c] = df[c].fillna("")

    print(f"[加载] 竞品归一化表: {COMP_FILE.name}  {df.shape}")

    agg = {}
    for pid, grp in df.groupby("商品ID"):
        pid = _clean(pid)
        if not pid:
            continue

        # 到手价：神价优先 否则折后价 否则原价
        def comp_price(row):
            shen = _num(row.get("神价"))
            if shen is not None:
                return shen
            zhe = _num(row.get("折后价"))
            if zhe is not None:
                return zhe
            return _num(row.get("原价"))

        prices = [p for p in (comp_price(r) for _, r in grp.iterrows()) if p is not None]
        final_price = min(prices) if prices else None

        # 所有SKU规格去重（保留顺序）
        specs = []
        seen = set()
        if "规格" in grp.columns:
            for v in grp["规格"].tolist():
                s = _clean(v)
                if s and s not in seen:
                    seen.add(s)
                    specs.append(s)

        agg[pid] = {
            "name": _first_nonempty(grp.get("清洗后商品名", pd.Series(dtype=object)))
                    or _first_nonempty(grp.get("商品名字", pd.Series(dtype=object))),
            "brand": _first_nonempty(grp.get("品牌（抽取）", pd.Series(dtype=object))),
            "model": _first_nonempty(grp.get("型号", pd.Series(dtype=object))),
            "spec": _first_nonempty(grp.get("规格", pd.Series(dtype=object))),
            "power": _first_nonempty(grp.get("功率", pd.Series(dtype=object))),
            "capacity": _first_nonempty(grp.get("容量", pd.Series(dtype=object))),
            "size": _first_nonempty(grp.get("尺寸", pd.Series(dtype=object))),
            "cat": _first_nonempty(grp.get("类目三级_归一", pd.Series(dtype=object)))
                   or _first_nonempty(grp.get("美团后台三级类", pd.Series(dtype=object))),
            "month_raw": _first_nonempty(grp.get("月售", pd.Series(dtype=object))),
            "final_price": final_price,
            "specs": specs,
        }
    print(f"       聚合后竞品商品ID数: {len(agg)}")
    return agg


# ============================================================
# 评估项构造
# ============================================================
def build_title(brand, model, power, capacity, size):
    """基于 品牌+型号+核心参数(功率/容量/尺寸) 拼接拟标题建议。"""
    parts = [p for p in [brand, model, power, capacity, size] if p]
    if not parts:
        return "待人工拟标题（竞品表无结构化品牌/型号/参数）"
    return " ".join(parts)


def build_price_range(final_price, suggest_range):
    """建议引入价区间：优先用选品建议清单的值，否则竞品到手价*1.0~1.1。"""
    s = _clean(suggest_range)
    if s:
        return s
    if final_price is not None and final_price > 0:
        low = round(final_price * 1.0, 2)
        high = round(final_price * 1.1, 2)
        return f"{low}~{high}"
    return ""


# 待办标注（明确告知哪些是AI做不了的，需人工）
IMG_TODO = "需补主图+详情图（待人工找图：当前仅能判有无，无法自动抓图）"
SOURCE_TODO = "待1688询价（无法自动抓取，需人工询价）"
PROFIT_TODO = "待货源询价后联动测算（依赖采购价，当前无）"


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 70)
    print("商品引进评估清单生成")
    print("=" * 70)

    # 1. 加载三源数据
    no_match_ids = load_no_match()
    intro_df = load_suggestions()
    comp_agg = load_comp()

    # 2. 合并去重（按竞品商品ID）
    #    来源标记：缺失品建议 / 无对应 / 无对应+缺失品建议
    merged = {}  # pid -> {"来源": str, "建议行": dict or None}

    # 2a. 选品建议清单（引入新品）
    for _, r in intro_df.iterrows():
        pid = _clean(r.get("竞品商品ID"))
        if not pid:
            continue
        merged[pid] = {"来源": "缺失品建议", "建议行": r.to_dict()}

    # 2b. 无对应品清单
    for pid in no_match_ids:
        if pid in merged:
            merged[pid]["来源"] = "无对应+缺失品建议"
        else:
            merged[pid] = {"来源": "无对应", "建议行": None}

    n_only_nomatch = sum(1 for v in merged.values() if v["来源"] == "无对应")
    n_only_sugg = sum(1 for v in merged.values() if v["来源"] == "缺失品建议")
    n_both = sum(1 for v in merged.values() if v["来源"] == "无对应+缺失品建议")
    print(f"\n[合并] 去重后竞品商品ID总数: {len(merged)}")
    print(f"       仅来自无对应品清单: {n_only_nomatch}")
    print(f"       仅来自缺失品建议: {n_only_sugg}")
    print(f"       两源合并: {n_both}")

    # 3. 逐条产出评估信息
    rows = []
    n_no_comp = 0
    for i, pid in enumerate(sorted(merged.keys()), start=1):
        m = merged[pid]
        src = m["来源"]
        srow = m["建议行"] or {}

        cinfo = comp_agg.get(pid, {})
        if not cinfo:
            n_no_comp += 1

        brand = cinfo.get("brand", "")
        model = cinfo.get("model", "")
        spec = cinfo.get("spec", "")
        power = cinfo.get("power", "")
        capacity = cinfo.get("capacity", "")
        size = cinfo.get("size", "")
        specs_list = cinfo.get("specs", [])

        # 基础字段：建议清单优先，缺失从竞品表补
        comp_name = _clean(srow.get("竞品商品名")) or cinfo.get("name", "")
        comp_cat = _clean(srow.get("竞品类目")) or cinfo.get("cat", "")
        month_raw = _clean(srow.get("竞品月售原值")) or cinfo.get("month_raw", "")
        hot_score = _clean(srow.get("竞品热度分"))
        comp_price = _num(srow.get("竞品到手价"))
        if comp_price is None:
            comp_price = cinfo.get("final_price")

        price_range = build_price_range(comp_price, srow.get("建议引入价区间"))
        title = build_title(brand, model, power, capacity, size)

        if specs_list:
            spec_range = "、".join(specs_list)
        elif spec:
            spec_range = spec
        else:
            spec_range = "（竞品表无规格字段，待人工确认）"

        # 备注
        remark_parts = []
        if src == "无对应":
            remark_parts.append("Web面板标记'找不到对应品'")
        elif src == "无对应+缺失品建议":
            remark_parts.append("Web面板标记'找不到对应品' + 选品建议清单'引入新品'")
        if not cinfo:
            remark_parts.append("竞品归一化表未找到该商品ID，基础信息缺失")
        remark = "；".join(remark_parts)

        rows.append({
            "建议ID": f"EVAL-{i:04d}",
            "来源": src,
            "竞品商品ID": pid,
            "竞品商品名": comp_name,
            "类目": comp_cat,
            "品牌": brand,
            "型号": model,
            "规格": spec,
            "月售原值": month_raw,
            "热度分": hot_score,
            "竞品到手价": comp_price if comp_price is not None else "",
            "建议引入价区间": price_range,
            "拟标题建议": title,
            "可引入规格档位": spec_range,
            "找图建议": IMG_TODO,
            "找货源建议": SOURCE_TODO,
            "利润测算": PROFIT_TODO,
            "状态": STATUS_PENDING,
            "备注": remark,
        })

    if n_no_comp:
        print(f"[注意] {n_no_comp} 条候选在竞品归一化表中未找到，基础信息可能缺失")

    # 4. 输出
    out = pd.DataFrame(rows, columns=OUTPUT_COLUMNS)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUT_FILE, index=False)
    print(f"\n[导出] {OUT_FILE}  共 {len(out)} 行 {out.shape[1]} 列")

    # 5. 统计 + 抽检
    print("\n" + "=" * 70)
    print("评估清单统计")
    print("=" * 70)
    print(f"来源分布: {out['来源'].value_counts().to_dict()}")
    print(f"状态分布: {out['状态'].value_counts().to_dict()}")
    has_title = int(out["拟标题建议"].astype(str).str.startswith("待人工").sum())
    has_spec = int(out["可引入规格档位"].astype(str).str.contains("、", na=False).sum())
    has_price = int(out["建议引入价区间"].astype(str).apply(lambda x: bool(x.strip())).sum())
    print(f"拟标题建议非空(含待人工标记): {len(out) - int(out['拟标题建议'].astype(str).str.startswith('待人工').sum())} 条有效 / {has_title} 条待人工")
    print(f"可引入规格档位(多档): {has_spec} 条")
    print(f"建议引入价区间已填: {has_price} 条")

    print("\n" + "=" * 70)
    print("抽检 - 引进评估清单（前 3 条）")
    print("=" * 70)
    for _, r in out.head(3).iterrows():
        print(f"\n[{r['建议ID']}] 来源={r['来源']}  竞品商品ID={r['竞品商品ID']}")
        print(f"    商品名: {str(r['竞品商品名'])[:36]}")
        print(f"    类目={r['类目']}  品牌={r['品牌']}  型号={r['型号']}  规格={r['规格']}")
        print(f"    月售={r['月售原值']}  热度分={r['热度分']}  到手价={r['竞品到手价']}")
        print(f"    建议引入价区间={r['建议引入价区间']}")
        print(f"    拟标题建议: {r['拟标题建议']}")
        print(f"    可引入规格档位: {str(r['可引入规格档位'])[:60]}")
        print(f"    找图: {r['找图建议']}")
        print(f"    找货源: {r['找货源建议']}")
        print(f"    利润测算: {r['利润测算']}")
        if r["备注"]:
            print(f"    备注: {r['备注']}")
        print(f"    状态={r['状态']}")

    print("\n完成。")


if __name__ == "__main__":
    main()

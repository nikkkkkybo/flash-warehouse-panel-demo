# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 阶段3：竞品热销品发现模块

功能：
  3.1 parse_month_sales  月售字段解析器
  3.2 calc_hot_score     热度分计算（可扩展骨架，v1 仅用月售归一化分）
  3.3 export_top_n       按美团后台三级类目取 Top20、全表取 Top100 并导出

输入: d:\\查图\\竞品_示例竞店B商品库_美团导出.xlsx
输出: output/热销品_按类目Top20.xlsx
      output/热销品_全表Top100.xlsx
"""

import os
import re
import sys
from pathlib import Path

import pandas as pd

try:
    import yaml
except ImportError:  # 极端情况下回退，但本环境已确认 pyyaml 可用
    yaml = None


# ============================================================
# 路径常量
# ============================================================
ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
INPUT_FILE = ROOT / "竞品_示例竞店B商品库_美团导出.xlsx"
CONFIG_FILE = Path(__file__).resolve().parent / "hot_score_config.yaml"
OUTPUT_DIR = ROOT / "output"
OUT_PER_CAT = OUTPUT_DIR / "热销品_按类目Top20.xlsx"
OUT_OVERALL = OUTPUT_DIR / "热销品_全表Top100.xlsx"


# ============================================================
# Task 3.1 月售解析器
# ============================================================
# 用两段独立正则分别匹配「已售」与「意向」，绝不混算。
_RE_SOLD = re.compile(r"^已售(\d+)(\+?)$")        # 已售1 / 已售100+
_RE_INTENT = re.compile(r"^(\d+)人想买$")         # 1人想买


def parse_month_sales(value):
    """解析月售字段，返回 dict: {数值:int, 类型:str, 是否区间:bool}

    规则：
      "已售1"     -> {数值:1,   类型:"已售", 是否区间:False}
      "已售100+"  -> {数值:100, 类型:"已售", 是否区间:True}   # 区间值取下限
      "1人想买"   -> {数值:1,   类型:"意向", 是否区间:False}
      空值/无法解析 -> {数值:0, 类型:"未知", 是否区间:False}
    """
    # 空值处理：None / NaN / 空串 / 字面量 nan
    if value is None:
        return {"数值": 0, "类型": "未知", "是否区间": False}
    s = str(value).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return {"数值": 0, "类型": "未知", "是否区间": False}

    # 已售（精确 / 区间）
    m = _RE_SOLD.match(s)
    if m:
        num = int(m.group(1))
        is_range = (m.group(2) == "+")
        return {"数值": num, "类型": "已售", "是否区间": is_range}

    # 意向（N人想买）
    m = _RE_INTENT.match(s)
    if m:
        num = int(m.group(1))
        return {"数值": num, "类型": "意向", "是否区间": False}

    # 兜底
    return {"数值": 0, "类型": "未知", "是否区间": False}


# ============================================================
# Task 3.2 热度分计算（可扩展骨架）
# ============================================================
def calc_hot_score(row, weights):
    """计算单行热度分，返回 0-1 归一化分。

    v1 仅启用「月售归一化分」一项，按 weights 配置加权求和后归一化。
    row 需含预处理好的「月售归一化分」字段。

    TODO: 后续加入评论数/差评关键词/上新频率
    """
    score = 0.0
    components = weights.get("components", {})

    # ---- 月售因子 ----
    ms = components.get("monthly_sales", {})
    if ms.get("enabled", True):
        w = float(ms.get("weight", 1.0))
        score += w * float(row.get("月售归一化分", 0.0))

    # TODO: 后续加入评论数/差评关键词/上新频率
    # ---- 预留：评论数 ----
    # rc = components.get("review_count", {})
    # if rc.get("enabled", False):
    #     score += float(rc.get("weight", 0.0)) * float(row.get("评论归一化分", 0.0))
    # ---- 预留：差评关键词（负向）----
    # nk = components.get("negative_keyword", {})
    # if nk.get("enabled", False):
    #     score += float(nk.get("weight", 0.0)) * float(row.get("差评关键词命中", 0.0))
    # ---- 预留：上新频率 ----
    # npf = components.get("new_product_freq", {})
    # if npf.get("enabled", False):
    #     score += float(npf.get("weight", 0.0)) * float(row.get("上新频率归一化分", 0.0))

    # 用「已启用」组件权重总和做归一化，保证最终落在 0-1
    total_w = 0.0
    for c in components.values():
        if isinstance(c, dict) and c.get("enabled", False):
            total_w += float(c.get("weight", 0.0))
    if total_w > 0:
        score = score / total_w

    # 裁剪到 [0, 1]
    score = max(0.0, min(1.0, score))
    return round(score, 4)


# ============================================================
# 配置加载 & 月售归一化分预处理
# ============================================================
def load_config():
    """读取 hot_score_config.yaml，失败时回退到内置 dict。"""
    if yaml is not None and CONFIG_FILE.exists():
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        if cfg and "hot_score" in cfg:
            return cfg
    # 回退内置配置（与 yaml 等价）
    return {
        "hot_score": {
            "version": "v1",
            "components": {
                "monthly_sales": {
                    "enabled": True,
                    "weight": 1.0,
                    "sold": {"normalize": "min_max", "scope": "global"},
                    "intent": {"normalize": "min_max", "weight": 0.3},
                }
            },
        },
        "export": {"per_category_top_n": 20, "overall_top_n": 100, "rank_method": "first"},
    }


def _min_max(series):
    """对一组数值做 min-max 归一化到 0-1，全相等时返回全 0。"""
    s_min, s_max = float(series.min()), float(series.max())
    if s_max <= s_min:
        return pd.Series(0.0, index=series.index)
    return (series.astype(float) - s_min) / (s_max - s_min)


def add_monthly_sales_score(df, hot_cfg):
    """解析月售并计算「月售归一化分」列。

    - 已售类型：按数值在全表做 min-max 归一化（区间值取下限）
    - 意向类型：按数值在全表做 min-max 归一化后 × intent_weight（0.3）
    - 未知类型：0 分
    """
    components = hot_cfg.get("components", {})
    ms_cfg = components.get("monthly_sales", {})
    intent_w = float(ms_cfg.get("intent", {}).get("weight", 0.3))

    # 解析
    parsed = df["月售"].apply(parse_month_sales)
    df["月售解析值"] = parsed.apply(lambda x: x["数值"]).astype(int)
    df["月售类型"] = parsed.apply(lambda x: x["类型"])
    df["是否区间"] = parsed.apply(lambda x: x["是否区间"])

    # 归一化分初始化
    df["月售归一化分"] = 0.0

    sold_mask = df["月售类型"] == "已售"
    intent_mask = df["月售类型"] == "意向"

    if sold_mask.any():
        df.loc[sold_mask, "月售归一化分"] = _min_max(df.loc[sold_mask, "月售解析值"])

    if intent_mask.any():
        df.loc[intent_mask, "月售归一化分"] = _min_max(df.loc[intent_mask, "月售解析值"]) * intent_w

    df["月售归一化分"] = df["月售归一化分"].fillna(0.0).clip(0.0, 1.0)
    return df


# ============================================================
# Task 3.3 Top N 导出
# ============================================================
OUTPUT_COLUMNS = [
    "商品ID", "商品名字", "一级类", "二级类", "美团后台三级类",
    "月售原值", "月售解析值", "月售类型", "热度分",
    "类目内排名", "全表排名",
]


def export_top_n(df, export_cfg):
    """按美团后台三级类目取 Top20，全表取 Top100，导出两个 Excel。

    排名规则：
      - 先按 热度分 desc, 月售解析值 desc, 商品ID asc 排序（确定性）
      - 全表排名：排序后的全局序号
      - 类目内排名：排序后每个类目内的序号
    """
    per_cat_n = int(export_cfg.get("per_category_top_n", 20))
    overall_n = int(export_cfg.get("overall_top_n", 100))

    # 确定性排序
    df_sorted = df.sort_values(
        by=["热度分", "月售解析值", "商品ID"],
        ascending=[False, False, True],
        kind="mergesort",
    ).reset_index(drop=True)

    # 全表排名
    df_sorted["全表排名"] = df_sorted.index + 1
    # 类目内排名（已按热度分降序，cumcount 即类目内名次）
    df_sorted["类目内排名"] = df_sorted.groupby("美团后台三级类").cumcount() + 1

    # 月售原值
    df_sorted["月售原值"] = df_sorted["月售"]

    # 按类目 Top20
    top_per_cat = df_sorted[df_sorted["类目内排名"] <= per_cat_n].copy()
    top_per_cat = top_per_cat.sort_values(
        by=["美团后台三级类", "类目内排名"], ascending=[True, True]
    ).reset_index(drop=True)

    # 全表 Top100
    top_overall = df_sorted.head(overall_n).copy()
    top_overall = top_overall.sort_values(by=["全表排名"], ascending=True).reset_index(drop=True)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    top_per_cat[OUTPUT_COLUMNS].to_excel(OUT_PER_CAT, index=False)
    top_overall[OUTPUT_COLUMNS].to_excel(OUT_OVERALL, index=False)

    return top_per_cat, top_overall


# ============================================================
# 主流程
# ============================================================
def main():
    # 0. 校验输入
    if not INPUT_FILE.exists():
        print(f"[错误] 输入文件不存在: {INPUT_FILE}", file=sys.stderr)
        sys.exit(1)

    cfg = load_config()
    hot_cfg = cfg.get("hot_score", {})
    export_cfg = cfg.get("export", {})

    # 1. 读取竞品表
    print(f"[1] 读取竞品表: {INPUT_FILE.name}")
    df = pd.read_excel(INPUT_FILE, dtype={"商品ID": str})
    print(f"    原始行数: {len(df)}, 列数: {df.shape[1]}")

    # 2. 月售解析 + 归一化分
    print("[2] 解析月售字段并计算月售归一化分 ...")
    df = add_monthly_sales_score(df, hot_cfg)
    type_counts = df["月售类型"].value_counts().to_dict()
    print(f"    类型分布: {type_counts}")

    # 3. 热度分
    print("[3] 计算热度分 (v1: 月售归一化分) ...")
    df["热度分"] = df.apply(lambda r: calc_hot_score(r, hot_cfg), axis=1)

    # 4. 导出 Top N
    print("[4] 导出 Top N 清单 ...")
    top_per_cat, top_overall = export_top_n(df, export_cfg)
    print(f"    按类目Top20: {OUT_PER_CAT.name}  共 {len(top_per_cat)} 行, "
          f"覆盖 {top_per_cat['美团后台三级类'].nunique()} 个类目")
    print(f"    全表Top100: {OUT_OVERALL.name}  共 {len(top_overall)} 行")

    # 5. 抽检
    print("\n[5] 抽检 - 月售解析:")
    for sample in ["已售1", "已售100+", "1人想买", "", None, "已售23", "2人想买"]:
        print(f"    {sample!r:>12} -> {parse_month_sales(sample)}")

    print("\n[5] 抽检 - 热度分对比:")
    s100 = df[df["月售"] == "已售100+"]
    s1 = df[df["月售"] == "已售1"]
    if len(s100):
        print(f"    已售100+ 热度分(均值): {s100['热度分'].mean():.4f}")
    if len(s1):
        print(f"    已售1    热度分(均值): {s1['热度分'].mean():.4f}")

    print("\n[5] Top5 热销品 (全表):")
    show_cols = ["全表排名", "商品名字", "美团后台三级类", "月售原值", "热度分"]
    print(top_overall[show_cols].head(5).to_string(index=False))

    print(f"\n输出目录: {OUTPUT_DIR}")
    print("完成。")


if __name__ == "__main__":
    main()

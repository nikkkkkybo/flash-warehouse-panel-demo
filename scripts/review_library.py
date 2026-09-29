# -*- coding: utf-8 -*-
"""
阶段7 Task 7.3 复盘库脚本

职责：
  1. 读取回看对比表 output/回看对比_YYYYMMDD.xlsx（默认取最新一份）。
  2. 筛选 |偏差率| 超阈值（默认 0.50，读自 scoring_weights.yaml）的记录进复盘库。
  3. 基于偏差方向反推打分维度权重调整「建议」：
       实际高于预期 (偏差率>0) -> 正向维度权重可能偏低，建议上调
       实际低于预期 (偏差率<0) -> 正向维度权重可能偏高，建议下调
  4. 写出 output/复盘库.xlsx：
       sheet 「复盘条目」   : 逐条高偏差记录 + 建议调整方向
       sheet 「权重调整建议」: 按维度汇总的建议（需人工确认后手动改 scoring_weights.yaml 才生效）
  5. 权重调整建议「绝不自动改写」scoring_weights.yaml，仅产出建议表。

注意：
  - 当前无真实上架后数据，回看对比表可能为空 -> 复盘库也为空，脚本仍可跑通。
  - 因阶段4/5未产出每条决策的逐维度得分，此处建议为「方向性建议」：
    对全部正向维度统一给出方向；后续接入逐维度得分后可细化为按维度归因。
"""
import os
import glob
import argparse
import pandas as pd
import numpy as np

try:
    import yaml
except ImportError:
    yaml = None

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
WEIGHTS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "scoring_weights.yaml")
REVIEW_GLOB = os.path.join(OUT_DIR, "回看对比_*.xlsx")
LIBRARY_PATH = os.path.join(OUT_DIR, "复盘库.xlsx")

DEFAULT_THRESHOLD = 0.50
DEFAULT_ADJUST_CAP = 0.20

# 内置回退权重（与 scoring_weights.yaml 等价；无 pyyaml 或读文件失败时用）
DEFAULT_WEIGHTS = {
    "scoring": {
        "version": "v1-placeholder",
        "dimensions": {
            "热度分":       {"weight": 0.30, "direction": "positive", "desc": "竞品热度分"},
            "排名优势":     {"weight": 0.20, "direction": "positive", "desc": "竞品排名"},
            "利润空间":     {"weight": 0.25, "direction": "positive", "desc": "毛利空间"},
            "类目匹配度":   {"weight": 0.10, "direction": "positive", "desc": "类目匹配"},
            "销量预期":     {"weight": 0.10, "direction": "positive", "desc": "预期销量"},
            "供应链稳定度": {"weight": 0.05, "direction": "positive", "desc": "供应链"},
        },
        "normalize": True,
    },
    "lookback": {
        "review_days": [7, 14, 30],
        "deviation_threshold": DEFAULT_THRESHOLD,
        "suggested_adjust_cap": DEFAULT_ADJUST_CAP,
    },
}


def load_weights():
    """读取 scoring_weights.yaml，失败回退内置 dict。返回 cfg dict。"""
    if yaml is not None and os.path.exists(WEIGHTS_FILE):
        with open(WEIGHTS_FILE, "r", encoding="utf-8") as f:
            cfg = yaml.safe_load(f)
        if cfg and "scoring" in cfg and "lookback" in cfg:
            return cfg
    return DEFAULT_WEIGHTS


def find_latest_review():
    """在 output/ 下找最新的 回看对比_YYYYMMDD.xlsx。"""
    files = glob.glob(REVIEW_GLOB)
    if not files:
        return ""
    # 按文件名里的日期字符串排序取最新
    files.sort(key=lambda p: os.path.basename(p), reverse=True)
    return files[0]


# 复盘条目列
ENTRY_COLUMNS = [
    "决策ID", "品身份证", "决策日期", "上架日期", "回看周期天",
    "预期销量", "实际销量", "偏差", "偏差率", "偏差方向",
    "建议调整维度", "建议调整方向", "建议调整幅度(相对)",
    "反推说明", "是否已确认生效",
]

# 权重调整建议列
WEIGHT_SUGGEST_COLUMNS = [
    "维度", "当前权重", "方向", "建议调整方向", "建议调整幅度(相对)",
    "建议新权重", "累计样本数", "反推依据", "是否已确认生效",
]


def _direction(rate):
    """偏差率 -> 方向文字。"""
    if rate is None or (isinstance(rate, float) and pd.isna(rate)):
        return ""
    if rate > 0:
        return "实际高于预期"
    if rate < 0:
        return "实际低于预期"
    return "持平"


def build_entries(detail, threshold, adjust_cap, dims):
    """
    由回看明细筛选高偏差记录，构造复盘条目。
    返回 (entries_df, pos_count, neg_count, mean_rate)。
    """
    if detail is None or len(detail) == 0:
        return pd.DataFrame(columns=ENTRY_COLUMNS), 0, 0, np.nan

    d = detail.copy()
    d["偏差率"] = pd.to_numeric(d["偏差率"], errors="coerce")
    d = d.dropna(subset=["偏差率"])
    mask = d["偏差率"].abs() > threshold
    high = d[mask].copy()
    if len(high) == 0:
        return pd.DataFrame(columns=ENTRY_COLUMNS), 0, 0, np.nan

    pos = int((high["偏差率"] > 0).sum())
    neg = int((high["偏差率"] < 0).sum())
    mean_rate = float(high["偏差率"].mean())

    rows = []
    # 建议调整维度：当前无逐维度得分，统一列出全部正向维度
    pos_dims = [k for k, v in dims.items() if v.get("direction", "positive") == "positive"]
    dims_str = "、".join(pos_dims)

    for _, r in high.iterrows():
        rate = float(r["偏差率"])
        direct = _direction(rate)
        # 建议方向：实际高于预期 -> 上调正向维度；反之下调
        if rate > 0:
            adj_dir = "上调"
            adj_mag = min(adjust_cap, abs(rate) * 0.5)
            why = ("实际销量高于预期，说明该品真实需求被低估，"
                   "正向打分维度权重可能偏低，建议上调后复算。")
        else:
            adj_dir = "下调"
            adj_mag = min(adjust_cap, abs(rate) * 0.5)
            why = ("实际销量低于预期，说明该品真实需求被高估，"
                   "正向打分维度权重可能偏高，建议下调后复算。")
        rows.append({
            "决策ID": r.get("决策ID", ""),
            "品身份证": r.get("品身份证", ""),
            "决策日期": r.get("决策日期", ""),
            "上架日期": r.get("上架日期", ""),
            "回看周期天": r.get("回看周期天", ""),
            "预期销量": r.get("预期销量", ""),
            "实际销量": r.get("实际销量", ""),
            "偏差": r.get("偏差", ""),
            "偏差率": rate,
            "偏差方向": direct,
            "建议调整维度": dims_str,
            "建议调整方向": adj_dir,
            "建议调整幅度(相对)": round(adj_mag, 4),
            "反推说明": why,
            "是否已确认生效": "否",
        })
    return pd.DataFrame(rows, columns=ENTRY_COLUMNS), pos, neg, mean_rate


def build_weight_suggestions(entries, dims, adjust_cap, pos, neg, mean_rate):
    """
    按维度汇总建议。无样本时返回空表(仅表头)。
    """
    if entries is None or len(entries) == 0:
        return pd.DataFrame(columns=WEIGHT_SUGGEST_COLUMNS)

    # 汇总方向：以高偏差样本的整体偏差率均值定基调
    # mean_rate>0 -> 偏向低估 -> 上调；<0 -> 偏向高估 -> 下调
    if mean_rate >= 0:
        adj_dir = "上调"
        basis = (f"高偏差样本{len(entries)}条中实际高于预期{pos}条/低于预期{neg}条，"
                 f"平均偏差率{mean_rate:.2%}，整体倾向低估，建议上调正向维度。")
    else:
        adj_dir = "下调"
        basis = (f"高偏差样本{len(entries)}条中实际高于预期{pos}条/低于预期{neg}条，"
                 f"平均偏差率{mean_rate:.2%}，整体倾向高估，建议下调正向维度。")

    rows = []
    for name, meta in dims.items():
        cur_w = float(meta.get("weight", 0))
        direct = meta.get("direction", "positive")
        if direct != "positive":
            # 负向维度反向：实际高于预期应下调负向维度（当前无负向维度，预留）
            d = "下调" if adj_dir == "上调" else "上调"
        else:
            d = adj_dir
        mag = min(adjust_cap, abs(mean_rate) * 0.5) if not pd.isna(mean_rate) else 0.0
        new_w = cur_w * (1 + mag) if d == "上调" else cur_w * (1 - mag)
        rows.append({
            "维度": name,
            "当前权重": cur_w,
            "方向": direct,
            "建议调整方向": d,
            "建议调整幅度(相对)": round(mag, 4),
            "建议新权重": round(new_w, 4),
            "累计样本数": len(entries),
            "反推依据": basis,
            "是否已确认生效": "否",
        })
    return pd.DataFrame(rows, columns=WEIGHT_SUGGEST_COLUMNS)


def run_library(review_path=None, library_path=LIBRARY_PATH):
    """执行复盘库生成。返回 (entries_df, weight_suggest_df, out_path)。"""
    cfg = load_weights()
    dims = cfg["scoring"].get("dimensions", {})
    threshold = float(cfg.get("lookback", {}).get("deviation_threshold", DEFAULT_THRESHOLD))
    adjust_cap = float(cfg.get("lookback", {}).get("suggested_adjust_cap", DEFAULT_ADJUST_CAP))

    review_path = review_path or find_latest_review()
    if not review_path or not os.path.exists(review_path):
        print("未找到回看对比表(output/回看对比_*.xlsx)。")
        print("请先运行: python review_lookback.py")
        print("当前无回看数据，复盘库输出空表(占位)。")
        entries = pd.DataFrame(columns=ENTRY_COLUMNS)
        weights = pd.DataFrame(columns=WEIGHT_SUGGEST_COLUMNS)
    else:
        print(f"读取回看对比表: {review_path}")
        detail = pd.read_excel(review_path, sheet_name="回看对比", dtype=str)
        entries, pos, neg, mean_rate = build_entries(detail, threshold, adjust_cap, dims)
        weights = build_weight_suggestions(entries, dims, adjust_cap, pos, neg, mean_rate)

    os.makedirs(OUT_DIR, exist_ok=True)
    # 若复盘库已存在，保留已确认生效状态（人工标注），不覆盖
    confirmed_map = {}
    if os.path.exists(library_path):
        try:
            old = pd.read_excel(library_path, sheet_name="权重调整建议", dtype=str)
            if "维度" in old.columns and "是否已确认生效" in old.columns:
                confirmed_map = dict(zip(old["维度"].astype(str),
                                         old["是否已确认生效"].astype(str)))
        except Exception:
            pass
    if len(weights) and confirmed_map:
        weights["是否已确认生效"] = weights["维度"].map(
            lambda x: confirmed_map.get(str(x), "否"))

    with pd.ExcelWriter(library_path, engine="openpyxl") as w:
        entries.to_excel(w, sheet_name="复盘条目", index=False)
        weights.to_excel(w, sheet_name="权重调整建议", index=False)

    print(f"复盘库已写出 -> {library_path}")
    print(f"  复盘条目     : {len(entries)} 条 (阈值 |偏差率|>{threshold:.0%})")
    print(f"  权重调整建议 : {len(weights)} 条")
    if len(entries) == 0:
        print("  (当前无高偏差样本，复盘库为空属正常——待上架N天后有真实销量再复跑)")
    else:
        print(f"  偏差方向分布 : 实际高于预期 {pos} 条 / 实际低于预期 {neg} 条")
        print("  提示: 权重调整建议需人工确认后，手动修改 scoring_weights.yaml 才生效。")
    return entries, weights, library_path


def main():
    ap = argparse.ArgumentParser(description="复盘库生成 (Task 7.3)")
    ap.add_argument("--input", default=None, help="回看对比 xlsx 路径(默认取最新)")
    ap.add_argument("--out", default=LIBRARY_PATH, help="复盘库输出路径")
    args = ap.parse_args()
    run_library(review_path=args.input, library_path=args.out)


if __name__ == "__main__":
    main()

# -*- coding: utf-8 -*-
"""
阶段2 人工确认回写脚本 (Task 2.5)

读取人工填写后的 对齐_人工确认池.xlsx，按"确认结果"列回写到主匹配表：
  确认      -> 并入 对齐_汇总.xlsx
  拒绝      -> 剔除（不进汇总）
  标记错配  -> 记入 对齐_错配库.xlsx

用法:
  python scripts/apply_review.py
  python scripts/apply_review.py --pool output/对齐_人工确认池.xlsx --summary output/对齐_汇总.xlsx

说明:
  - "确认结果"列留空的行视为未处理，跳过（可分批填写、反复运行）。
  - 确认项会追加到汇总表（保留原候选的 匹配层/匹配依据/置信度），
    并打上 "审核状态=已确认" 标记，与第1/2层自动确认项区分。
  - 汇总表若不存在则新建。
"""
import os
import argparse
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
DEFAULT_POOL = os.path.join(OUT_DIR, "对齐_人工确认池.xlsx")
DEFAULT_SUMMARY = os.path.join(OUT_DIR, "对齐_汇总.xlsx")
MISMATCH_PATH = os.path.join(OUT_DIR, "对齐_错配库.xlsx")

# 确认结果的合法取值（去空格、大小写不敏感）
RESULT_CONFIRM = {"确认", "confirm", "yes", "y", "是"}
RESULT_REJECT = {"拒绝", "reject", "no", "n", "否"}
RESULT_MISMATCH = {"标记错配", "错配", "mismatch", "错"}


def parse_result(v):
    """归一确认结果。返回 '确认'/'拒绝'/'标记错配'/''(未处理)。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    if s in RESULT_CONFIRM:
        return "确认"
    if s in RESULT_REJECT:
        return "拒绝"
    if s in RESULT_MISMATCH:
        return "标记错配"
    # 未识别
    return f"未知:{s}"


def main():
    ap = argparse.ArgumentParser(description="人工确认回写")
    ap.add_argument("--pool", default=DEFAULT_POOL, help="人工确认池 xlsx 路径")
    ap.add_argument("--summary", default=DEFAULT_SUMMARY, help="主匹配汇总 xlsx 路径")
    args = ap.parse_args()

    if not os.path.exists(args.pool):
        print(f"错误: 人工确认池不存在 -> {args.pool}")
        return

    print(f"读取人工确认池: {args.pool}")
    pool = pd.read_excel(args.pool, dtype=str)
    if "确认结果" not in pool.columns:
        print("错误: 池中缺少 '确认结果' 列，请先在 对齐_人工确认池.xlsx 中填写。")
        return

    pool["__res"] = pool["确认结果"].apply(parse_result)
    n_total = len(pool)
    n_blank = int((pool["__res"] == "").sum())
    n_unknown = int(pool["__res"].str.startswith("未知:", na=False).sum())
    confirmed = pool[pool["__res"] == "确认"].drop(columns="__res")
    rejected = pool[pool["__res"] == "拒绝"].drop(columns="__res")
    mismatch = pool[pool["__res"] == "标记错配"].drop(columns="__res")

    print(f"  池总行数        : {n_total}")
    print(f"  未处理(留空)    : {n_blank}")
    print(f"  确认            : {len(confirmed)}")
    print(f"  拒绝            : {len(rejected)}")
    print(f"  标记错配        : {len(mismatch)}")
    if n_unknown:
        print(f"  未识别结果(跳过): {n_unknown}")

    # ---- 读取现有汇总 ----
    if os.path.exists(args.summary):
        summary = pd.read_excel(args.summary, dtype=str)
        print(f"读取现有汇总: {args.summary}  ({len(summary)} 行)")
    else:
        summary = pd.DataFrame()
        print(f"汇总不存在，将新建: {args.summary}")

    # 确认项追加：打审核标记
    if len(confirmed):
        confirmed = confirmed.copy()
        if "审核状态" not in confirmed.columns:
            confirmed["审核状态"] = "已确认(人工)"
        else:
            confirmed["审核状态"] = confirmed["审核状态"].fillna("已确认(人工)")
        # 去重：避免重复回写（按 竞品Sku_id + 自家SKUID 判重）
        key_cols = ["竞品Sku_id", "自家SKUID"]
        existing_keys = set()
        if len(summary) and all(c in summary.columns for c in key_cols):
            existing_keys = set(zip(summary[key_cols[0]].astype(str),
                                    summary[key_cols[1]].astype(str)))
        new_rows = []
        for _, r in confirmed.iterrows():
            k = (str(r.get(key_cols[0], "")), str(r.get(key_cols[1], "")))
            if k in existing_keys:
                continue
            new_rows.append(r)
        if new_rows:
            summary = pd.concat([summary, pd.DataFrame(new_rows)], ignore_index=True)
            print(f"  新增确认匹配    : {len(new_rows)} 条")
        else:
            print("  无新增确认匹配（均已存在于汇总）")

    # 确保汇总有审核状态列（第1/2层自动确认项标记）
    if len(summary):
        if "审核状态" not in summary.columns:
            summary["审核状态"] = ""
        summary["审核状态"] = summary["审核状态"].fillna("")
        # 匹配层为 1/2 且审核状态为空的 -> 自动确认
        if "匹配层" in summary.columns:
            mask = (summary["匹配层"].astype(str).isin(["1", "2"])) & (summary["审核状态"] == "")
            summary.loc[mask, "审核状态"] = "自动确认"

    # ---- 写出汇总 ----
    summary.to_excel(args.summary, index=False)
    print(f"回写汇总 -> {args.summary}  ({len(summary)} 行)")

    # ---- 错配库 ----
    prev_mm = []
    if os.path.exists(MISMATCH_PATH):
        prev_mm_df = pd.read_excel(MISMATCH_PATH, dtype=str)
        prev_mm = [prev_mm_df]
    if len(mismatch):
        mismatch = mismatch.copy()
        mismatch["错配来源"] = "人工标记"
        prev_mm.append(mismatch)
    if prev_mm:
        mm_all = pd.concat(prev_mm, ignore_index=True)
        # 按 竞品Sku_id + 自家SKUID 去重
        if all(c in mm_all.columns for c in ["竞品Sku_id", "自家SKUID"]):
            mm_all = mm_all.drop_duplicates(["竞品Sku_id", "自家SKUID"], keep="last")
        mm_all.to_excel(MISMATCH_PATH, index=False)
        print(f"错配库 -> {MISMATCH_PATH}  ({len(mm_all)} 行)")
    else:
        print(f"错配库: 无错配记录（{MISMATCH_PATH} 未生成）")

    # ---- 拒绝项日志 ----
    if len(rejected):
        rej_path = os.path.join(OUT_DIR, "对齐_拒绝记录.xlsx")
        rejected.to_excel(rej_path, index=False)
        print(f"拒绝记录 -> {rej_path}  ({len(rejected)} 行)")

    print("\n回写完成。")
    print(f"  汇总最终行数  : {len(summary)}")
    if len(summary) and "审核状态" in summary.columns:
        print(f"  其中自动确认  : {int((summary['审核状态']=='自动确认').sum())}")
        print(f"  其中人工确认  : {int((summary['审核状态']=='已确认(人工)').sum())}")


if __name__ == "__main__":
    main()

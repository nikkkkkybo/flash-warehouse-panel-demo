# -*- coding: utf-8 -*-
"""
阶段5 建议清单人工确认回写 (Task 5.3)

读取人工填写后的 选品建议清单.xlsx，按"确认结果"列流转状态：
  确认      -> 状态改为"待入库"，写入 建议清单_已确认_待入库.xlsx
  修改      -> 按"修改后值"调整建议值，状态改为"待入库"，写入 建议清单_已确认_待入库.xlsx
  拒绝      -> 状态改为"已归档"，写入 建议清单_已拒绝_归档.xlsx

状态流转链：待人工确认 -> 待入库 -> 已入库
  （本脚本只负责 待人工确认 -> 待入库；"已入库"由实际入库操作另行标记）

说明:
  - "确认结果"列留空的行视为未处理，跳过（可分批填写、反复运行）。
  - 回写时基于"建议ID"做主键去重，避免重复流转。
  - 修改类会把"修改后值"合并到对应建议值字段（覆盖"建议引入价区间"或
    "价格建议_建议售价"等，按建议类型路由）。

用法:
  python scripts/apply_suggestion_review.py
  python scripts/apply_suggestion_review.py --input output/选品建议清单.xlsx
"""
import os
import argparse
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
DEFAULT_INPUT = os.path.join(OUT_DIR, "选品建议清单.xlsx")
CONFIRMED_PATH = os.path.join(OUT_DIR, "建议清单_已确认_待入库.xlsx")
REJECTED_PATH = os.path.join(OUT_DIR, "建议清单_已拒绝_归档.xlsx")

# 确认结果合法取值（大小写不敏感）
RESULT_CONFIRM = {"确认", "confirm", "yes", "y", "是"}
RESULT_MODIFY = {"修改", "modify", "edit", "改"}
RESULT_REJECT = {"拒绝", "reject", "no", "n", "否"}

STATUS_PENDING = "待人工确认"
STATUS_TODO = "待入库"
STATUS_ARCHIVED = "已归档"

# 修改后值要回写的目标列（按建议类型路由）
MODIFY_TARGET_COL_INTRO = "建议引入价区间"      # 引入新品：修改后值覆盖引入价区间
MODIFY_TARGET_COL_OPT = "价格建议_建议售价"      # 优化现有：修改后值覆盖调价建议


def parse_result(v):
    """归一确认结果。返回 '确认'/'修改'/'拒绝'/''(未处理)。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    if s in RESULT_CONFIRM:
        return "确认"
    if s in RESULT_MODIFY:
        return "修改"
    if s in RESULT_REJECT:
        return "拒绝"
    return f"未知:{s}"


def _clean(v):
    """空值统一为空串。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def apply_modify(row):
    """对"修改"类行，把"修改后值"合并到对应建议值列。

    规则：
      - 引入新品 -> 覆盖"建议引入价区间"
      - 优化现有 -> 覆盖"价格建议_建议售价"
      其余列保持原样。
    """
    r = row.copy()
    modify_val = _clean(r.get("修改后值"))
    if not modify_val:
        return r
    rtype = _clean(r.get("建议类型"))
    if rtype == "引入新品" and MODIFY_TARGET_COL_INTRO in r:
        r[MODIFY_TARGET_COL_INTRO] = modify_val
    elif rtype == "优化现有" and MODIFY_TARGET_COL_OPT in r:
        r[MODIFY_TARGET_COL_OPT] = modify_val
    return r


def main():
    ap = argparse.ArgumentParser(description="选品建议清单人工确认回写")
    ap.add_argument("--input", default=DEFAULT_INPUT, help="选品建议清单 xlsx 路径")
    ap.add_argument("--confirmed", default=CONFIRMED_PATH, help="待入库输出路径")
    ap.add_argument("--rejected", default=REJECTED_PATH, help="已拒绝归档输出路径")
    args = ap.parse_args()

    if not os.path.exists(args.input):
        print(f"错误: 建议清单不存在 -> {args.input}")
        return

    print(f"读取建议清单: {args.input}")
    df = pd.read_excel(args.input, dtype=str)
    if "确认结果" not in df.columns:
        print("错误: 清单中缺少 '确认结果' 列，请先在 选品建议清单.xlsx 中填写。")
        return
    if "建议ID" not in df.columns:
        print("错误: 清单中缺少 '建议ID' 列，无法做主键去重。")
        return

    df["__res"] = df["确认结果"].apply(parse_result)
    n_total = len(df)
    n_blank = int((df["__res"] == "").sum())
    n_unknown = int(df["__res"].astype(str).str.startswith("未知:", na=False).sum())

    confirmed = df[df["__res"] == "确认"].drop(columns="__res")
    modified = df[df["__res"] == "修改"].drop(columns="__res")
    rejected = df[df["__res"] == "拒绝"].drop(columns="__res")

    print(f"  清单总行数      : {n_total}")
    print(f"  未处理(留空)    : {n_blank}")
    print(f"  确认            : {len(confirmed)}")
    print(f"  修改            : {len(modified)}")
    print(f"  拒绝            : {len(rejected)}")
    if n_unknown:
        print(f"  未识别结果(跳过): {n_unknown}")

    # ---- 处理"修改"类：合并修改后值 ----
    if len(modified):
        modified = modified.apply(apply_modify, axis=1)

    # ---- 合并 确认 + 修改 -> 待入库 ----
    todo = pd.concat([confirmed, modified], ignore_index=True)
    if len(todo):
        todo["状态"] = STATUS_TODO

    # ---- 拒绝 -> 已归档 ----
    if len(rejected):
        rejected = rejected.copy()
        rejected["状态"] = STATUS_ARCHIVED

    # ---- 基于建议ID与既有待入库表去重（保留本次最新）----
    if os.path.exists(args.confirmed):
        prev_todo = pd.read_excel(args.confirmed, dtype=str)
        prev_ids = set(prev_todo["建议ID"].astype(str).unique())
        new_todo = todo[~todo["建议ID"].astype(str).isin(prev_ids)]
        kept_todo = pd.concat([prev_todo, new_todo], ignore_index=True)
        print(f"  既有待入库      : {len(prev_todo)}  本次新增: {len(new_todo)}")
    else:
        kept_todo = todo
        print(f"  待入库表新建    : {len(todo)} 条")

    if len(kept_todo):
        kept_todo.to_excel(args.confirmed, index=False)
        print(f"待入库 -> {args.confirmed}  ({len(kept_todo)} 行)")
    else:
        print(f"待入库: 无确认/修改记录（{args.confirmed} 未生成）")

    # ---- 拒绝归档（按建议ID去重，保留最新）----
    if os.path.exists(args.rejected):
        prev_rej = pd.read_excel(args.rejected, dtype=str)
        prev_rej_ids = set(prev_rej["建议ID"].astype(str).unique())
        new_rej = rejected[~rejected["建议ID"].astype(str).isin(prev_rej_ids)]
        kept_rej = pd.concat([prev_rej, new_rej], ignore_index=True)
    else:
        kept_rej = rejected

    if len(kept_rej):
        kept_rej.to_excel(args.rejected, index=False)
        print(f"已拒绝归档 -> {args.rejected}  ({len(kept_rej)} 行)")
    else:
        print(f"已拒绝: 无拒绝记录（{args.rejected} 未生成）")

    # ---- 回写原清单的状态列（便于后续查看处理进度）----
    status_map = {}
    for _, r in todo.iterrows():
        status_map[_clean(r.get("建议ID"))] = STATUS_TODO
    for _, r in rejected.iterrows():
        status_map[_clean(r.get("建议ID"))] = STATUS_ARCHIVED
    df["状态"] = df.apply(
        lambda r: status_map.get(_clean(r.get("建议ID")), _clean(r.get("状态"))),
        axis=1,
    )
    df = df.drop(columns="__res")
    df.to_excel(args.input, index=False)
    print(f"\n原清单状态已回写 -> {args.input}")

    # ---- 汇总 ----
    print("\n回写完成。")
    print(f"  待入库(确认+修改): {len(todo)} 条")
    print(f"  已归档(拒绝)     : {len(rejected)} 条")
    print(f"  未处理           : {n_blank} 条")


if __name__ == "__main__":
    main()

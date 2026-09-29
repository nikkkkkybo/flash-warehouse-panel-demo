# -*- coding: utf-8 -*-
"""
阶段7 Task 7.2 回看脚本

职责：
  读取 output/选品决策记录.xlsx，对「已上架」且上架日期满 N 天（默认 7/14/30，
  可通过 --days 指定单一周期，如 --days 30 只回看满30天的）的记录，
  关联本店归一化表取该品实际月销量，计算偏差并产出对比表。

数据源：
  - 决策记录: output/选品决策记录.xlsx
  - 实际销量: output/自家_归一化后.xlsx 的 月销量 字段（按 品身份证中的 自家SKUID 关联）

计算口径：
  - 偏差     = 实际销量 - 预期销量
  - 偏差率   = 偏差 / 预期销量     (预期销量为0时记为空，避免除零)
  - 排名对比 = 当前竞品同期排名 - 竞品当时排名（当前竞品排名需人工补，缺省留空）

输出：
  output/回看对比_YYYYMMDD.xlsx
  - sheet 「回看对比」: 逐条对比明细
  - sheet 「未回看说明」: 被跳过记录的原因（未上架/上架不满N天/缺SKUID等）

说明：
  - 当前没有上架N天后的真实数据，脚本仍可跑通：输出空明细表 + 占位说明。
  - 逻辑完整：一旦决策表里有满足条件的已上架记录，自动产出对比。
"""
import os
import argparse
import pandas as pd
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
RECORD_PATH = os.path.join(OUT_DIR, "选品决策记录.xlsx")
OWN_NORM_PATH = os.path.join(OUT_DIR, "自家_归一化后.xlsx")

DEFAULT_REVIEW_DAYS = [7, 14, 30]

# 回看对比表输出列
LOOKBACK_COLUMNS = [
    "决策ID", "决策日期", "决策类型", "品身份证", "自家SKUID",
    "竞品当时排名", "竞品当时热度分", "竞品当时月售原值",
    "引入理由快照", "预期排名", "预期销量",
    "采购价快照", "建议售价", "上架日期", "回看周期天",
    "实际销量", "偏差", "偏差率",
    "当前竞品同期排名", "排名变化",
    "备注",
]


def _to_date(s):
    """把各种格式解析为 pandas Timestamp，失败返回 NaT。"""
    if s is None:
        return pd.NaT
    if isinstance(s, float) and pd.isna(s):
        return pd.NaT
    t = pd.to_datetime(s, errors="coerce")
    return t


def _parse_skuid(pid):
    """从 品身份证 '竞品商品ID|自家SKUID' 拆出自家SKUID。"""
    if not isinstance(pid, str) or "|" not in pid:
        return ""
    return pid.split("|", 1)[1].strip()


def load_actual_sales(own_norm_path=OWN_NORM_PATH):
    """
    读取本店归一化表，返回 {SKUID: 月销量}。
    文件不存在或无 月销量 列时返回空 dict（脚本仍可跑通，实际销量记为空）。
    """
    if not os.path.exists(own_norm_path):
        print(f"警告: 本店归一化表不存在 -> {own_norm_path}（实际销量将留空）")
        return {}
    df = pd.read_excel(own_norm_path, dtype=str)
    if "SKUID" not in df.columns:
        print(f"警告: 本店归一化表无 SKUID 列 -> {own_norm_path}")
        return {}
    if "月销量" not in df.columns:
        print(f"警告: 本店归一化表无 月销量 列 -> {own_norm_path}")
        return {}
    df = df[["SKUID", "月销量"]].copy()
    df["SKUID"] = df["SKUID"].astype(str).str.strip()
    df["月销量"] = pd.to_numeric(df["月销量"], errors="coerce")
    # 同一 SKUID 取第一条非空
    df = df.dropna(subset=["月销量"]).drop_duplicates("SKUID", keep="first")
    return dict(zip(df["SKUID"], df["月销量"]))


def run_lookback(record_path=RECORD_PATH, review_days=None, today=None,
                 own_norm_path=OWN_NORM_PATH):
    """
    执行回看。返回 (明细 DataFrame, 未回看说明 DataFrame)。

    参数:
      review_days: list[int] 回看周期，对每条记录按其达到的最大周期标注「回看周期天」。
                   传入单元素列表如 [30] 则只回看满30天的。
      today: 指定“今天”用于测试（默认系统当前日期）。
    """
    if review_days is None:
        review_days = DEFAULT_REVIEW_DAYS
    review_days = sorted(set(int(d) for d in review_days))
    today_ts = pd.to_datetime(today) if today else pd.Timestamp.now().normalize()

    if not os.path.exists(record_path):
        print(f"错误: 决策记录表不存在 -> {record_path}")
        print("请先运行: python record_decision.py --init")
        return pd.DataFrame(columns=LOOKBACK_COLUMNS), pd.DataFrame()

    df = pd.read_excel(record_path, dtype=str)
    # 校验字段
    for c in ["决策状态", "上架日期", "品身份证", "预期销量"]:
        if c not in df.columns:
            print(f"错误: 决策记录表缺少必要列 '{c}'")
            return pd.DataFrame(columns=LOOKBACK_COLUMNS), pd.DataFrame()

    sales_map = load_actual_sales(own_norm_path)

    detail_rows = []
    skip_rows = []

    for _, r in df.iterrows():
        base = {
            "决策ID": r.get("决策ID", ""),
            "决策日期": r.get("决策日期", ""),
            "决策类型": r.get("决策类型", ""),
            "品身份证": r.get("品身份证", ""),
            "竞品当时排名": r.get("竞品当时排名", ""),
            "竞品当时热度分": r.get("竞品当时热度分", ""),
            "竞品当时月售原值": r.get("竞品当时月售原值", ""),
            "引入理由快照": r.get("引入理由快照", ""),
            "预期排名": r.get("预期排名", ""),
            "预期销量": r.get("预期销量", ""),
            "采购价快照": r.get("采购价快照", ""),
            "建议售价": r.get("建议售价", ""),
            "上架日期": r.get("上架日期", ""),
            "回看周期天": "",
            "实际销量": "",
            "偏差": "",
            "偏差率": "",
            "当前竞品同期排名": "",
            "排名变化": "",
            "备注": "",
        }

        status = str(r.get("决策状态", "")).strip()
        shelve = _to_date(r.get("上架日期"))

        # 跳过判定
        if status != "已上架":
            skip_rows.append({**base, "备注": f"未回看: 决策状态={status or '空'}(非已上架)"})
            continue
        if pd.isna(shelve):
            skip_rows.append({**base, "备注": "未回看: 已上架但上架日期为空"})
            continue

        days_on_shelf = (today_ts - shelve.normalize()).days
        # 取该记录已达成的最大回看周期
        reached = [d for d in review_days if days_on_shelf >= d]
        if not reached:
            need = review_days[0]
            skip_rows.append({
                **base,
                "备注": f"未回看: 上架{days_on_shelf}天，未满最小回看周期{need}天",
            })
            continue
        cycle = max(reached)

        # 实际销量：按品身份证拆出 自家SKUID 关联本店归一化表
        skuid = _parse_skuid(r.get("品身份证"))
        base["自家SKUID"] = skuid
        actual = sales_map.get(skuid, np.nan)

        exp_sales = pd.to_numeric(r.get("预期销量"), errors="coerce")

        if pd.isna(actual):
            actual_val = ""
            dev = ""
            dev_rate = ""
            note = "实际销量未取到(本店归一化表无此SKUID或月销量为空)"
        else:
            actual_val = float(actual)
            if pd.isna(exp_sales):
                dev = ""
                dev_rate = ""
                note = "预期销量为空，无法算偏差"
            else:
                dev = actual_val - float(exp_sales)
                dev_rate = (dev / float(exp_sales)) if float(exp_sales) != 0 else ""
                note = ""

        base.update({
            "回看周期天": cycle,
            "实际销量": actual_val,
            "偏差": dev,
            "偏差率": dev_rate,
            "备注": note,
        })

        # 排名变化：当前竞品同期排名需人工补到 决策记录 或外部表，当前缺省留空
        # 若 决策记录 备注里写了「当前排名=NN」，可解析（占位增强）
        cur_rank = ""
        rank_chg = ""
        base["当前竞品同期排名"] = cur_rank
        base["排名变化"] = rank_chg

        detail_rows.append(base)

    detail = pd.DataFrame(detail_rows, columns=LOOKBACK_COLUMNS) if detail_rows \
        else pd.DataFrame(columns=LOOKBACK_COLUMNS)
    skip = pd.DataFrame(skip_rows) if skip_rows else pd.DataFrame()

    return detail, skip


def main():
    ap = argparse.ArgumentParser(description="选品回看对比 (Task 7.2)")
    ap.add_argument("--record", default=RECORD_PATH, help="决策记录 xlsx 路径")
    ap.add_argument("--own", default=OWN_NORM_PATH, help="本店归一化表 xlsx 路径")
    ap.add_argument("--days", type=int, nargs="+", default=None,
                    help="回看周期(天)，默认 7 14 30；如 --days 30 只回看满30天的")
    ap.add_argument("--today", default=None, help="指定今天日期(测试用)，如 2026-09-15")
    args = ap.parse_args()

    review_days = args.days if args.days else DEFAULT_REVIEW_DAYS

    print("回看配置:")
    print(f"  决策记录   : {args.record}")
    print(f"  本店归一化 : {args.own}")
    print(f"  回看周期   : {review_days} 天")
    print(f"  今天       : {args.today or '(系统当前)'}")
    print()

    detail, skip = run_lookback(
        record_path=args.record,
        review_days=review_days,
        today=args.today,
        own_norm_path=args.own,
    )

    today_str = (pd.to_datetime(args.today) if args.today
                 else pd.Timestamp.now()).strftime("%Y%m%d")
    out_path = os.path.join(OUT_DIR, f"回看对比_{today_str}.xlsx")
    os.makedirs(OUT_DIR, exist_ok=True)

    with pd.ExcelWriter(out_path, engine="openpyxl") as w:
        detail.to_excel(w, sheet_name="回看对比", index=False)
        if len(skip):
            skip.to_excel(w, sheet_name="未回看说明", index=False)
        else:
            pd.DataFrame(columns=["说明"]).to_excel(w, sheet_name="未回看说明", index=False)

    print(f"回看完成 -> {out_path}")
    print(f"  对比明细   : {len(detail)} 条")
    print(f"  未回看说明 : {len(skip)} 条")
    if len(detail) == 0:
        print("  (当前无满足回看条件的已上架记录，对比明细为空属正常——上架N天后才有真实数据)")
    if len(detail):
        n_valid = detail["偏差"].apply(lambda x: isinstance(x, (int, float)) and not pd.isna(x)).sum()
        print(f"  其中可算偏差: {int(n_valid)} 条")
        if n_valid:
            rates = pd.to_numeric(detail["偏差率"], errors="coerce").dropna()
            if len(rates):
                print(f"  偏差率: min={rates.min():.2%}, max={rates.max():.2%}, mean={rates.mean():.2%}")


if __name__ == "__main__":
    main()

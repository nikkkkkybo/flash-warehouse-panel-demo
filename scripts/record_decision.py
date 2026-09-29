# -*- coding: utf-8 -*-
"""
阶段7 Task 7.1 选品决策记录脚本

职责：
  1. 定义选品决策记录表的字段结构（表头），写出模板 xlsx。
  2. 提供 record_decision(decision_dict) 函数，把一次决策追加到
     output/选品决策记录.xlsx（若不存在则用模板创建）。
  3. 提供 CLI 入口：
       python record_decision.py --init            生成模板 + 空记录表
       python record_decision.py --add [字段参数]  追加一条记录（缺省字段交互式补录）
       python record_decision.py --mock            追加一条 mock 决策（自检用）

字段结构（表头顺序固定）：
  决策ID、决策日期、决策类型、品身份证、竞品当时排名、竞品当时热度分、
  竞品当时月售原值、引入理由快照、预期排名、预期销量、采购价快照、
  建议售价、决策状态、上架日期、备注

说明：
  - 品身份证格式约定为 "竞品商品ID|自家SKUID"（例：8888888801|9999999999999999999），
    回看脚本 review_lookback.py 会按 "|" 拆出 自家SKUID 与本店归一化表关联取实际销量。
  - 决策ID 自动生成：D + 三位流水号（D001、D002...），调用方无需传入。
  - 决策状态取值：待执行 / 已上架 / 已下架（缺省 待执行）。
"""
import os
import sys
import argparse
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
TEMPLATE_PATH = os.path.join(OUT_DIR, "选品决策记录_模板.xlsx")
RECORD_PATH = os.path.join(OUT_DIR, "选品决策记录.xlsx")

# 表头（顺序即输出顺序）
DECISION_COLUMNS = [
    "决策ID",
    "决策日期",
    "决策类型",
    "品身份证",
    "竞品当时排名",
    "竞品当时热度分",
    "竞品当时月售原值",
    "引入理由快照",
    "预期排名",
    "预期销量",
    "采购价快照",
    "建议售价",
    "决策状态",
    "上架日期",
    "备注",
]

DECISION_TYPES = {"引入新品", "优化现有"}
STATUS_TYPES = {"待执行", "已上架", "已下架"}


def _empty_frame():
    """构造一张只有表头的空表。"""
    return pd.DataFrame(columns=DECISION_COLUMNS)


def create_template():
    """生成模板 xlsx（仅表头）。同时确保 output 目录存在。"""
    os.makedirs(OUT_DIR, exist_ok=True)
    df = _empty_frame()
    df.to_excel(TEMPLATE_PATH, index=False)
    print(f"模板已写出 -> {TEMPLATE_PATH}")
    return TEMPLATE_PATH


def _next_id(existing):
    """根据现有记录生成下一个决策ID（D001起）。"""
    if existing is None or len(existing) == 0:
        return "D001"
    ids = existing["决策ID"].dropna().astype(str).tolist()
    max_n = 0
    for i in ids:
        if i.startswith("D") and i[1:].isdigit():
            max_n = max(max_n, int(i[1:]))
    return f"D{max_n + 1:03d}"


def _today():
    return pd.Timestamp.now().strftime("%Y-%m-%d")


def record_decision(decision_dict):
    """
    把一次决策追加到 选品决策记录.xlsx。

    参数:
      decision_dict: dict，键为 DECISION_COLUMNS 中的字段名（决策ID可省略，自动生成）。
                     缺省字段填空字符串；决策日期缺省取今天；决策状态缺省取 待执行。

    返回:
      (record_path, new_id) 追加后的文件路径与本次生成的决策ID。
    """
    os.makedirs(OUT_DIR, exist_ok=True)

    # 读取现有记录（不存在则用模板结构新建空表）
    if os.path.exists(RECORD_PATH):
        existing = pd.read_excel(RECORD_PATH, dtype=str)
        # 补齐可能缺失的列，保证结构一致
        for c in DECISION_COLUMNS:
            if c not in existing.columns:
                existing[c] = ""
        existing = existing[DECISION_COLUMNS]
    else:
        existing = _empty_frame()

    # 组装新行
    row = {c: "" for c in DECISION_COLUMNS}
    for k, v in decision_dict.items():
        if k in row and v is not None:
            row[k] = str(v).strip()
    # 自动字段
    row["决策ID"] = _next_id(existing)
    if not row.get("决策日期"):
        row["决策日期"] = _today()
    if not row.get("决策状态"):
        row["决策状态"] = "待执行"
    # 规范取值
    if row["决策类型"] and row["决策类型"] not in DECISION_TYPES:
        print(f"警告: 决策类型 '{row['决策类型']}' 不在 {DECISION_TYPES}，已原样保留。")
    if row["决策状态"] not in STATUS_TYPES:
        print(f"警告: 决策状态 '{row['决策状态']}' 不在 {STATUS_TYPES}，已原样保留。")

    new_id = row["决策ID"]
    existing = pd.concat([existing, pd.DataFrame([row])], ignore_index=True)
    existing = existing[DECISION_COLUMNS]
    existing.to_excel(RECORD_PATH, index=False)
    print(f"已追加决策记录 {new_id} -> {RECORD_PATH}  (共 {len(existing)} 条)")
    return RECORD_PATH, new_id


# ============================================================
# CLI
# ============================================================
def _prompt(label, default=""):
    """交互式录入一行。回车走默认值。"""
    s = input(f"{label}" + (f" [{default}]: " if default else ": ")).strip()
    return s if s else default


def cmd_init(args):
    create_template()
    if not os.path.exists(RECORD_PATH):
        _empty_frame().to_excel(RECORD_PATH, index=False)
        print(f"空记录表已创建 -> {RECORD_PATH}")
    else:
        print(f"记录表已存在，保留不动 -> {RECORD_PATH}")


def cmd_mock(args):
    """追加一条 mock 决策用于自检。"""
    mock = {
        "决策日期": _today(),
        "决策类型": "引入新品",
        "品身份证": "8888888801|9999999999999999999",
        "竞品当时排名": "3",
        "竞品当时热度分": "0.8721",
        "竞品当时月售原值": "已售320",
        "引入理由快照": "竞品LED光源类目Top3、热度分高、本店缺货、毛利空间充足(mock)",
        "预期排名": "8",
        "预期销量": "120",
        "采购价快照": "9.5",
        "建议售价": "18.9",
        "决策状态": "待执行",
        "上架日期": "",
        "备注": "mock自检记录",
    }
    record_decision(mock)


def cmd_add(args):
    """追加一条记录：命令行给了的字段直接用，没给的字段交互式补录。"""
    def val(name, cli_val):
        return cli_val if cli_val is not None else None

    d = {
        "决策日期":      val("决策日期", args.date) or _today(),
        "决策类型":      val("决策类型", args.type),
        "品身份证":      val("品身份证", args.product),
        "竞品当时排名":  val("竞品当时排名", args.comp_rank),
        "竞品当时热度分": val("竞品当时热度分", args.comp_score),
        "竞品当时月售原值": val("竞品当时月售原值", args.comp_sales),
        "引入理由快照":  val("引入理由快照", args.reason),
        "预期排名":      val("预期排名", args.exp_rank),
        "预期销量":      val("预期销量", args.exp_sales),
        "采购价快照":    val("采购价快照", args.cost),
        "建议售价":      val("建议售价", args.price),
        "决策状态":      val("决策状态", args.status) or "待执行",
        "上架日期":      val("上架日期", args.shelve_date),
        "备注":          val("备注", args.remark),
    }

    # 交互式补录关键字段（仅当未通过参数提供）
    if not d["决策类型"]:
        d["决策类型"] = _prompt("决策类型(引入新品/优化现有)", "引入新品")
    if not d["品身份证"]:
        d["品身份证"] = _prompt("品身份证(竞品商品ID|自家SKUID)")
    if not d["引入理由快照"]:
        d["引入理由快照"] = _prompt("引入理由快照")
    if not d["预期销量"]:
        d["预期销量"] = _prompt("预期销量")
    if not d["建议售价"]:
        d["建议售价"] = _prompt("建议售价")

    record_decision(d)


def build_parser():
    ap = argparse.ArgumentParser(
        description="选品决策记录 (Task 7.1)",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    sub = ap.add_subparsers(dest="cmd")

    p_init = sub.add_parser("init", help="生成模板 + 空记录表")
    p_init.set_defaults(func=cmd_init)

    p_mock = sub.add_parser("mock", help="追加一条 mock 决策(自检)")
    p_mock.set_defaults(func=cmd_mock)

    p_add = sub.add_parser("add", help="追加一条决策记录(参数缺省则交互补录)")
    p_add.add_argument("--date", help="决策日期 YYYY-MM-DD")
    p_add.add_argument("--type", dest="type", help="决策类型: 引入新品/优化现有")
    p_add.add_argument("--product", help="品身份证: 竞品商品ID|自家SKUID")
    p_add.add_argument("--comp-rank", dest="comp_rank", help="竞品当时排名")
    p_add.add_argument("--comp-score", dest="comp_score", help="竞品当时热度分")
    p_add.add_argument("--comp-sales", dest="comp_sales", help="竞品当时月售原值")
    p_add.add_argument("--reason", help="引入理由快照")
    p_add.add_argument("--exp-rank", dest="exp_rank", help="预期排名")
    p_add.add_argument("--exp-sales", dest="exp_sales", help="预期销量")
    p_add.add_argument("--cost", help="采购价快照")
    p_add.add_argument("--price", help="建议售价")
    p_add.add_argument("--status", help="决策状态: 待执行/已上架/已下架")
    p_add.add_argument("--shelve-date", dest="shelve_date", help="上架日期 YYYY-MM-DD")
    p_add.add_argument("--remark", help="备注")
    p_add.set_defaults(func=cmd_add)

    # 顶层快捷方式: python record_decision.py --init / --add / --mock
    ap.add_argument("--init", action="store_true", help="等价 init 子命令")
    ap.add_argument("--add", action="store_true", help="等价 add 子命令(交互式)")
    ap.add_argument("--mock", action="store_true", help="等价 mock 子命令")
    return ap


def main(argv=None):
    ap = build_parser()
    args = ap.parse_args(argv)

    # 顶层快捷方式兼容
    if getattr(args, "init", False):
        args.func = cmd_init
    elif getattr(args, "mock", False):
        args.func = cmd_mock
    elif getattr(args, "add", False) and not getattr(args, "func", None):
        args.func = cmd_add

    if not getattr(args, "func", None):
        ap.print_help()
        print("\n示例:")
        print("  python record_decision.py --init")
        print("  python record_decision.py --mock")
        print("  python record_decision.py add --type 引入新品 --product 8888888801|9999999999999999999 --exp-sales 120 --price 18.9")
        return

    args.func(args)


if __name__ == "__main__":
    # 兼容直接 python record_decision.py（无子命令）与带子命令两种调用
    main()

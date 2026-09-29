# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 主控脚本（菜单式 CLI）

串联全流程：数据归一化 → 跨店对齐 → 热销品发现 → 差异打分 → 建议生成
        → 入库导出 → 效果追踪 → 商品引进评估。

用法:
  python scripts/run.py            # 交互式菜单
  python scripts/run.py --auto     # 非交互式跑完整流程（跳过需人工确认的步骤）

设计:
  - 用 subprocess.run 调用各子脚本，子进程输出实时显示到当前终端
  - 每步显示执行时间；出错时显示错误并返回菜单
  - 状态检查：检查关键 output 文件存在情况和行数
  - 完整流程按顺序跑 ①-⑥，跳过需人工的回写步骤
    （apply_review.py 对齐确认回写 / apply_suggestion_review.py 建议确认回写）
"""
import os
import sys
import time
import argparse
import subprocess
from pathlib import Path

# Windows 控制台 UTF-8 输出（子进程输出编码由各脚本自身决定）
try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
SCRIPTS_DIR = Path(__file__).resolve().parent          # <项目根目录>\scripts
OUTPUT_DIR = ROOT / "output"
PY = sys.executable                                     # 当前 Python 解释器


# ============================================================
# 关键 output 文件（用于状态检查）
#   (文件名, 阶段标签)
# ============================================================
STATUS_FILES = [
    ("竞品_归一化后.xlsx", "①归一化-竞品"),
    ("自家_归一化后.xlsx", "①归一化-自家"),
    ("类目映射表.xlsx", "①归一化-类目映射"),
    ("对齐_汇总.xlsx", "②对齐-汇总(已匹配)"),
    ("对齐_人工确认池.xlsx", "②对齐-确认池(待人工)"),
    ("对齐_总览统计.xlsx", "②对齐-总览"),
    ("热销品_全表Top100.xlsx", "③热销品-Top100"),
    ("热销品_按类目Top20.xlsx", "③热销品-按类目"),
    ("差异打分_汇总.xlsx", "④差异打分"),
    ("选品建议清单.xlsx", "⑤建议清单"),
    ("建议清单_已确认_待入库.xlsx", "⑤建议-已确认待入库"),
    ("入库表_待导入.xlsx", "⑥入库表"),
    ("选品决策记录.xlsx", "⑦决策记录"),
    ("引进评估清单.xlsx", "引进评估"),
]


# ============================================================
# 工具函数
# ============================================================
def _row_count(path: Path) -> int:
    """读取 xlsx 数据行数（不含表头）。失败返回 -1。"""
    try:
        import pandas as pd
        df = pd.read_excel(path, sheet_name=0)
        return len(df)
    except Exception:
        return -1


def show_status():
    """显示当前数据状态：各关键 output 文件存在情况和行数。"""
    print("\n" + "=" * 60)
    print("当前数据状态")
    print("=" * 60)
    n_exist = 0
    for fname, label in STATUS_FILES:
        p = OUTPUT_DIR / fname
        if p.exists():
            n = _row_count(p)
            row_str = f"{n} 行" if n >= 0 else "(读取失败)"
            print(f"  [存在] {label:<22} {fname}  →  {row_str}")
            n_exist += 1
        else:
            print(f"  [缺失] {label:<22} {fname}")
    print("-" * 60)
    print(f"共 {len(STATUS_FILES)} 个关键文件，已存在 {n_exist} 个")


def run_script(script_name: str, args=None, label: str = "") -> bool:
    """调用子脚本，实时显示输出，返回是否成功。

    子进程的 stdout/stderr 直接继承当前终端，输出实时可见。
    """
    args = args or []
    script_path = SCRIPTS_DIR / script_name
    if not script_path.exists():
        print(f"[错误] 脚本不存在: {script_path}")
        return False

    cmd = [PY, str(script_path)] + [str(a) for a in args]
    title = label or script_name
    print("\n" + "#" * 60)
    print(f"# 执行: {title}")
    print(f"# 命令: {PY} {script_path} {' '.join(str(a) for a in args)}".rstrip())
    print("#" * 60)

    t0 = time.time()
    try:
        # 不捕获输出，子进程直接写到当前终端（实时）
        ret = subprocess.run(cmd, cwd=str(ROOT))
    except KeyboardInterrupt:
        print("\n[中断] 用户中止执行。")
        return False
    except Exception as e:
        print(f"[异常] 调用失败: {e}")
        return False

    elapsed = time.time() - t0
    ok = ret.returncode == 0
    status = "成功" if ok else f"失败(returncode={ret.returncode})"
    print(f"\n[{status}] {title}  耗时 {elapsed:.1f}s")
    if not ok:
        print(f"[错误] {title} 执行出错，请查看上方日志。")
    return ok


def run_full_flow():
    """完整流程：①→②→③→④→⑤→⑥，跳过需人工的回写步骤。

    - 跳过 apply_review.py（对齐后人工确认回写）
    - 跳过 apply_suggestion_review.py（建议确认回写）
    - ⑥ 入库导出依赖"建议清单_已确认_待入库.xlsx”，若不存在则提示跳过。
    """
    print("\n" + "=" * 60)
    print("完整流程（一键跑 ①-⑥，跳过需人工确认的回写步骤）")
    print("=" * 60)
    print("说明：")
    print("  - 跳过 apply_review.py        （②对齐后人工确认回写）")
    print("  - 跳过 apply_suggestion_review.py （⑤建议确认回写）")
    print("  - ⑥入库导出需先人工确认建议；若未确认则提示跳过\n")

    steps = [
        ("normalize.py", [], "① 数据归一化"),
        ("align.py", [], "② 跨店对齐"),
        ("hot_products.py", [], "③ 热销品发现"),
        ("score_diff.py", [], "④ 差异打分"),
        ("build_suggestions.py", [], "⑤ 建议生成"),
    ]

    results = []
    for script, args, label in steps:
        ok = run_script(script, args, label)
        results.append((label, ok))
        if not ok:
            print(f"\n[中止] {label} 失败，后续步骤终止。")
            break

    # ⑥ 入库导出：依赖人工确认后的待入库清单
    pending_file = OUTPUT_DIR / "建议清单_已确认_待入库.xlsx"
    if not pending_file.exists():
        print("\n[跳过] ⑥ 入库导出：缺少 output/建议清单_已确认_待入库.xlsx")
        print("       请先在Web面板/Excel中确认选品建议，")
        print("       再运行 apply_suggestion_review.py 产出待入库清单后单独跑 ⑥。")
        results.append(("⑥ 入库导出", None))
    else:
        ok = run_script("export_to_template.py", ["--generic"], "⑥ 入库导出(--generic)")
        results.append(("⑥ 入库导出", ok))

    # 汇总
    print("\n" + "=" * 60)
    print("完整流程汇总")
    print("=" * 60)
    for label, ok in results:
        if ok is True:
            mark = "[成功]"
        elif ok is False:
            mark = "[失败]"
        else:
            mark = "[跳过]"
        print(f"  {mark} {label}")


# ============================================================
# 菜单
# ============================================================
MENU = """
==== 闪电仓选品工具 ====
请选择：
 0. 查看当前数据状态（不执行）
 1. 完整流程（一键跑①-⑥，跳过需人工的步骤）
 2. ① 数据归一化（normalize.py）
 3. ② 跨店对齐（align.py）
 4. 导入新竞品表（import_adapter.py）
 5. 竞品优点分析（build_advantage_analysis.py）
 6. ③ 热销品发现（hot_products.py）
 7. ④ 差异打分（score_diff.py）
 8. ⑤ 建议生成（build_suggestions.py）
 9. ⑥ 入库导出（export_to_template.py --generic）
10. ⑦ 效果追踪-记录决策（record_decision.py --mock）
11. ⑦ 效果追踪-回看（review_lookback.py）
12. 启动Web人工确认面板（web_panel.py）
13. 商品引进评估（introduce_eval.py）
14. 退出
"""

# 菜单项 -> (脚本, 参数, 标签)
# 注：4（导入新竞品表）需交互输入，在 menu_loop 中单独处理，不在此表
MENU_ACTIONS = {
    "2": ("normalize.py", [], "① 数据归一化"),
    "3": ("align.py", [], "② 跨店对齐"),
    "5": ("build_advantage_analysis.py", [], "竞品优点分析"),
    "6": ("hot_products.py", [], "③ 热销品发现"),
    "7": ("score_diff.py", [], "④ 差异打分"),
    "8": ("build_suggestions.py", [], "⑤ 建议生成"),
    "9": ("export_to_template.py", ["--generic"], "⑥ 入库导出(--generic)"),
    "10": ("record_decision.py", ["--mock"], "⑦ 记录决策(--mock)"),
    "11": ("review_lookback.py", [], "⑦ 效果回看"),
    "13": ("introduce_eval.py", [], "商品引进评估"),
}


def _prompt(msg: str):
    """读取一行输入；stdin 耗尽或 Ctrl+C 时返回 None（信号退出菜单）。"""
    try:
        return input(msg).strip()
    except (EOFError, KeyboardInterrupt):
        return None


def run_import_adapter():
    """导入新竞品表：提示输入文件路径和来源名，先扫描保存 profile 再提取列。

    流程：
      1) import_adapter.py --scan <xlsx> --profile <来源名>   扫描列名+保存profile
      2) import_adapter.py --extract <xlsx> --profile <来源名> 按profile提取列
    """
    print("\n" + "=" * 60)
    print("导入新竞品表（import_adapter.py）")
    print("=" * 60)
    excel_path = _prompt("请输入竞品 Excel 文件路径：")
    if not excel_path:
        print("[取消] 未输入文件路径。")
        return
    # 去除两端引号（拖拽路径常带引号）
    excel_path = excel_path.strip().strip('"').strip("'")
    if not Path(excel_path).exists():
        print(f"[错误] 文件不存在: {excel_path}")
        return
    source_name = _prompt("请输入来源名（用于保存/加载 profile，如 美团导出）：")
    if not source_name:
        print("[取消] 未输入来源名。")
        return

    # 1) 扫描列名 + 保存 profile
    ok1 = run_script(
        "import_adapter.py",
        ["--scan", excel_path, "--profile", source_name],
        f"扫描列名并保存 profile（{source_name}）",
    )
    if not ok1:
        print("[中止] 扫描失败，未进行提取。")
        return
    # 2) 提取列输出标准格式表
    run_script(
        "import_adapter.py",
        ["--extract", excel_path, "--profile", source_name],
        f"提取列输出标准格式表（profile={source_name}）",
    )


def menu_loop():
    """交互式菜单主循环。"""
    while True:
        # 顶部状态摘要
        show_status()
        print(MENU)
        choice = _prompt("请输入选项编号 > ")
        if choice is None:
            print("\n再见。")
            return

        if choice == "14" or choice.lower() in ("q", "quit", "exit"):
            print("再见。")
            return
        if choice == "0":
            # 已在顶部显示，再次详细展示
            show_status()
            continue
        if choice == "1":
            run_full_flow()
            if _prompt("\n按回车返回菜单...") is None:
                print("\n再见。")
                return
            continue
        if choice == "4":
            run_import_adapter()
            if _prompt("\n按回车返回菜单...") is None:
                print("\n再见。")
                return
            continue
        if choice == "12":
            web_path = SCRIPTS_DIR / "web_panel.py"
            if not web_path.exists():
                print(f"\n[提示] web_panel.py 不存在（{web_path}）。")
            else:
                print("\n[启动] Web人工确认面板，浏览器打开 http://127.0.0.1:5000")
                print("[提示] Ctrl+C 停止服务后回到此菜单")
                proc = subprocess.Popen(
                    [sys.executable, str(web_path)],
                    cwd=str(ROOT_DIR)
                )
                try:
                    proc.wait()
                except KeyboardInterrupt:
                    proc.terminate()
                    proc.wait(timeout=5)
                    print("\n[已停止] Web 面板")
            if _prompt("\n按回车返回菜单...") is None:
                print("\n再见。")
                return
            continue
        if choice in MENU_ACTIONS:
            script, args, label = MENU_ACTIONS[choice]
            run_script(script, args, label)
            if _prompt("\n按回车返回菜单...") is None:
                print("\n再见。")
                return
            continue

        print(f"[提示] 无效选项: {choice}（请输入 0-14）")


# ============================================================
# 入口
# ============================================================
def main():
    ap = argparse.ArgumentParser(description="闪电仓选品工具主控")
    ap.add_argument("--auto", action="store_true",
                    help="非交互式：直接跑完整流程后退出（跳过需人工确认的步骤）")
    args = ap.parse_args()

    if args.auto:
        print("[auto 模式] 非交互式执行完整流程...")
        show_status()
        run_full_flow()
        print("\n[auto 模式] 完成。")
        return

    menu_loop()


if __name__ == "__main__":
    main()

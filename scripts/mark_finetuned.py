# -*- coding: utf-8 -*-
"""FINETUNED 标记守卫：只有满足全部条件才允许写入 scripts/models/FINETUNED。

条件（缺一拒绝）：
  1. ft 模型目录存在且含 metrics.json，且 metrics.accepted == true（训练侧三门槛）
  2. 该目录存在 model_quantized.onnx（验收环境 = 线上一致的 int8）
  3. 用户显式传入 --i-confirm（人工确认，绝无自动放标记路径）

写入内容：标记文件内记录 metrics 摘要，供 align 层与审计读取。
用法：
  python scripts/mark_finetuned.py --ft3 --i-confirm
"""
import argparse
import json
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MARKER = os.path.join(BASE, "scripts", "models", "FINETUNED")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--ft3", action="store_true", help="使用 bge-small-zh-v1.5-ft3")
    ap.add_argument("--i-confirm", action="store_true", help="人工确认放行")
    args = ap.parse_args()
    ft_dir = os.path.join(BASE, "scripts", "models",
                          "bge-small-zh-v1.5-ft3" if args.ft3 else "bge-small-zh-v1.5-ft")
    mp = os.path.join(ft_dir, "metrics.json")
    checks = {}
    checks["模型目录存在"] = os.path.isdir(ft_dir)
    try:
        mt = json.load(open(mp, encoding="utf-8"))
        checks["metrics.accepted(训练三门槛)"] = bool(mt.get("accepted"))
        best = mt.get("best_metrics") or {}
    except Exception as e:
        mt, best = None, {}
        checks["metrics.accepted(训练三门槛)"] = False
    checks["ONNX int8 已导出"] = os.path.exists(os.path.join(ft_dir, "model_quantized.onnx"))
    checks["人工确认(--i-confirm)"] = bool(args.i_confirm)
    for k, v in checks.items():
        print(f"  [{'PASS' if v else 'FAIL'}] {k}")
    if not all(checks.values()):
        print("拒绝放标记（任何一项 FAIL 都不放）")
        sys.exit(1)
    body = {"granted_at": __import__("time").strftime("%Y-%m-%d %H:%M:%S"),
            "ft_dir": os.path.basename(ft_dir),
            "best_metrics": best,
            "model_sha256": (mt or {}).get("model_sha256", ""),
            "note": "FINETUNED 由 mark_finetuned.py 写入，三门槛+人工确认双守卫"}
    with open(MARKER, "w", encoding="utf-8") as f:
        json.dump(body, f, ensure_ascii=False, indent=1)
    print(f"✅ FINETUNED 标记已写入 {MARKER}")
    print(f"   recall@5={best.get('recall5')} 误报={best.get('neg_gt_062')} AUC={best.get('auc')}")


if __name__ == "__main__":
    main()

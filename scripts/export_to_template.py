# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 阶段6 入库表导出 (Task 6.1 / 6.2)

把"建议清单_已确认_待入库.xlsx"中已确认的记录导出为入库模板格式。

支持三种模式（互斥）:
  --generic                  使用 scripts/template_field_mapping.yaml 中定义的通用模板字段
  --template <xlsx路径>      读取用户提供的入库模板表头，按 yaml 配置 + 自动相似度匹配
  --make-generic-template    仅生成/刷新 output/入库模板_通用.xlsx（与 --generic 同源）

可选:
  --dry-run                  只输出预览不写文件，显示行数 + 字段缺失统计
  --input <xlsx>             待入库输入（默认 output/建议清单_已确认_待入库.xlsx）
  --output <xlsx>            输出路径（默认 output/入库表_待导入.xlsx）
  --mapping <yaml>           字段映射配置（默认 scripts/template_field_mapping.yaml）

字段缺失项统一用"待补"标记，不静默填空。

用法:
  python scripts/export_to_template.py --generic --dry-run
  python scripts/export_to_template.py --generic
  python scripts/export_to_template.py --template 用户提供的入库模板.xlsx
  python scripts/export_to_template.py --make-generic-template
"""
import os
import sys
import re
import argparse
from pathlib import Path
from difflib import SequenceMatcher

import pandas as pd
import yaml

BASE = Path(__file__).resolve().parent.parent
OUT_DIR = BASE / "output"
SCRIPTS_DIR = BASE / "scripts"

DEFAULT_INPUT = OUT_DIR / "建议清单_已确认_待入库.xlsx"
DEFAULT_OUTPUT = OUT_DIR / "入库表_待导入.xlsx"
DEFAULT_GENERIC_TEMPLATE = OUT_DIR / "入库模板_通用.xlsx"
DEFAULT_MAPPING = SCRIPTS_DIR / "template_field_mapping.yaml"


def _clean(v):
    """空值统一为空串。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def load_mapping(yaml_path):
    with open(yaml_path, encoding="utf-8") as f:
        return yaml.safe_load(f)


def read_template_columns(template_path):
    """读取模板文件的首行表头（仅读表头，避免大数据加载）。"""
    if not os.path.exists(template_path):
        print(f"错误: 模板文件不存在 -> {template_path}")
        sys.exit(1)
    df = pd.read_excel(template_path, nrows=0, dtype=str)
    return list(df.columns)


def _tokenize(s):
    return set(re.findall(r"[\w]+", str(s or "").lower()))


def _similarity(a, b, use_token_overlap=True):
    """字段名相似度：取整体 ratio 与 token 重叠的较大值。"""
    a = str(a or "").strip().lower()
    b = str(b or "").strip().lower()
    if not a or not b:
        return 0.0
    base = SequenceMatcher(None, a, b).ratio()
    if use_token_overlap:
        ta, tb = _tokenize(a), _tokenize(b)
        if ta and tb:
            overlap = len(ta & tb) / max(len(ta | tb), 1)
            base = max(base, overlap)
    return base


def auto_map_one(target_col, source_cols, threshold=0.6, use_token_overlap=True):
    """对单个 target_col，从 source_cols 找最相似的列，返回 (best_src, score) 或 (None, 0)。"""
    best_src, best_score = None, 0.0
    for sc in source_cols:
        score = _similarity(target_col, sc, use_token_overlap)
        if score > best_score:
            best_score, best_src = score, sc
    if best_score >= threshold and best_src:
        return best_src, round(best_score, 2)
    return None, 0.0


def ensure_generic_template(mapping_cfg):
    """如果通用模板文件缺失，按 yaml 中字段定义生成一个空表头文件。"""
    fields = list(mapping_cfg["generic_template"]["fields"])
    if not DEFAULT_GENERIC_TEMPLATE.exists():
        df_empty = pd.DataFrame(columns=fields)
        df_empty.to_excel(str(DEFAULT_GENERIC_TEMPLATE), index=False)
        print(f"通用模板已生成 -> {DEFAULT_GENERIC_TEMPLATE}")
    return fields


def build_field_map(mapping_cfg, source_cols, target_cols):
    """构建 target_col -> {sources, default} 取值规则。

    优先使用 yaml mapping 中的显式规则；未声明的 target_col 用相似度自动匹配。
    """
    field_map_cfg = mapping_cfg.get("mapping", {})
    placeholder = mapping_cfg.get("placeholder", "待补")
    auto_cfg = mapping_cfg.get("auto_match", {})
    threshold = float(auto_cfg.get("similarity_threshold", 0.6))
    use_token = bool(auto_cfg.get("use_token_overlap", True))

    result = {}
    auto_results = {}  # 记录哪些 target 是自动匹配的，便于日志
    for tc in target_cols:
        rule = field_map_cfg.get(tc, {}) or {}
        declared = bool(rule)  # 只要 yaml 中存在该 key 即视为已声明
        sources = []
        if rule.get("source"):
            sources.append(rule["source"])
        for fb in (rule.get("fallback") or []):
            if fb and fb not in sources:
                sources.append(fb)
        # 过滤掉源中不存在的列
        sources = [s for s in sources if s in source_cols]
        default = rule.get("default", placeholder)
        result[tc] = {"sources": sources, "default": default, "declared": declared}

        # 自动匹配规则：
        #   - 已显式声明（即便只设了 default）-> 尊重配置，不自动匹配
        #   - 完全未声明 -> 按字段名相似度自动匹配
        if not declared:
            best_src, score = auto_map_one(tc, source_cols, threshold, use_token)
            if best_src:
                result[tc]["sources"] = [best_src]
                result[tc]["auto"] = (best_src, score)
                auto_results[tc] = (best_src, score, "auto")
    return result, auto_results


def fill_rows(df, target_cols, field_map):
    """按 field_map 规则把 df 每行映射到 target_cols。
    返回 (out_rows, missing_stats)。"""
    out_rows = []
    missing_stats = {tc: 0 for tc in target_cols}
    for _, row in df.iterrows():
        out_row = {}
        for tc in target_cols:
            rule = field_map[tc]
            val = ""
            for src in rule["sources"]:
                v = _clean(row.get(src))
                if v:
                    val = v
                    break
            if not val:
                val = rule["default"]
                missing_stats[tc] += 1
            out_row[tc] = val
        out_rows.append(out_row)
    return out_rows, missing_stats


def print_preview(target_cols, out_rows, missing_stats, placeholder):
    n = len(out_rows)
    print(f"\n=== 导出预览 ===")
    print(f"  将导出行数: {n}")
    if not n:
        print("  (空表，仅含表头)")
        return
    print(f"\n字段缺失统计（每列缺失行数 / 总行数 {n}）:")
    for tc in target_cols:
        cnt = missing_stats[tc]
        pct = (cnt / n * 100) if n else 0
        flag = " [占位]" if cnt == n else ""
        print(f"  {tc:<10} 缺失 {cnt:>4} 行 ({pct:5.1f}%){flag}")
    print(f"\n前 3 行预览:")
    for i, r in enumerate(out_rows[:3]):
        print(f"  --- row {i+1} ---")
        for tc in target_cols:
            v = r[tc]
            tag = "  [缺失标记]" if v == placeholder else ""
            print(f"    {tc}: {v}{tag}")


def main():
    ap = argparse.ArgumentParser(description="入库表导出：把已确认待入库建议导出为入库模板格式")
    mode = ap.add_mutually_exclusive_group()
    mode.add_argument("--generic", action="store_true", help="使用通用入库模板字段")
    mode.add_argument("--template", help="读取用户提供的入库模板表头，自动映射字段")
    mode.add_argument("--make-generic-template", action="store_true",
                      help="仅生成/刷新 output/入库模板_通用.xlsx 后退出")
    ap.add_argument("--dry-run", action="store_true", help="只输出预览不写文件")
    ap.add_argument("--input", default=str(DEFAULT_INPUT), help="待入库输入 xlsx")
    ap.add_argument("--output", default=str(DEFAULT_OUTPUT), help="输出 xlsx")
    ap.add_argument("--mapping", default=str(DEFAULT_MAPPING), help="字段映射配置 yaml")
    args = ap.parse_args()

    # 默认行为：未指定模式时走 generic
    if not args.generic and not args.template and not args.make_generic_template:
        args.generic = True

    # ---- 加载映射配置 ----
    if not os.path.exists(args.mapping):
        print(f"错误: 字段映射配置不存在 -> {args.mapping}")
        return
    mapping_cfg = load_mapping(args.mapping)
    placeholder = mapping_cfg.get("placeholder", "待补")

    # ---- 模式：仅生成通用模板 ----
    if args.make_generic_template:
        ensure_generic_template(mapping_cfg)
        fields = list(mapping_cfg["generic_template"]["fields"])
        print(f"通用模板字段（共 {len(fields)} 列）: {fields}")
        print(f"路径: {DEFAULT_GENERIC_TEMPLATE}")
        return

    # ---- 决定目标列 ----
    if args.template:
        target_cols = read_template_columns(args.template)
        print(f"读取用户模板表头: {args.template}")
        print(f"  目标列（共 {len(target_cols)}）: {target_cols}")
    else:
        target_cols = ensure_generic_template(mapping_cfg)
        print(f"使用通用模板字段（共 {len(target_cols)} 列）")
        print(f"  目标列: {target_cols}")

    if not target_cols:
        print("错误: 目标列为空，无法导出。")
        return

    # ---- 读取输入 ----
    if not os.path.exists(args.input):
        print(f"错误: 待入库输入文件不存在 -> {args.input}")
        return
    df = pd.read_excel(args.input, dtype=str)
    source_cols = list(df.columns)
    print(f"\n读取待入库输入: {args.input}")
    print(f"  源列（共 {len(source_cols)}）: {source_cols}")
    print(f"  行数: {len(df)}")

    # ---- 构建字段映射 ----
    field_map, auto_results = build_field_map(mapping_cfg, source_cols, target_cols)

    print(f"\n字段映射结果:")
    for tc in target_cols:
        rule = field_map[tc]
        sources = rule["sources"]
        default = rule["default"]
        if rule.get("declared"):
            marker = " (配置)"
        elif "auto" in rule:
            marker = f" (自动 score={rule['auto'][1]})"
        else:
            marker = " (无匹配)"
        src_str = " / ".join(sources) if sources else "(无源)"
        print(f"  {tc:<10} <- {src_str:<32} default={default!r}{marker}")

    # ---- 填充输出 ----
    out_rows, missing_stats = fill_rows(df, target_cols, field_map)
    out_df = pd.DataFrame(out_rows, columns=target_cols) if target_cols else pd.DataFrame()

    # ---- 预览 ----
    print_preview(target_cols, out_rows, missing_stats, placeholder)

    if args.dry_run:
        print(f"\n[dry-run] 未写出文件。")
        return

    # ---- 写出 ----
    out_df.to_excel(args.output, index=False)
    n = len(out_df)
    print(f"\n导出 -> {args.output}  ({n} 行)")


if __name__ == "__main__":
    main()

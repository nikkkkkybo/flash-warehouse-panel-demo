# -*- coding: utf-8 -*-
"""
阶段0 Task 0.1 数据体检脚本
对竞品表、自家表做字段非空率、类型、值域检查，输出体检报告。
输出: output/data_check_report.md
"""
import os
import sys
import pandas as pd
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMP_PATH = os.path.join(BASE, "竞品_示例竞店B商品库_美团导出.xlsx")
OWN_PATH = os.path.join(BASE, "自家_门店商品精简_对齐用.xlsx")
OUT_DIR = os.path.join(BASE, "output")
OUT_REPORT = os.path.join(OUT_DIR, "data_check_report.md")

os.makedirs(OUT_DIR, exist_ok=True)


def clean_vc(d):
    """value_counts dict -> 普通 int 值，避免 np.int64 污染报告"""
    return {str(k): int(v) for k, v in dict(d).items()}


def fmt_val(v):
    """格式化单个值用于报告展示"""
    if v is None:
        return "None"
    if isinstance(v, float):
        if np.isnan(v):
            return "<空>"
        if v.is_integer():
            return str(int(v))
        return f"{v:.4f}"
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.lower() == "nan":
            return "<空字符串>"
        return s
    return str(v)


def col_stats(df, col):
    """返回单列统计信息字典"""
    s = df[col]
    n = len(s)
    # 非空率：去掉 NaN 和 空字符串/nan 文本
    def is_real(x):
        if x is None:
            return False
        if isinstance(x, float) and np.isnan(x):
            return False
        if isinstance(x, str):
            t = x.strip()
            if t == "" or t.lower() == "nan":
                return False
            return True
        return True
    real_mask = s.apply(is_real)
    non_null = int(real_mask.sum())
    non_null_rate = non_null / n * 100 if n else 0.0
    # 唯一值样本（前5个非空唯一值）
    real_vals = s[real_mask]
    uniq = real_vals.drop_duplicates().head(5).tolist()
    samples = [fmt_val(x) for x in uniq]
    dtype = str(s.dtype)
    return {
        "col": col,
        "non_null": non_null,
        "total": n,
        "non_null_rate": non_null_rate,
        "dtype": dtype,
        "samples": samples,
    }


def df_report_block(df, label):
    lines = []
    lines.append(f"### {label}（共 {len(df)} 行，{df.shape[1]} 列）\n")
    lines.append("| 列名 | 非空数 | 总数 | 非空率 | dtype | 样本值(前5唯一) |")
    lines.append("|---|---|---|---|---|---|")
    for c in df.columns:
        st = col_stats(df, c)
        sample_str = " / ".join(st["samples"]) if st["samples"] else "<无>"
        # 转义 markdown 表格中的 |
        sample_str = sample_str.replace("|", "\\|")
        lines.append(
            f"| {st['col']} | {st['non_null']} | {st['total']} | "
            f"{st['non_null_rate']:.2f}% | {st['dtype']} | {sample_str} |"
        )
    lines.append("")
    return lines


def value_domain_check(df, label):
    """关键字段值域检查"""
    lines = [f"### {label} 关键字段值域检查\n"]
    if label.startswith("竞品"):
        # 月售 文本格式分布
        if "月售" in df.columns:
            s = df["月售"].dropna().astype(str)
            # 归类格式
            import re
            patterns = {
                "已售N": r"^已售\d+\+?$",
                "N人想买": r"^\d+人想买$",
                "N人购买": r"^\d+人购买$",
                "纯数字": r"^\d+$",
                "其他": r".*",
            }
            lines.append("- **月售（文本）格式分布**:")
            for name, pat in patterns.items():
                cnt = int(s.str.match(pat).sum())
                lines.append(f"  - {name}: {cnt}")
            # 纠正"其他"计数 = 总数 - 已分类
            classified = sum(int(s.str.match(p).sum()) for p in [patterns["已售N"], patterns["N人想买"], patterns["N人购买"], patterns["纯数字"]])
            lines[-1] = f"  - 其他(未归类): {int(len(s)) - classified}"
            lines.append(f"  - 样本: {s.head(8).tolist()}")
        # 价格值域
        for pc in ["原价", "折后价", "神价"]:
            if pc in df.columns:
                v = pd.to_numeric(df[pc], errors="coerce").dropna()
                if len(v):
                    lines.append(f"- **{pc}**: min={v.min()}, max={v.max()}, mean={v.mean():.2f}")
        # 条形码 长度分布
        if "条形码" in df.columns:
            s = df["条形码"].dropna()
            lens = s.apply(lambda x: len(str(int(x))) if pd.notna(x) else 0)
            lines.append(f"- **条形码**: 非空 {len(s)} 条, 长度分布: {clean_vc(lens.value_counts())}")
        # 剩余库存 字符型
        if "剩余库存" in df.columns:
            lines.append(f"- **剩余库存** dtype={df['剩余库存'].dtype}, 样本: {df['剩余库存'].dropna().head(5).tolist()}")
    else:
        # 自家表
        if "品牌名称" in df.columns:
            s = df["品牌名称"].dropna().astype(str)
            lines.append(f"- **品牌名称**: 非空 {len(s)}, 其他品牌计数={int((s=='其他品牌').sum())}")
            lines.append(f"  - top10: {clean_vc(s.value_counts().head(10))}")
        if "商品描述" in df.columns:
            lines.append(f"- **商品描述**: 非空 {df['商品描述'].notna().sum()}")
        if "商品卖点" in df.columns:
            lines.append(f"- **商品卖点**: 非空 {df['商品卖点'].notna().sum()}")
        if "商品条码" in df.columns:
            s = df["商品条码"].dropna().astype(str)
            lens = s.str.replace(r"\.0$", "", regex=True).str.len()
            lines.append(f"- **商品条码**: 非空 {len(s)}, 长度分布: {clean_vc(lens.value_counts().head(10))}")
        if "美团闪购渠道售价" in df.columns:
            v = pd.to_numeric(df["美团闪购渠道售价"], errors="coerce").dropna()
            if len(v):
                lines.append(f"- **美团闪购渠道售价**: min={v.min()}, max={v.max()}, mean={v.mean():.2f}")
        if "采购价" in df.columns:
            v = pd.to_numeric(df["采购价"], errors="coerce").dropna()
            if len(v):
                lines.append(f"- **采购价**: min={v.min()}, max={v.max()}, 非空{len(v)}")
        if "是否有图片" in df.columns:
            lines.append(f"- **是否有图片**: {clean_vc(df['是否有图片'].value_counts())}")
        if "售卖状态" in df.columns:
            lines.append(f"- **售卖状态**: {clean_vc(df['售卖状态'].value_counts())}")
    lines.append("")
    return lines


def field_mapping_table(dfc, dfo):
    """两表字段对应关系表（类目/条码/价格/规格/销量5要素）"""
    lines = ["### 两表字段对应关系表（5要素）\n"]
    lines.append("| 对齐要素 | 竞品表字段 | 自家表字段 | 对齐情况 |")
    lines.append("|---|---|---|---|")
    mapping = [
        ("类目(三级)", "美团后台三级类", "美团闪购类目三级", "枚举值需对齐（竞品172 / 自家247），见类目映射表"),
        ("条码", "条形码", "商品条码", "竞品仅1.1%非空，自家100%；以条码对齐覆盖率极低，需以商品名+型号+规格对齐为主"),
        ("价格", "折后价/神价", "美团闪购渠道售价", "口径不同（折后 vs 售价），仅可作参考对比"),
        ("规格", "规格", "规格名称", "竞品规格多为档位/包装（白光/单控/2个），自家多为尺寸/容量（24cm/1.6L），需归一化"),
        ("销量", "月售(文本)", "月销量(数字)", "竞品为文本需解析为数字，自家已为数字"),
    ]
    for m in mapping:
        lines.append("| " + " | ".join(m) + " |")
    lines.append("")
    return lines


def main():
    print("加载数据...")
    dfc = pd.read_excel(COMP_PATH)
    dfo = pd.read_excel(OWN_PATH)
    print(f"竞品: {dfc.shape}, 自家: {dfo.shape}")

    md = []
    md.append("# 数据体检报告（阶段0 Task 0.1）\n")
    md.append(f"- 生成时间: {pd.Timestamp.now().strftime('%Y-%m-%d %H:%M:%S')}")
    md.append(f"- 竞品表: {COMP_PATH}")
    md.append(f"- 自家表: {OWN_PATH}\n")

    md.append("## 1. 字段非空率与类型检查\n")
    md.append("## A. 竞品表字段统计\n")
    md.extend(df_report_block(dfc, "竞品表"))
    md.extend(value_domain_check(dfc, "竞品表"))

    md.append("## B. 自家表字段统计\n")
    md.extend(df_report_block(dfo, "自家表"))
    md.extend(value_domain_check(dfo, "自家表"))

    md.append("## 2. 已知硬伤清单\n")
    md.extend([
        "| 序号 | 硬伤 | 影响 | 应对策略 |",
        "|---|---|---|---|",
        "| 1 | 竞品条形码非空率仅1.1%（55/4832） | 无法以条码为主键对齐 | 改以 品牌+型号+规格+类目 多键对齐 |",
        "| 2 | 自家品牌名称非空率14.63%（含'其他品牌'占位） | 大量品牌缺失 | 从商品名抽取补全；'其他品牌'视为缺失需抽取 |",
        "| 3 | 自家商品描述非空率约2% | 描述字段基本不可用 | 对齐阶段忽略描述字段 |",
        "| 4 | 竞品月售为文本（'已售N'/'N人想买'/'已售100+'） | 无法直接数值比较 | 正则解析为整数销量 |",
        "| 5 | 自家品牌名含英文括号（如'公牛（BULL）'） | 跨表品牌无法直接相等匹配 | 归一化为中文名'公牛' |",
        "| 6 | 竞品条形码为float类型（如4891199000027.0） | 易丢精度/前导零丢失 | 统一转为字符串 |",
        "| 7 | 竞品'剩余库存'为str型、含非数字 | 类型不一致 | 需清洗 |",
        "| 8 | 竞品三级类目172项 vs 自家247项，枚举不一致 | 类目对不齐 | 建立类目映射表归一 |",
        "| 9 | 规格/规格名称字段含义异构（档位 vs 尺寸 vs 包装） | 同款难以规格相等 | 拆解为 功率/容量/尺寸/电压/颜色档位 多维 |",
        "| 10 | 颜色档位（白光/暖光/中性光）在同款下分多行 | 易被误判为不同商品 | 颜色档位单独抽取，匹配时作同款不同色 |",
        "",
    ])

    md.append("## 3. 两表字段对应关系表\n")
    md.extend(field_mapping_table(dfc, dfo))

    md.append("## 4. 三级类目枚举值对比（摘要）\n")
    cc = dfc["美团后台三级类"].dropna().astype(str)
    oc = dfo["美团闪购类目三级"].dropna().astype(str)
    cset, oset = set(cc.unique()), set(oc.unique())
    both = cset & oset
    only_c = cset - oset
    only_o = oset - cset
    md.append(f"- 竞品三级类目枚举数: {len(cset)}")
    md.append(f"- 自家三级类目枚举数: {len(oset)}")
    md.append(f"- 两表共有（名称一致）: {len(both)}")
    md.append(f"- 仅竞品有: {len(only_c)}")
    md.append(f"- 仅自家有: {len(only_o)}")
    md.append(f"- 共有类目示例: {sorted(both)[:15]}")
    md.append(f"- 仅竞品示例: {sorted(only_c)[:15]}")
    md.append(f"- 仅自家示例: {sorted(only_o)[:15]}\n")

    with open(OUT_REPORT, "w", encoding="utf-8") as f:
        f.write("\n".join(md))
    print(f"体检报告已写出: {OUT_REPORT}")


if __name__ == "__main__":
    main()

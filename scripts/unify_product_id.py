# -*- coding: utf-8 -*-
"""
品身份证列统一脚本

为以下4个文件在末尾追加"品身份证"列（不破坏原有列）：
  - 对齐_汇总.xlsx:        品身份证 = 竞品品ID + "-" + 自家SKUID
                            （自家SKUID 为空时用 "竞品品ID-（自家空）"）
  - 差异打分_汇总.xlsx:     品身份证 = 竞品商品ID + "-" + 自家SKUID（值同上，从匹配组关联）
  - 热销品_全表Top100.xlsx: 品身份证 = 竞品商品ID（自家未匹配，仅竞品侧）
  - 热销品_按类目Top20.xlsx:品身份证 = 竞品商品ID（自家未匹配，仅竞品侧）

可重复运行：若"品身份证"列已存在，先删除再重新追加。
"""
import os
import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")

COL_NAME = "品身份证"


def _clean(v):
    """统一清洗单元格值：None/NaN/'nan'/'None' -> ''，其余转 str。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v)
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def make_id_pair(competitor_id, own_sku):
    """竞品品ID + "-" + 自家SKUID；自家为空时占位"（自家空）"。"""
    cid = _clean(competitor_id)
    oid = _clean(own_sku)
    if not oid:
        oid = "（自家空）"
    return f"{cid}-{oid}"


def _append_column(path, series):
    """读取 -> 删除已有品身份证列 -> 追加新列 -> 写回。返回 (df, 样例值)。"""
    df = pd.read_excel(path)
    before_cols = list(df.columns)
    before_rows = len(df)
    if COL_NAME in df.columns:
        df = df.drop(columns=[COL_NAME])
    df[COL_NAME] = series
    df.to_excel(path, index=False)
    after_cols = list(df.columns)
    sample = df[COL_NAME].iloc[0] if len(df) else ""
    print(f"[OK] {os.path.basename(path)}  行数 {before_rows}->{len(df)}  "
          f"列数 {len(before_cols)}->{len(after_cols)}  样例: {sample}")
    return df


def process_align_summary():
    """对齐_汇总.xlsx: 竞品品ID + "-" + 自家SKUID"""
    path = os.path.join(OUT_DIR, "对齐_汇总.xlsx")
    df = pd.read_excel(path)
    series = df.apply(
        lambda r: make_id_pair(r.get("竞品品ID"), r.get("自家SKUID")), axis=1
    )
    return _append_column(path, series)


def process_diff_summary():
    """差异打分_汇总.xlsx: 竞品商品ID + "-" + 自家SKUID（从匹配组关联，列已存在）"""
    path = os.path.join(OUT_DIR, "差异打分_汇总.xlsx")
    df = pd.read_excel(path)
    series = df.apply(
        lambda r: make_id_pair(r.get("竞品商品ID"), r.get("自家SKUID")), axis=1
    )
    return _append_column(path, series)


def process_hot_top100():
    """热销品_全表Top100.xlsx: 品身份证 = 竞品商品ID"""
    path = os.path.join(OUT_DIR, "热销品_全表Top100.xlsx")
    df = pd.read_excel(path)
    series = df["商品ID"].apply(_clean)
    return _append_column(path, series)


def process_hot_by_category():
    """热销品_按类目Top20.xlsx: 品身份证 = 竞品商品ID"""
    path = os.path.join(OUT_DIR, "热销品_按类目Top20.xlsx")
    df = pd.read_excel(path)
    series = df["商品ID"].apply(_clean)
    return _append_column(path, series)


def main():
    print("=" * 60)
    print("品身份证列统一脚本")
    print("=" * 60)
    process_align_summary()
    process_diff_summary()
    process_hot_top100()
    process_hot_by_category()
    print("\n完成。")


if __name__ == "__main__":
    main()

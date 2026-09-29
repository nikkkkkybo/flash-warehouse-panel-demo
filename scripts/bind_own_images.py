# -*- coding: utf-8 -*-
"""把用户提供的自家店铺图片按条形码对应写入自家商品库（自家_归一化后.xlsx 新增"商品图片url"列）。
用法：python bind_own_images.py "<excel路径>" [--dry]
  --dry 只报告匹配率不写文件
"""
import argparse
import os
import shutil

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("excel")
    ap.add_argument("--dry", action="store_true")
    args = ap.parse_args()

    src = pd.read_excel(args.excel, dtype=str).fillna("")
    print(f"源表格: {len(src)} 行")
    # 过滤：无条形码、无图片、店铺信息行（商品ID=-100）
    src = src[(src["条形码"].astype(str).str.strip() != "") &
              (src["图片url"].astype(str).str.strip() != "") &
              (src["商品ID"].astype(str) != "-100")].copy()
    print(f"有效行: {len(src)}（含条形码+图片）")
    src["条形码"] = src["条形码"].astype(str).str.strip()
    src["图片url"] = src["图片url"].astype(str).str.strip()
    # 同条码多行（同品多规格）：取第一条图片
    img_map = src.drop_duplicates("条形码", keep="first").set_index("条形码")["图片url"].to_dict()
    print(f"去重后条码数: {len(img_map)}")

    own = pd.read_excel(OWN_FILE, dtype={"商品条码": str, "SKUID": str, "SPUID": str})
    print(f"自家商品库: {len(own)} 行")
    own["商品条码"] = own["商品条码"].astype(str).str.strip()

    if "商品图片url" not in own.columns:
        own["商品图片url"] = ""
    m = own["商品条码"].isin(img_map.keys())
    own.loc[m, "商品图片url"] = own.loc[m, "商品条码"].map(img_map)
    n = int(m.sum())
    print(f"按条码匹配到图片: {n} 行 / 自家总 {len(own)}（匹配率 {n / len(own) * 100:.1f}%）")
    # 条码级未命中的：尝试按 商品ID 再匹配（源表有商品ID列吗？有）
    if n < len(img_map):
        not_found = set(img_map.keys()) - set(own.loc[m, "商品条码"])
        print(f"源条码未命中: {len(not_found)} 个（样本: {list(not_found)[:5]}）")

    if not args.dry:
        bak = OWN_FILE + ".bak_bindimg"
        if not os.path.exists(bak):
            shutil.copy2(OWN_FILE, bak)
            print(f"备份: {bak}")
        own.to_excel(OWN_FILE, index=False)
        print(f"已写入: {OWN_FILE}（图片列: 商品图片url）")
    else:
        print("（--dry 模式，未写文件）")


if __name__ == "__main__":
    main()

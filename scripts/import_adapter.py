# -*- coding: utf-8 -*-
"""
竞品表导入适配层
解决新竞品 Excel 表格式不保证一致时，自动猜列映射 + 人工兜底 + 来源配置复用。

主要 API:
  scan_columns(excel_path)             扫描 Excel 所有列名
  guess_mapping(columns)               用关键词词典自动猜列映射到标准字段
  save_profile(source_name, mapping)   保存为 config/import_profiles/<来源名>.yaml
  load_profile(source_name)            加载已有配置
  extract_columns(excel_path, mapping, output_path)
                                       按映射提取需要的列，输出标准格式表

CLI:
  python scripts/import_adapter.py --scan <xlsx> [--profile <来源名>]
  python scripts/import_adapter.py --extract <xlsx> --profile <来源名> [--output <xlsx>]
  python scripts/import_adapter.py --extract <xlsx> --config <yaml>   [--output <xlsx>]
  python scripts/import_adapter.py --save <来源名> --mapping-json '<json>'
"""
import os
import sys
import json
import argparse
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import yaml

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROFILE_DIR = os.path.join(BASE, "config", "import_profiles")
DEFAULT_OUTPUT = os.path.join(BASE, "output", "竞品_标准化.xlsx")

# 标准字段定义（工具内部统一用这些）
STANDARD_FIELDS = [
    "name", "spec", "price", "sales", "image_url", "brand",
    "category1", "category2", "category3", "barcode",
    "original_price", "discount_price", "god_price",
    "stock", "shop_name", "comp_id", "sku_id",
]

# 关键词典（猜映射用）—— 任务给定
KEYWORDS: Dict[str, List[str]] = {
    "name": ["商品名", "商品名字", "品名", "标题", "title", "名称", "商品名称"],
    "spec": ["规格", "规格名称", "规格型号", "型号规格"],
    "price": ["价格", "售价", "现价", "单价", "价格(元)", "美团闪购渠道售价"],
    "sales": ["月销", "月售", "销量", "已售", "售量", "月销量", "销售"],
    "image_url": ["图片", "图片url", "图片链接", "主图", "图片地址", "image", "img", "图片url"],
    "brand": ["品牌", "品牌名称", "品牌型号"],
    "category1": ["一级类", "类目一级", "一级分类", "美团闪购类目一级", "美团后台一级类"],
    "category2": ["二级类", "类目二级", "二级分类", "美团闪购类目二级", "美团后台二级类"],
    "category3": ["三级类", "类目三级", "三级分类", "美团闪购类目三级", "美团后台三级类"],
    "barcode": ["条码", "条形码", "商品条码", "barcode"],
    "original_price": ["原价", "划线价", "吊牌价", "市场价", "单件原价"],
    "discount_price": ["折后价", "折扣价", "优惠价", "活动价", "单件折扣价"],
    "god_price": ["神价", "到手价", "实付价"],
    "stock": ["库存", "渠道库存", "实物库存", "总库存"],
    "shop_name": ["店铺", "店铺名", "门店", "竞店", "店名"],
    "comp_id": ["商品ID", "商品id", "品ID", "商品编码", "SPU", "spu_id", "商品id"],
    "sku_id": ["Sku_id", "SKUID", "sku_id", "sku", "sku编码", "SKU编码"],
}

# 需要跳过的列名后缀（美团后台新版导出含“价格带”统计列，会与真实价格列抢映射）
SKIP_SUFFIXES = ["价格带"]

# 标准字段 -> 中文显示名（Web 面板下拉 / 界面展示用）
FIELD_CN = {
    "name": "商品名", "spec": "规格", "price": "价格", "sales": "月销量",
    "image_url": "图片链接", "brand": "品牌", "category1": "一级类目",
    "category2": "二级类目", "category3": "三级类目", "barcode": "商品条码",
    "original_price": "原价", "discount_price": "折后价", "god_price": "神价",
    "stock": "库存", "shop_name": "竞店名称", "comp_id": "竞品商品编号",
    "sku_id": "竞品单品编号",
}

# 竞店标准库列：英文内部名 -> 中文表头（写 Excel / 界面展示用）
STD_CN = {
    "comp_id": "竞品商品编号", "sku_id": "竞品单品编号",
    "category1": "一级类目", "category2": "二级类目", "category3": "三级类目",
    "name": "商品名称", "sales": "月销量", "price": "到手价",
    "discount_price": "折后价", "god_price": "神价", "spec": "规格",
    "barcode": "商品条码", "stock": "库存", "image_url": "图片链接",
    "shop_name": "竞店名称",
    "barcode_raw": "原始条码", "barcode_type": "条码类型", "category3_source": "类目来源",
    "original_price": "原价",
}
# 中文表头 -> 英文内部名（读取标准库时转回）
CN_STD = {v: k for k, v in STD_CN.items()}

UNSPECIFIED = "待指定"
# 竞品表必须有的核心列（price 可缺，三价兜底后填充）
CORE_FIELDS = ["name", "sales"]
# 价格三选一：缺失 price 时按此顺序兜底
PRICE_FALLBACK = ["god_price", "discount_price", "original_price"]


def scan_columns(excel_path: str) -> List[str]:
    """扫描 Excel 所有列名（仅读表头，nrows=1）。"""
    df = pd.read_excel(excel_path, nrows=1)
    return list(df.columns)


def guess_mapping(columns: List[str]) -> Dict[str, str]:
    """
    用关键词词典自动猜列映射到标准字段。
    返回 {excel列名: 标准字段 或 "待指定"}。

    - 对每个 Excel 列名（忽略大小写）检查是否包含关键词词典中任一关键词
    - 命中则映射到对应标准字段
    - 一个标准字段被多列命中时，取第一个命中的列，其余列标 "待指定"
    """
    col_to_field: Dict[str, str] = {}
    field_to_col: Dict[str, str] = {}  # 已占用的字段 -> 首个命中列

    for col in columns:
        col_l = str(col).lower()
        # 跳过价格带等统计列，避免与真实价格列抢映射
        if any(suf.lower() in col_l for suf in SKIP_SUFFIXES):
            col_to_field[col] = UNSPECIFIED
            continue
        candidate: Optional[str] = None
        # 按 KEYWORDS 字段顺序遍历，第一个命中的字段胜出
        for field, kws in KEYWORDS.items():
            for kw in kws:
                if kw.lower() in col_l:
                    candidate = field
                    break
            if candidate:
                break

        if candidate is None:
            col_to_field[col] = UNSPECIFIED
            continue

        # 冲突处理：同字段被多列命中，取首个
        if candidate in field_to_col:
            col_to_field[col] = UNSPECIFIED
        else:
            field_to_col[candidate] = col
            col_to_field[col] = candidate
    return col_to_field


def save_profile(source_name: str, mapping: Dict[str, str]) -> str:
    """保存为 config/import_profiles/<来源名>.yaml，返回路径。
    "待指定" 的列不写入配置文件。"""
    os.makedirs(PROFILE_DIR, exist_ok=True)
    clean = {k: v for k, v in mapping.items() if v != UNSPECIFIED}
    data = {
        "source_name": source_name,
        "mapping": clean,
        "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    path = os.path.join(PROFILE_DIR, f"{source_name}.yaml")
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)
    return path


def load_profile(source_name: str) -> Optional[Dict]:
    """加载已有配置；不存在返回 None。"""
    path = os.path.join(PROFILE_DIR, f"{source_name}.yaml")
    if not os.path.exists(path):
        return None
    with open(path, "r", encoding="utf-8") as f:
        return yaml.safe_load(f)


def extract_columns(excel_path: str, mapping: Dict[str, str],
                    output_path: str = DEFAULT_OUTPUT) -> str:
    """
    按映射提取需要的列，输出标准格式表。
    - 按 mapping 重命名列（excel列名 → 标准字段名）
    - 只保留 mapping 中有明确标准字段的列（"待指定" 的列丢弃）
    - 竞品表必须有的核心列 name/price/sales 缺则报错
    - 输出到 output_path（默认 output/竞品_标准化.xlsx）
    """
    # ID/条码列逐单元格原始读取：19 位数字（Sku_id/商品ID/条形码）若整列被 pandas
    # 推断为 float64 会丢精度（1e19 > int64 上限 9.2e18 或列内混 NaN 触发）。
    # converters 拿到的是 openpyxl 原始单元格值（float/int/str），对 float 整数
    # 还原为整数字符串，避免 19 位数字变科学计数法/截断。
    def _id_conv(v):
        if isinstance(v, float):
            if v != v:          # NaN
                return str(v)
            if v.is_integer():
                return str(int(v))
        return str(v)

    id_fields = ("comp_id", "sku_id", "barcode")
    conv = {col: _id_conv for col, field in mapping.items() if field in id_fields}
    df = pd.read_excel(excel_path, converters=conv or None)
    # 校验 mapping 中的列在 Excel 中存在
    missing_cols = [c for c in mapping if c not in df.columns]
    if missing_cols:
        raise ValueError(f"Excel 中找不到列: {missing_cols}")

    # 只保留有明确字段的列
    keep = {col: field for col, field in mapping.items() if field != UNSPECIFIED}

    # 缺核心列报错
    missing_core = [FIELD_CN.get(f, f) for f in CORE_FIELDS if f not in keep.values()]
    if missing_core:
        raise ValueError(
            f"缺少核心列: {'、'.join(missing_core)}"
            f"（核心列必须包含 商品名、月销量；价格缺失时自动用神价/折后价/原价兜底，"
            f"请在映射中补全对应 Excel 列）"
        )

    sub = df[list(keep.keys())].rename(columns=keep)

    # price 三价兜底：无 price 列时从 god_price/discount_price/original_price 复制
    if "price" not in sub.columns:
        for pf in PRICE_FALLBACK:
            if pf in sub.columns:
                sub["price"] = sub[pf]
                print(f"价格列缺失，用{FIELD_CN.get(pf, pf)}兜底")
                break
        else:
            raise ValueError("缺少核心列: 价格（神价/折后价/原价 均无，无法自动兜底，"
                             f"请在映射中补全价格类列）")

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    # 输出表头中文化（与 CN_STD 严格互逆，消费端可转回英文内部名）
    sub = sub.rename(columns={k: v for k, v in STD_CN.items() if k in sub.columns})
    sub.to_excel(output_path, index=False)
    return output_path


def _print_mapping(mapping: Dict[str, str]) -> None:
    print(f"{'Excel列名':<20}\t{'标准字段'}")
    print("-" * 50)
    for col, field in mapping.items():
        print(f"{str(col):<20}\t{field}")
    total = len(mapping)
    hit = sum(1 for v in mapping.values() if v != UNSPECIFIED)
    rate = hit / total * 100 if total else 0
    print(f"\n命中 {hit}/{total} 列，准确率 {rate:.1f}%")


def main():
    parser = argparse.ArgumentParser(
        description="竞品表导入适配层：扫描/猜映射/保存配置/提取列"
    )
    parser.add_argument("--scan", help="扫描 Excel 列名并输出猜映射结果")
    parser.add_argument("--extract", help="按映射提取列输出标准格式表")
    parser.add_argument("--profile", help="来源名（加载或保存 profile）")
    parser.add_argument("--config", help="直接指定 yaml 配置路径")
    parser.add_argument("--output", help="提取输出路径（默认 output/竞品_标准化.xlsx）")
    parser.add_argument("--save", help="保存来源配置（配合 --mapping-json）")
    parser.add_argument("--mapping-json", help="映射 JSON 字符串（配合 --save）")
    args = parser.parse_args()

    # 1) 保存配置
    if args.save:
        if not args.mapping_json:
            print("--save 需配合 --mapping-json '<json>'")
            sys.exit(1)
        mapping = json.loads(args.mapping_json)
        path = save_profile(args.save, mapping)
        print(f"已保存 profile: {path}")
        return

    # 2) 扫描 + 猜映射
    if args.scan:
        cols = scan_columns(args.scan)
        print(f"扫描到 {len(cols)} 列:")
        print(cols)
        print()
        mapping = guess_mapping(cols)
        print("猜映射结果:")
        _print_mapping(mapping)
        if args.profile:
            path = save_profile(args.profile, mapping)
            print(f"\n已保存 profile: {path}")
        return

    # 3) 提取列输出标准格式表
    if args.extract:
        if args.config:
            with open(args.config, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f)
            mapping = data.get("mapping", {})
            print(f"使用配置: {args.config}")
        elif args.profile:
            data = load_profile(args.profile)
            if not data:
                print(f"profile 不存在: {args.profile}，先 --scan --profile 保存或人工配置")
                sys.exit(1)
            mapping = data.get("mapping", {})
            print(f"使用 profile: {args.profile}")
        else:
            # 没指定 profile/config，现场扫描+猜
            cols = scan_columns(args.extract)
            mapping = guess_mapping(cols)
            print("未指定 profile，使用现场猜映射结果:")
            _print_mapping(mapping)

        out = args.output or DEFAULT_OUTPUT
        try:
            path = extract_columns(args.extract, mapping, out)
        except ValueError as e:
            print(f"提取失败: {e}")
            sys.exit(1)
        print(f"\n提取完成 -> {path}")
        out_df = pd.read_excel(path, nrows=1)
        print(f"输出列: {list(out_df.columns)}")
        return

    parser.print_help()


if __name__ == "__main__":
    main()

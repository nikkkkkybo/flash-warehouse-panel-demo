# -*- coding: utf-8 -*-
"""
多竞店标准库构建工具（批量 + 增量导入 + 自家端增强）

批量模式（处理一个目录下所有竞店表）:
  python scripts/build_comp_library.py --batch D:\\扒数据\\把数据

增量模式（后续新增竞店表一键导入）:
  python scripts/build_comp_library.py --add <新竞店表.xlsx> [--force] [--align]

自家端增强（区分 EAN 条码 / 美团内码）:
  python scripts/build_comp_library.py --own

产出（默认输出到 output/竞店标准库/）:
  - 竞品_<店名>_标准化.xlsx   每店一个标准表（统一 schema）
  - 多竞店汇总库.xlsx         全店合并（shop_name 标识来源店）
  - 竞店数据质量报告.xlsx     五维非空率 + 质量分 + 使用优先级
  - archive/<店名>_<日期>.xlsx  版本快照（供趋势分析积累历史）
"""
import os
import re
import sys
import glob
import shutil
import argparse
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from import_adapter import (scan_columns, guess_mapping, extract_columns,
                            save_profile, load_profile, STD_CN, CN_STD)

LIB_DIR = os.path.join(BASE, "output", "竞店标准库")
ARCHIVE_DIR = os.path.join(LIB_DIR, "archive")
SUMMARY_FILE = os.path.join(LIB_DIR, "多竞店汇总库.xlsx")
REPORT_FILE = os.path.join(LIB_DIR, "竞店数据质量报告.xlsx")
PROFILE_DIR = os.path.join(BASE, "config", "import_profiles")
OWN_SOURCE = os.path.join(BASE, "自家商品库_标准版本.xlsx")
OWN_OUT = os.path.join(BASE, "output", "自家_标准版本_对齐增强.xlsx")

# 统一 schema（与现有竞品_标准化.xlsx 对齐，保证 align.py 链路可直接使用）
STD_COLS = [
    "comp_id", "sku_id", "category1", "category2", "category3", "name",
    "sales", "price", "discount_price", "god_price", "spec", "barcode",
    "stock", "image_url", "shop_name",
]
# 扩展列（清洗追溯用）
EXTRA_COLS = ["barcode_raw", "barcode_type", "category3_source"]

# 质量分权重（五维）
WEIGHTS = {"barcode": 0.30, "cat3": 0.25, "spec": 0.20, "img": 0.15, "price": 0.10}


def extract_shop_name(filename: str) -> str:
    """从文件名提取店名：去扩展名/日期后缀/序号/前后缀说明词。"""
    stem = os.path.splitext(os.path.basename(filename))[0]
    # 去尾部日期：_20260804 / -20260804
    stem = re.sub(r"[_\-]?\d{8}$", "", stem)
    # 去尾部序号：(1) (1)(3) (3)
    stem = re.sub(r"\((?:\d+\)?\(?)*\d+\)$", "", stem)
    # 去前缀「数据导出-」「竞品_」「竞品-」
    stem = re.sub(r"^(数据导出[-_])", "", stem)
    stem = re.sub(r"^竞品[-_]", "", stem)
    # 去后缀说明词（长词优先）
    for w in ["商品数据_带类型", "商品数据", "_带类型", "数据_带类型",
              "_美团导出", "_美团后台导出", "_后台导出", "_美团", "_导出", "_数据"]:
        stem = stem.replace(w, "")
    # 清理首尾括号/空格/下划线
    stem = stem.strip().strip("（( \t-_")
    return stem


def clean_barcode(b):
    """条码清洗：仅保留 13 位纯数字 EAN；返回 (barcode, raw, type)。"""
    if b is None or (isinstance(b, float) and pd.isna(b)):
        return "", "", "空"
    # 浮点整数（如 6902083881405.0）先转 int，避免小数点尾巴干扰位数判断
    if isinstance(b, float) and b.is_integer():
        b = int(b)
    s = str(b).strip()
    if s.lower() in ("nan", "none") or s == "":
        return "", "", "空"
    digits = re.sub(r"\D", "", s)
    if len(digits) == 13:
        return digits, s, "标准13位条码"
    if len(digits) == 19:
        return "", s, "美团内码19位"
    return "", s, "其他"


def to_price(v):
    """价格转 float，失败返回空。"""
    try:
        f = float(v)
        return f if f > 0 else ""
    except (TypeError, ValueError):
        return ""


def standardize_one(excel_path: str, shop_name: str, out_dir: str) -> pd.DataFrame:
    """单店标准化：猜列 -> profile 复用/保存 -> 提取 -> 清洗。返回标准表 df。"""
    cols = scan_columns(excel_path)
    profile = load_profile(shop_name)
    if profile:
        mapping = profile.get("mapping", {})
        print(f"  复用配置: {shop_name}.yaml")
    else:
        mapping = guess_mapping(cols)
        save_profile(shop_name, mapping)
        print(f"  自动识别列并保存配置: {shop_name}.yaml")

    # 提取到临时标准表
    tmp = os.path.join(out_dir, f"_tmp_{shop_name}.xlsx")
    extract_columns(excel_path, mapping, tmp)

    # ID/条码列强制字符串读取（中文表头），避免 19 位 ID 与 13 位条码被 float 截断/变形
    df = pd.read_excel(tmp, dtype={"竞品商品编号": str, "竞品单品编号": str,
                                   "商品条码": str, "原始条码": str})
    os.remove(tmp)
    # 兼容旧英文列名（历史产物）
    df = df.rename(columns={k: v for k, v in CN_STD.items() if k in df.columns})

    # 清洗：条码
    barcode_res = df["barcode"].apply(clean_barcode) if "barcode" in df.columns \
        else pd.Series([("", "", "空")] * len(df))
    df["barcode"] = [r[0] for r in barcode_res]
    df["barcode_raw"] = [r[1] for r in barcode_res]
    df["barcode_type"] = [r[2] for r in barcode_res]

    # 清洗：价格数值化
    for c in ["price", "discount_price", "god_price"]:
        if c in df.columns:
            df[c] = df[c].apply(to_price)

    # price 兜底增强：若 price 非空率低，从其余价格列选非空率最高的补缺口
    if "price" in df.columns:
        rate = (df["price"].fillna("").astype(str).str.strip() != "").mean()
        if rate < 0.6:
            best_pf, best_rate = None, 0.0
            for c in ["god_price", "discount_price"]:
                if c in df.columns:
                    r = (df[c].fillna("").astype(str).str.strip() != "").mean()
                    if r > best_rate:
                        best_pf, best_rate = c, r
            if best_pf and best_rate > rate:
                mask = df["price"].fillna("").astype(str).str.strip() == ""
                df.loc[mask, "price"] = df.loc[mask, best_pf]
                cn_name = {"god_price": "神价", "discount_price": "折后价"}[best_pf]
                print(f"  价格列非空率 {rate:.0%} 偏低，用{cn_name}({best_rate:.0%})补缺口")

    # 类目兜底：category3 缺时用 category2
    cat3 = df["category3"].fillna("").astype(str) if "category3" in df.columns \
        else pd.Series([""] * len(df))
    cat2 = df["category2"].fillna("").astype(str) if "category2" in df.columns \
        else pd.Series([""] * len(df))
    df["category3_source"] = ["meituan" if c else ("shop" if c2 else "无") for c, c2 in zip(cat3, cat2)]
    df["category3"] = cat3.where(cat3 != "", cat2)

    # 补齐缺失列
    for c in STD_COLS + EXTRA_COLS:
        if c not in df.columns:
            df[c] = ""
    df["shop_name"] = shop_name
    # 输出给用户看的表头中文化（内部逻辑仍用英文列名）
    df = df.rename(columns={k: v for k, v in STD_CN.items() if k in df.columns})
    return df


def pick_best_version(files: list) -> str:
    """同店多文件时选最优：优先美团三级类目非空率最高，其次行数最多，最后最新。"""
    best = None
    best_key = None
    for f in files:
        try:
            df = pd.read_excel(f, nrows=None)
        except Exception:
            continue
        has_cat3 = any("美团三级" in str(c) or "美团后台三级" in str(c) for c in df.columns)
        n = len(df)
        mtime = os.path.getmtime(f)
        key = (has_cat3, n, mtime)
        if best_key is None or key > best_key:
            best_key = key
            best = f
    return best


def merge_shop_aliases(groups: dict) -> dict:
    """店名别名合并：若 A 是 B 的子串（视为同一家店的不同版本命名），合并到 B。"""
    names = list(groups.keys())
    merged = {}
    used = set()
    for n in sorted(names, key=len, reverse=True):
        if n in used:
            continue
        # 找所有是 n 子串的其他店名（短名 -> 长名）
        aliases = [m for m in names if m != n and m in n and m not in used]
        for a in aliases:
            print(f"  店名合并: [{a}] 并入 [{n}]（视为同店不同命名）")
            used.add(a)
        merged[n] = groups[n]
        for a in aliases:
            merged[n].extend(groups[a])
        used.add(n)
    return merged


def build_batch(batch_dir: str) -> dict:
    """批量处理目录下所有竞店表。返回 {店名: 标准表 df}。"""
    os.makedirs(LIB_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)

    files = [f for f in glob.glob(os.path.join(batch_dir, "*.xlsx"))
             if not os.path.basename(f).startswith("~$")]
    # 按店名分组（同店多版本去重）
    groups = {}
    for f in sorted(files):
        shop = extract_shop_name(f)
        groups.setdefault(shop, []).append(f)
    groups = merge_shop_aliases(groups)

    result = {}
    for shop, fs in groups.items():
        if len(fs) > 1:
            chosen = pick_best_version(fs)
            print(f"[{shop}] 同店 {len(fs)} 个文件，选择: {os.path.basename(chosen)}")
            for f in fs:
                if f != chosen:
                    print(f"    跳过重复版本: {os.path.basename(f)}")
        else:
            chosen = fs[0]
        print(f"标准化 [{shop}] <- {os.path.basename(chosen)}")
        df = standardize_one(chosen, shop, LIB_DIR)
        # 价格观测快照（优化项C）：批量导入 = 各店首次观测（基线）
        try:
            import price_obs
            n_obs = price_obs.snapshot_shop(shop, df, source="批量导入基线")
            print(f"  价格观测快照: {n_obs} 行 -> {price_obs.HIST_FILE}")
        except Exception as e:
            print(f"  [警告] 价格观测快照失败（不影响导入）: {e}")
        out_path = os.path.join(LIB_DIR, f"竞品_{shop}_标准化.xlsx")
        df.to_excel(out_path, index=False)
        print(f"  -> {out_path} ({len(df)} 行)")
        # 版本快照（趋势分析积累用）
        stamp = datetime.now().strftime("%Y%m%d")
        snap = os.path.join(ARCHIVE_DIR, f"{shop}_{stamp}.xlsx")
        df.to_excel(snap, index=False)
        result[shop] = df
    return result


def build_summary(result: dict, force_replace: dict = None):
    """合并汇总库（内部英文列名拼接，输出中文表头）。"""
    if os.path.exists(SUMMARY_FILE):
        old = pd.read_excel(SUMMARY_FILE, dtype={"竞品商品编号": str, "竞品单品编号": str})
        old = old.rename(columns={k: v for k, v in CN_STD.items() if k in old.columns})
        keep_shops = [s for s in old["shop_name"].unique() if s not in result]
        parts = [old[old["shop_name"] == s] for s in keep_shops] if keep_shops else []
    else:
        parts = []
    parts.extend(df.rename(columns={k: v for k, v in CN_STD.items() if k in df.columns})
                 for df in result.values())
    summary = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
    if len(summary):
        summary["上次导入时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary = summary.rename(columns={k: v for k, v in STD_CN.items() if k in summary.columns})
    summary.to_excel(SUMMARY_FILE, index=False)
    print(f"汇总库 -> {SUMMARY_FILE} ({len(summary)} 行, "
          f"{summary['竞店名称'].nunique() if len(summary) else 0} 店)")


def price_anomalies(result: dict) -> pd.DataFrame:
    """价格异常检测（优化项E·规则版）。
    规则1: 到手价 < 类目25分位×0.3 -> 疑似脏价/促销价
    规则2: 到手价 > 类目75分位×3   -> 疑似单位错误（元/分填反、规格混填）
    规则3: 同店同品名到手价差异>10× -> 规格混填（"1个装"和"100个装"同名）
    类目样本 <5 不做分位判定（防小样本误报）。"""
    rows = []
    MIN_CAT = 5
    ANOMALY_COLS = ["店名", "竞品单品编号", "商品名称", "三级类目", "到手价",
                    "异常类型", "参考值", "说明"]
    for shop, df in result.items():
        if not len(df) or "到手价" not in df.columns:
            continue
        prices = pd.to_numeric(df["到手价"], errors="coerce")
        if prices.notna().sum() == 0:
            continue
        work = df.assign(__p=prices)
        # 规则1/2：类目价格带
        if "三级类目" in work.columns:
            for cat, g in work[work["__p"].notna()].groupby("三级类目"):
                if len(g) < MIN_CAT:
                    continue
                p25 = float(g["__p"].quantile(0.25))
                p75 = float(g["__p"].quantile(0.75))
                for _, r in g.iterrows():
                    p = float(r["__p"])
                    if p25 > 0 and p < p25 * 0.3:
                        rows.append({"店名": shop, "竞品单品编号": r.get("竞品单品编号", ""),
                                     "商品名称": r.get("商品名称", ""), "三级类目": cat,
                                     "到手价": p, "异常类型": "疑似脏价/促销价",
                                     "参考值": f"类目25分位¥{round(p25, 2)}",
                                     "说明": f"低于价格带下限（25分位×0.3=¥{round(p25 * 0.3, 2)}）"})
                    elif p75 > 0 and p > p75 * 3:
                        rows.append({"店名": shop, "竞品单品编号": r.get("竞品单品编号", ""),
                                     "商品名称": r.get("商品名称", ""), "三级类目": cat,
                                     "到手价": p, "异常类型": "疑似单位错误",
                                     "参考值": f"类目75分位¥{round(p75, 2)}",
                                     "说明": f"高于价格带上限（75分位×3=¥{round(p75 * 3, 2)}）"})
        # 规则3：同店同品名价差>10×
        if "商品名称" in work.columns:
            for _name, g in work[work["__p"].notna()].groupby("商品名称"):
                if len(g) < 2:
                    continue
                pmin, pmax = float(g["__p"].min()), float(g["__p"].max())
                if pmin > 0 and pmax / pmin > 10:
                    for _, r in g.iterrows():
                        rows.append({"店名": shop, "竞品单品编号": r.get("竞品单品编号", ""),
                                     "商品名称": r.get("商品名称", ""),
                                     "三级类目": r.get("三级类目", ""),
                                     "到手价": float(r["__p"]), "异常类型": "规格混填(同名价差)",
                                     "参考值": f"同名价区间¥{round(pmin, 2)}~¥{round(pmax, 2)}",
                                     "说明": "同店同名到手价差异>10×，疑规格混填"})
    return pd.DataFrame(rows, columns=ANOMALY_COLS)


def quality_report(result: dict):
    """五维非空率 + 质量分 + 使用优先级 + 价格异常清单（优化项E）。"""
    rows = []
    for shop, df in result.items():
        n = len(df)
        rate = lambda c: (df[c].fillna("").astype(str).str.strip() != "").sum() / n * 100 if n else 0
        barcode_rate = (df["条码类型"] == "标准13位条码").sum() / n * 100 if n else 0
        row = {
            "店名": shop,
            "行数": n,
            "标准条码率%": round(barcode_rate, 1),
            "类目3率%": round(rate("三级类目"), 1),
            "规格率%": round(rate("规格"), 1),
            "图片率%": round(rate("图片链接"), 1),
            "价格率%": round(rate("到手价"), 1),
            "源文件": "",
        }
        rows.append(row)
    report = pd.DataFrame(rows)
    # 质量分 = 加权（图片率固定按 100 计满分权重，图片缺失主要影响第4层）
    score = (
        report["标准条码率%"] * WEIGHTS["barcode"]
        + report["类目3率%"] * WEIGHTS["cat3"]
        + report["规格率%"] * WEIGHTS["spec"]
        + report["图片率%"] * WEIGHTS["img"]
        + report["价格率%"] * WEIGHTS["price"]
    )
    report["质量分"] = score.round(1)
    report["建议优先级"] = report["质量分"].apply(lambda s: "高" if s >= 85 else ("中" if s >= 60 else "低"))
    report = report.sort_values("质量分", ascending=False).reset_index(drop=True)
    anomalies = price_anomalies(result)
    with pd.ExcelWriter(REPORT_FILE, engine="openpyxl") as w:
        report.to_excel(w, sheet_name="质量总览", index=False)
        if len(anomalies):
            anomalies.to_excel(w, sheet_name="价格异常清单", index=False)
        else:
            pd.DataFrame([{"说明": "未检出价格异常（规则：类目25分位×0.3下限 / 75分位×3上限 / 同名价差>10×）"}]) \
                .to_excel(w, sheet_name="价格异常清单", index=False)
    print(f"质量报告 -> {REPORT_FILE}（价格异常 {len(anomalies)} 条）")
    print(report[["店名", "行数", "标准条码率%", "类目3率%", "规格率%", "图片率%", "价格率%", "质量分", "建议优先级"]].to_string(index=False))
    if len(anomalies):
        print("价格异常分布:", anomalies["异常类型"].value_counts().to_dict())
    return report


def _detect_existing_shop(shop, summary):
    """同店识别：精确店名 + 别名（子串包含）。返回已有店名或 None。"""
    if summary is None or not len(summary):
        return None
    known = summary["shop_name"].unique().tolist()
    if shop in known:
        return shop
    for k in known:
        if shop in k or k in shop:
            return k
    return None


def rerun_align(shop: str) -> bool:
    """重跑该店对齐（normalize + align，输出 output/试跑_<店>/）。"""
    import subprocess
    import sys as _sys
    std_file = os.path.join(LIB_DIR, f"竞品_{shop}_标准化.xlsx")
    trial_dir = os.path.join(BASE, "output", f"试跑_{shop}")
    norm_out = os.path.join(trial_dir, "竞品_归一化后.xlsx")
    own_out = os.path.join(BASE, "output", "自家_归一化后.xlsx")
    map_file = os.path.join(BASE, "output", "类目映射表.xlsx")
    cfg = os.path.join(BASE, "config", "import_profiles", "标准库格式.yaml")
    cmds = [
        [_sys.executable, os.path.join(BASE, "scripts", "normalize.py"),
         "--comp-path", std_file, "--config", cfg, "--comp-out", norm_out,
         "--skip-own", "--own-out", own_out, "--apply-map", map_file],
        [_sys.executable, os.path.join(BASE, "scripts", "align.py"),
         "--comp", norm_out, "--own", own_out, "--outdir", trial_dir],
    ]
    for cmd in cmds:
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", cwd=BASE)
        if r.returncode != 0:
            print(f"  [警告] 对齐重跑失败: {(r.stderr or '')[-300:]}")
            return False
    return True


def add_one(excel_path: str, force: bool = False, shop: str = None):
    """增量导入一家竞店表。
    同店导入 = 数据更新（不是新店）：--force 后覆盖标准库/汇总库/archive 快照。
    无论新店/更新店，导入后都自动跑该店对齐（品级细化+4层+两层汇总库同步），
    人工池按新对齐重建（确认进度按竞品Sku_id 由面板迁移）。
    shop: 显式指定店名（Web 导入用来源名）；缺省从文件名提取。"""
    os.makedirs(LIB_DIR, exist_ok=True)
    os.makedirs(ARCHIVE_DIR, exist_ok=True)
    shop = shop or extract_shop_name(excel_path)

    if os.path.exists(SUMMARY_FILE):
        summary = pd.read_excel(SUMMARY_FILE, dtype={"竞品商品编号": str, "竞品单品编号": str})
        summary = summary.rename(columns={k: v for k, v in CN_STD.items() if k in summary.columns})
    else:
        summary = pd.DataFrame()

    existing_shop = _detect_existing_shop(shop, summary)
    if existing_shop:
        if not force:
            print(f"增量导入 [{shop}] <- {os.path.basename(excel_path)}")
            print(f"  检测到同店 [{existing_shop}] 已导入过，本次为新数据更新（价格变动/新品对齐）。")
            print(f"  覆盖更新将：替换该店标准库、留 archive 快照、重跑该店对齐。")
            print(f"  确认更新请加 --force；若确为不同店铺，请用不同文件名区分店名。")
            return None
        shop = existing_shop  # 保持店名稳定，避免同店被拆成两家
        print(f"覆盖更新 [{shop}] <- {os.path.basename(excel_path)}（数据更新模式）")
    else:
        print(f"新店导入 [{shop}] <- {os.path.basename(excel_path)}")

    df = standardize_one(excel_path, shop, LIB_DIR)
    df = df.rename(columns={k: v for k, v in CN_STD.items() if k in df.columns})
    # 价格观测快照（优化项C）：本次导入 = 一次全量价格观测，先落历史再覆盖库
    try:
        import price_obs
        n_obs = price_obs.snapshot_shop(shop, df)
        print(f"  价格观测快照: {n_obs} 行 -> {price_obs.HIST_FILE}")
    except Exception as e:
        print(f"  [警告] 价格观测快照失败（不影响导入）: {e}")
    out_path = os.path.join(LIB_DIR, f"竞品_{shop}_标准化.xlsx")
    df_cn = df.rename(columns={k: v for k, v in STD_CN.items() if k in df.columns})
    df_cn.to_excel(out_path, index=False)
    stamp = datetime.now().strftime("%Y%m%d")
    df_cn.to_excel(os.path.join(ARCHIVE_DIR, f"{shop}_{stamp}.xlsx"), index=False)
    print(f"  -> {out_path} ({len(df)} 行)")

    # 更新汇总库（覆盖该店）
    if len(summary):
        summary = summary[summary["shop_name"] != shop]
    df["上次导入时间"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    summary = pd.concat([summary, df], ignore_index=True)
    summary = summary.rename(columns={k: v for k, v in STD_CN.items() if k in summary.columns})
    summary.to_excel(SUMMARY_FILE, index=False)
    print(f"汇总库已更新 ({len(summary)} 行, {summary['竞店名称'].nunique()} 店)")

    # 更新质量报告（重算全店）
    result = {}
    shop_col = "竞店名称" if "竞店名称" in summary.columns else "shop_name"
    for s in summary[shop_col].unique():
        f = os.path.join(LIB_DIR, f"竞品_{s}_标准化.xlsx")
        if os.path.exists(f):
            result[s] = pd.read_excel(f, dtype={"竞品商品编号": str, "竞品单品编号": str})
    quality_report(result)

    # 导入后自动跑该店对齐（新店=首次对齐；更新店=重跑对齐），
    # 归一化+对齐+两层汇总库同步由 rerun_align 完成
    print("  自动对齐（归一化 + 对齐 + 两层汇总库同步）...")
    ok = rerun_align(shop)
    if ok:
        print(f"  对齐已生成: output/试跑_{shop}/（人工池/两层汇总库已同步，确认进度按竞品Sku_id 自动迁移）")
    else:
        print("  [警告] 对齐失败，请手动重跑 scripts/align.py")

    # 变动报告（优化项C）：diff 最近两次观测（降价/涨价/新品/下架/断货/undercut）
    try:
        import price_diff
        _, pd_meta = price_diff.build_report(shop)
        if pd_meta.get("note"):
            print(f"  变动报告: {pd_meta['note']}")
        else:
            print(f"  变动报告: {pd_meta.get('rows', 0)} 条 -> {pd_meta.get('report')}")
    except Exception as e:
        print(f"  [警告] 变动报告生成失败（不影响导入）: {e}")
    return df


def enhance_own():
    """自家端增强：区分真实条码(13位) / 美团内码，输出对齐增强版。"""
    print("自家端增强...")
    df = pd.read_excel(OWN_SOURCE, dtype={"SPUID": str, "SKUID": str, "商品条码": str})
    res = df["商品条码"].apply(clean_barcode)
    df["真实条码(13位)"] = [r[0] for r in res]
    df["条码类型"] = [r[2] for r in res]
    n = len(df)
    print(f"  读取: {OWN_SOURCE} ({n} 行)")
    for t in ["标准13位条码", "美团内码19位", "其他", "空"]:
        print(f"    条码类型 {t}: {int((df['条码类型'] == t).sum())}")
    img_col = "图片url" if "图片url" in df.columns else None
    if img_col:
        n_img = (df[img_col].fillna("").astype(str).str.strip() != "").sum()
        print(f"  图片url 非空率: {n_img}/{n} ({n_img/n*100:.1f}%)，缺口待数据①(美团后台带图url表)补全")
    has_img = "是否有图片" if "是否有图片" in df.columns else None
    if has_img:
        n_has = (df[has_img].fillna("").astype(str).str.strip() != "").sum()
        print(f"  是否有图片 列非空率: {n_has}/{n} ({n_has/n*100:.1f}%)")
    # 源表列「图片url」与 ID 列输出为中文（其余源表列名保持与美团后台一致）
    df = df.rename(columns={"图片url": "图片链接", "SPUID": "商品编号", "SKUID": "单品编号"})
    df.to_excel(OWN_OUT, index=False)
    print(f"  -> {OWN_OUT}")


def main():
    parser = argparse.ArgumentParser(description="多竞店标准库构建工具")
    parser.add_argument("--batch", help="批量处理目录下所有竞店表")
    parser.add_argument("--add", help="增量导入单个竞店表")
    parser.add_argument("--force", action="store_true", help="增量导入时覆盖同店旧数据")
    parser.add_argument("--own", action="store_true", help="自家端增强（EAN条码/条码类型）")
    args = parser.parse_args()

    if args.own:
        enhance_own()
        return
    if args.batch:
        result = build_batch(args.batch)
        build_summary(result)
        quality_report(result)
        return
    if args.add:
        add_one(args.add, force=args.force)
        return
    parser.print_help()


if __name__ == "__main__":
    main()

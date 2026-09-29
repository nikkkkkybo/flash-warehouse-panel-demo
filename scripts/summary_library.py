# -*- coding: utf-8 -*-
"""
两层对齐汇总库读写工具（align 引擎 / Web 面板 / 迁移脚本共用）。

文件: output/对齐汇总库.xlsx，两个 sheet：
  - SPU匹配层（粗匹配，品级）：键 = 竞店名称 + 竞品品ID
  - SKU匹配层（细匹配，规格级）：键 = 竞店名称 + 竞品Sku_id

合并规则：
  1. 确认来源=人工确认 的行数据列永不被机器覆盖（仅刷新 状态/更新时间）；
  2. 机器行按键覆盖更新，新键追加；
  3. 店内编号不在本次竞品表 -> 状态=已下架（行保留不删）；重新出现恢复 有效。
"""
import os
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIBRARY_FILE = os.path.join(BASE, "output", "对齐汇总库.xlsx")

DEFAULT_SHOP = "示例竞店B"   # 全局 output/ 目录归属店（旧单店链路）
CONFIRM_AUTO = "机器自动"
CONFIRM_MANUAL = "人工确认"
CONFIRM_AI = "AI预筛"          # 预筛分流（A2 triage）自动写入行
CONFIRM_VEC_AUTO = "向量自动"   # 向量高置信自动确认（M5·2.4a 试点；参考级，可撤销）
STATUS_ACTIVE = "有效"
STATUS_OFF = "已下架"
STATUS_REVIEW = "待抽检"       # AI预筛行专用：抽检通过前不计入有效对齐

# 对应级别（2026-08-28 口径：等同级只有 条码/人工 两种来源；机器语义一律参考级）
LEVEL_EQ = "等同"
LEVEL_REF = "参考"

SPU_SHEET = "SPU匹配层"
SKU_SHEET = "SKU匹配层"

SPU_COLS = ["竞店名称", "竞品品ID", "竞品商品名", "竞品品牌", "竞品型号", "竞品类目",
            "自家品组ID", "自家品名", "置信度", "确认来源", "更新时间", "状态",
            "对应级别", "来源"]
SKU_COLS = ["竞店名称", "竞品品ID", "竞品Sku_id", "竞品商品名", "竞品条形码",
            "自家SPUID", "自家SKUID", "自家商品名",
            "匹配层", "匹配依据", "置信度", "确认来源", "更新时间", "状态",
            "对应级别", "来源"]


def derive_level(confirm_src, match_layer=""):
    """对应级别派生：人工确认→等同；AI预筛→参考；机器自动仅条码层(匹配层=1)→等同，其余→参考。"""
    if str(confirm_src or "") == CONFIRM_MANUAL:
        return LEVEL_EQ
    if str(confirm_src or "") == CONFIRM_AI:
        return LEVEL_REF
    if str(match_layer or "").strip() in ("1", "1.0"):
        return LEVEL_EQ
    return LEVEL_REF


def load_library(path=None):
    """读取汇总库，返回 (spu_df, sku_df)。文件/sheet 不存在时返回带表头空表。"""
    path = path or LIBRARY_FILE

    def _read(sheet, cols):
        if os.path.exists(path):
            try:
                df = pd.read_excel(path, sheet_name=sheet, dtype=str)
                df = df.fillna("")
                for c in cols:
                    if c not in df.columns:
                        df[c] = ""
                # 存量行级别回填（迁移）：对应级别为空的按来源派生；来源列同步补值
                lvl_mask = df["对应级别"].astype(str).str.strip() == ""
                if lvl_mask.any():
                    df.loc[lvl_mask, "对应级别"] = [
                        derive_level(s, m) for s, m in
                        zip(df.loc[lvl_mask, "确认来源"],
                            df.loc[lvl_mask, "匹配层" if "匹配层" in df.columns else "确认来源"])]
                src_mask = df["来源"].astype(str).str.strip() == ""
                if src_mask.any():
                    if "匹配层" in df.columns:
                        df.loc[src_mask, "来源"] = df.loc[src_mask, "匹配层"]
                    df.loc[df["来源"].astype(str).str.strip() == "", "来源"] = \
                        df.loc[df["来源"].astype(str).str.strip() == "", "确认来源"]
                return df
            except (ValueError, FileNotFoundError):
                pass
        return pd.DataFrame(columns=cols)

    return _read(SPU_SHEET, SPU_COLS), _read(SKU_SHEET, SKU_COLS)


def save_library(spu_df, sku_df, path=None):
    """写回汇总库（两 sheet）。"""
    path = path or LIBRARY_FILE
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        spu_df.to_excel(writer, sheet_name=SPU_SHEET, index=False)
        sku_df.to_excel(writer, sheet_name=SKU_SHEET, index=False)


def upsert_rows(df, new_rows, key_col, shop, active_keys, now=None):
    """单层表的店内 upsert。

    df         : 该层现有全量表（全店）
    new_rows   : 本次机器产出行 list[dict]（不含 竞店名称/确认来源/更新时间/状态）
    key_col    : "竞品品ID"（SPU层）或 "竞品Sku_id"（SKU层）
    shop       : 本次店名
    active_keys: 该店当前竞品表全部编号集合（在表=有效，不在=已下架）
    返回更新后的该层全量表。
    """
    cols = SPU_COLS if key_col == "竞品品ID" else SKU_COLS
    now = now or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    shop_mask = df["竞店名称"].astype(str) == shop
    if len(df):
        keys = df[key_col].astype(str).str.strip()
        # 1) 店内现有行状态刷新：在现表=有效，不在=已下架（人工行数据列不动，仅状态）
        #    AI预筛行不参与刷新（保住"待抽检"标记，抽检前不计入有效对齐）
        refresh = shop_mask & (df["确认来源"].astype(str) != CONFIRM_AI)
        df.loc[refresh & keys.isin(active_keys), "状态"] = STATUS_ACTIVE
        df.loc[refresh & ~keys.isin(active_keys), "状态"] = STATUS_OFF

    # 2) 受保护行的键（人工确认 + AI预筛）：机器行不得覆盖
    manual_keys = set(
        df.loc[shop_mask & (df["确认来源"].astype(str) != CONFIRM_AUTO), key_col]
        .astype(str).str.strip()
    )
    existing_keys = set(df.loc[shop_mask, key_col].astype(str).str.strip())

    add_rows = []
    for r in new_rows:
        k = str(r.get(key_col, "")).strip()
        if not k or k in manual_keys:
            continue
        row = dict(r)
        row["竞店名称"] = shop
        row["确认来源"] = CONFIRM_AUTO
        row["更新时间"] = now
        row["状态"] = STATUS_ACTIVE if k in active_keys else STATUS_OFF
        if not str(row.get("对应级别") or "").strip():
            row["对应级别"] = derive_level(CONFIRM_AUTO, row.get("匹配层", ""))
        if not str(row.get("来源") or "").strip():
            row["来源"] = str(row.get("匹配层", "")) or CONFIRM_AUTO
        if k in existing_keys:
            mask = shop_mask & (df[key_col].astype(str).str.strip() == k)
            for c in cols:
                if c in row:
                    df.loc[mask, c] = row[c]
        else:
            add_rows.append({c: row.get(c, "") for c in cols})
            existing_keys.add(k)

    if add_rows:
        df = pd.concat([df, pd.DataFrame(add_rows, columns=cols)], ignore_index=True)
    return df


def upsert_auto(df, key_col, shop, row, source, now=None):
    """机器高置信自动行 upsert（确认来源=source，如 向量自动）。

    保护规则：
      - 该键已是 人工确认 / AI预筛 行 → 不覆盖（人与预筛优先）；
      - 该键已是 机器自动 行（更早机器层结论）→ 不覆盖（防跨层翻转）；
      - 该键已是 向量自动 行 → 覆盖数据列（幂等重算）；
      - 否则追加。
    返回更新后的该层全量表与是否写入。"""
    cols = SPU_COLS if key_col == "竞品品ID" else SKU_COLS
    now = now or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    k = str(row.get(key_col, "")).strip()
    if not k:
        return df, False
    shop_mask = df["竞店名称"].astype(str) == shop
    mask = shop_mask & (df[key_col].astype(str).str.strip() == k)
    if mask.any():
        src_now = str(df.loc[mask, "确认来源"].iloc[0] or "")
        if src_now in (CONFIRM_MANUAL, CONFIRM_AI, CONFIRM_AUTO):
            return df, False
        for c in cols:
            if c in row:
                df.loc[mask, c] = str(row.get(c, ""))
        df.loc[mask, ["竞店名称", "确认来源", "更新时间", "状态"]] = \
            [shop, source, now, STATUS_ACTIVE]
        return df, True
    full = {c: str(row.get(c, "")) for c in cols}
    full["竞店名称"] = shop
    full["确认来源"] = source
    full["更新时间"] = now
    full["状态"] = STATUS_ACTIVE
    df = pd.concat([df, pd.DataFrame([full], columns=cols)], ignore_index=True)
    return df, True


def upsert_manual(df, key_col, shop, row, now=None):
    """人工确认行 upsert（确认来源=人工确认，机器重跑不覆盖数据列）。
    row: 含键列与各数据列的 dict。返回更新后的该层全量表。"""
    cols = SPU_COLS if key_col == "竞品品ID" else SKU_COLS
    now = now or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    k = str(row.get(key_col, "")).strip()
    if not k:
        return df
    full = {c: str(row.get(c, "")) for c in cols}
    full["竞店名称"] = shop
    full["确认来源"] = CONFIRM_MANUAL
    full["更新时间"] = now
    full["状态"] = STATUS_ACTIVE
    if not str(full.get("对应级别") or "").strip():
        full["对应级别"] = LEVEL_EQ
    if not str(full.get("来源") or "").strip():
        full["来源"] = "人工"

    shop_mask = df["竞店名称"].astype(str) == shop
    mask = shop_mask & (df[key_col].astype(str).str.strip() == k)
    if mask.any():
        for c in cols:
            df.loc[mask, c] = full[c]
    else:
        df = pd.concat([df, pd.DataFrame([full], columns=cols)], ignore_index=True)
    return df


def upsert_ai(df, key_col, shop, row, now=None):
    """AI预筛行 upsert（确认来源=AI预筛，状态=待抽检，机器重跑不覆盖）。
    人工优先：该键已是人工确认行时不覆盖（AI 只补位，不抢人工结论）。
    row: 含键列与各数据列的 dict。返回更新后的该层全量表。"""
    cols = SPU_COLS if key_col == "竞品品ID" else SKU_COLS
    now = now or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]

    k = str(row.get(key_col, "")).strip()
    if not k:
        return df
    shop_mask = df["竞店名称"].astype(str) == shop
    mask = shop_mask & (df[key_col].astype(str).str.strip() == k)
    if mask.any() and (df.loc[mask, "确认来源"].astype(str) == CONFIRM_MANUAL).any():
        return df  # 人工行优先，AI 不覆盖
    full = {c: str(row.get(c, "")) for c in cols}
    full["竞店名称"] = shop
    full["确认来源"] = CONFIRM_AI
    full["更新时间"] = now
    full["状态"] = STATUS_REVIEW
    if not str(full.get("对应级别") or "").strip():
        full["对应级别"] = LEVEL_REF
    if not str(full.get("来源") or "").strip():
        full["来源"] = "AI预筛"
    if mask.any():
        for c in cols:
            df.loc[mask, c] = full[c]
    else:
        df = pd.concat([df, pd.DataFrame([full], columns=cols)], ignore_index=True)
    return df

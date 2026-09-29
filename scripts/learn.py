# -*- coding: utf-8 -*-
"""
学习闭环：人工确认反哺系统（越用越准）。
  record()        确认后记录样本（学习样本库）+ 品牌/型号补全记录
  apply_brand_fill()  对齐引擎加载竞品表后，用已启用品牌补全填充空品牌/型号
  extract_synonyms()  从确认配对提炼同义词（大模型）-> 待启用清单
  启用机制：同义词写入 config/synonyms.yaml（normalize 生效）；品牌补全标记启用后 align 套用
"""
import os
import json
import re
import time

import pandas as pd
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
SAMPLES_FILE = os.path.join(OUTPUT, "学习样本库.xlsx")
BRAND_FILL_FILE = os.path.join(OUTPUT, "品牌补全库.xlsx")
SYN_PENDING_FILE = os.path.join(BASE, "config", "synonyms_pending.yaml")
SYNONYMS_FILE = os.path.join(BASE, "config", "synonyms.yaml")
KEYWORD_LIB_FILE = os.path.join(OUTPUT, "自家关键词库.xlsx")

SAMPLES_COLS = ["时间", "竞店名称", "竞品Sku_id", "竞品商品名", "自家SKUID", "自家商品名",
                "确认来源", "竞品品牌", "竞品型号", "自家品牌", "自家型号"]
BRAND_COLS = ["竞品Sku_id", "竞店名称", "推断品牌", "推断型号", "来源", "时间", "状态"]
KEYWORD_COLS = ["关键词", "品组ID", "品名", "类目", "来源", "时间", "状态"]

# 停用词（营销/通用/特性词；与 align.KEYWORD_STOP 一致，独立存放避免循环导入）
KEYWORD_STOP = {"正品", "包邮", "家用", "通用", "专用", "新款", "批发", "加厚", "高配",
                "升级", "加强", "精品", "促销", "特价", "优惠", "赠品", "免费", "上门",
                "安装", "维修", "厂家", "直销", "现货", "下单", "客服", "联系", "咨询",
                "备注", "型号", "规格", "颜色", "款式", "左右", "上下", "系列", "套装",
                "闪购", "美团", "闪电仓", "热销", "爆款", "断货", "补货", "到货"}


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _read_df(path, cols):
    if os.path.exists(path):
        try:
            df = pd.read_excel(path, dtype=str).fillna("")
            for c in cols:
                if c not in df.columns:
                    df[c] = ""
            return df
        except Exception:
            pass
    return pd.DataFrame(columns=cols)


def _append_row(path, row_dict, cols):
    df = _read_df(path, cols)
    df = pd.concat([df, pd.DataFrame([{c: _clean(row_dict.get(c, "")) for c in cols}])],
                   ignore_index=True)
    df.to_excel(path, index=False)


def record(shop, comp_sku, comp_name, own_sku, own_name, source="人工确认",
           comp_brand="", comp_model="", own_brand="", own_model=""):
    """确认后调用：写学习样本库；竞品缺品牌/型号时写品牌补全库（待启用）。"""
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        _append_row(SAMPLES_FILE, {
            "时间": now, "竞店名称": shop, "竞品Sku_id": comp_sku, "竞品商品名": comp_name,
            "自家SKUID": own_sku, "自家商品名": own_name, "确认来源": source,
            "竞品品牌": comp_brand, "竞品型号": comp_model,
            "自家品牌": own_brand, "自家型号": own_model,
        }, SAMPLES_COLS)
    except Exception:
        pass
    # 品牌补全：竞品缺品牌/型号且自家有 -> 待启用
    if (not comp_brand or not comp_model) and (own_brand or own_model):
        try:
            _append_row(BRAND_FILL_FILE, {
                "竞品Sku_id": comp_sku, "竞店名称": shop,
                "推断品牌": own_brand, "推断型号": own_model,
                "来源": source + "（" + comp_name[:40] + "）",
                "时间": now, "状态": "待启用",
            }, BRAND_COLS)
        except Exception:
            pass
    return True


def load_brand_fill(active_only=True):
    """读取品牌补全库；active_only=True 仅取状态=启用。返回 {竞品Sku_id: {品牌, 型号}}。"""
    df = _read_df(BRAND_FILL_FILE, BRAND_COLS)
    if not len(df):
        return {}
    if active_only:
        df = df[df["状态"].astype(str) == "启用"]
    out = {}
    for _, r in df.iterrows():
        sku = _clean(r.get("竞品Sku_id"))
        if sku and sku.lower() != "nan":
            out.setdefault(sku, {})
            if r.get("推断品牌"):
                out[sku]["品牌"] = _clean(r.get("推断品牌"))
            if r.get("推断型号"):
                out[sku]["型号"] = _clean(r.get("推断型号"))
    return out


def apply_brand_fill(dfc):
    """对齐引擎用：用已启用品牌补全，填充竞品行缺失的品牌/型号。返回填充行数。"""
    fill = load_brand_fill(active_only=True)
    if not fill or len(dfc) == 0:
        return 0
    n = 0
    sku_col = "Sku_id" if "Sku_id" in dfc.columns else "sku_id"
    if sku_col not in dfc.columns:
        return 0
    for i, row in dfc.iterrows():
        f = fill.get(_clean(row.get(sku_col)))
        if not f:
            continue
        if "品牌" in f and not _clean(row.get("品牌（抽取）")):
            dfc.at[i, "品牌（抽取）"] = f["品牌"]
            n += 1
        if "型号" in f and not _clean(row.get("型号")):
            dfc.at[i, "型号"] = f["型号"]
    return n


def extract_synonyms(limit=30, config=None):
    """从学习样本库提炼同义词（大模型）：返回新增待启用词对列表。"""
    from ai_analysis import load_config
    config = config or load_config()
    api_key = (config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("未配置 AI 接口密钥，请先在「选店分析」页配置 DeepSeek key")
    df = _read_df(SAMPLES_FILE, SAMPLES_COLS)
    if not len(df):
        raise RuntimeError("学习样本库为空，先确认几笔配对再提炼同义词")
    pairs = []
    for _, r in df.tail(limit).iterrows():
        c, o = _clean(r.get("竞品商品名")), _clean(r.get("自家商品名"))
        if c and o:
            pairs.append(f"竞品：{c}\n自家：{o}")
    prompt = ("以下是已人工确认的「竞品-自家」配对（同一商品），请找出两边名称中「对应但写法不同」的词对"
              "（如 竞品「尖泡」 vs 自家「蜡烛灯」；品牌词/品名词/规格词都算）。\n"
              "只输出 JSON 数组，不要输出其他文字：\n"
              '[{"variant": "竞品写法", "standard": "自家标准写法"}, ...]\n\n' + "\n\n".join(pairs))

    import requests
    from ai_client import call_llm
    content = call_llm("learn_synonyms",
                       [{"role": "user", "content": prompt}],
                       config=config, timeout=120,
                       temperature=0.2, max_tokens=800)
    try:
        words = json.loads(content)
    except (ValueError, TypeError):
        s, e = content.find("["), content.rfind("]")
        words = json.loads(content[s:e + 1]) if 0 <= s < e else []
    if not isinstance(words, list):
        words = []

    # 去重（与现有 synonyms.yaml 和待启用清单比对）后写入待启用
    existing = set()
    if os.path.exists(SYNONYMS_FILE):
        try:
            import yaml
            with open(SYNONYMS_FILE, "r", encoding="utf-8") as f:
                existing.update((yaml.safe_load(f) or {}).keys())
        except Exception:
            pass
    pending = load_syn_pending()
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added = []
    for w in words:
        if not isinstance(w, dict):
            continue
        var = _clean(w.get("variant"))
        std = _clean(w.get("standard"))
        if not var or not std or var == std:
            continue
        if var in existing or any(p["variant"] == var for p in pending):
            continue
        pending.append({"variant": var, "standard": std, "time": now, "status": "待启用"})
        added.append({"variant": var, "standard": std})
    save_syn_pending(pending)
    return added


def load_syn_pending():
    if os.path.exists(SYN_PENDING_FILE):
        try:
            import yaml
            with open(SYN_PENDING_FILE, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or []
            return data if isinstance(data, list) else []
        except Exception:
            pass
    return []


def save_syn_pending(pending):
    import yaml
    os.makedirs(os.path.dirname(SYN_PENDING_FILE), exist_ok=True)
    with open(SYN_PENDING_FILE, "w", encoding="utf-8") as f:
        yaml.safe_dump(pending, f, allow_unicode=True, sort_keys=False)


def enable_synonym(variant, standard, enable=True):
    """启用/忽略同义词：启用写入 synonyms.yaml（normalize 生效）。"""
    pending = load_syn_pending()
    for p in pending:
        if p.get("variant") == variant:
            p["status"] = "启用" if enable else "忽略"
            break
    save_syn_pending(pending)
    if enable:
        import yaml
        syn = {}
        if os.path.exists(SYNONYMS_FILE):
            with open(SYNONYMS_FILE, "r", encoding="utf-8") as f:
                syn = yaml.safe_load(f) or {}
        syn[variant] = standard
        with open(SYNONYMS_FILE, "w", encoding="utf-8") as f:
            yaml.safe_dump(syn, f, allow_unicode=True, sort_keys=False)


def enable_brand_fill(sku, enable=True):
    """启用/忽略品牌补全记录。"""
    df = _read_df(BRAND_FILL_FILE, BRAND_COLS)
    if not len(df):
        return
    mask = df["竞品Sku_id"].astype(str) == str(sku)
    df.loc[mask, "状态"] = "启用" if enable else "忽略"
    df.to_excel(BRAND_FILL_FILE, index=False)


def learn_status():
    """学习状态统计：待启用同义词数/待启用品牌补全数/样本数。"""
    return {
        "samples": len(_read_df(SAMPLES_FILE, SAMPLES_COLS)),
        "pending_synonyms": sum(1 for p in load_syn_pending() if p.get("status") == "待启用"),
        "pending_brand": int((_read_df(BRAND_FILL_FILE, BRAND_COLS)["状态"] == "待启用").sum())
        if len(_read_df(BRAND_FILL_FILE, BRAND_COLS)) else 0,
        "keywords": len(_read_df(KEYWORD_LIB_FILE, KEYWORD_COLS)),
        "keyword_groups": int(_read_df(KEYWORD_LIB_FILE, KEYWORD_COLS)["品组ID"].nunique())
        if len(_read_df(KEYWORD_LIB_FILE, KEYWORD_COLS)) else 0,
    }


def add_keyword(own_name, spuid, cat="", source="确认反哺"):
    """确认反哺：从确认过的自家品名提取词段写入关键词库（零成本，词库越用越准）。"""
    import re as _re
    if not own_name or not spuid:
        return 0
    words = set()
    for seg in _re.findall(r"[\u4e00-\u9fff]{2,}", str(own_name)):
        if seg in KEYWORD_STOP or len(seg) < 2:
            continue
        words.add(seg)
    if not words:
        return 0
    df = _read_df(KEYWORD_LIB_FILE, KEYWORD_COLS)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    added = 0
    for w in sorted(words):
        if len(df) and ((df["关键词"] == w) & (df["品组ID"] == spuid)).any():
            continue
        df = pd.concat([df, pd.DataFrame([{"关键词": w, "品组ID": spuid, "品名": own_name,
                                           "类目": cat, "来源": source,
                                           "时间": now, "状态": "启用"}])], ignore_index=True)
        added += 1
    if added:
        df.to_excel(KEYWORD_LIB_FILE, index=False)
    return added

# -*- coding: utf-8 -*-
"""
阶段2 跨店商品对齐引擎 (Task 2.1-2.5)
把竞品表与自家表的商品对齐到同一物理商品。

4 层匹配，逐层收敛、避免重复：
  第1层 条码精确匹配          置信度 0.95  自动确认
  第2层 强组合键(类目+品牌+型号+规格) 置信度 0.90  自动确认
  第3层 弱组合键(类目+品牌+型号, 规格不一致) 置信度 0.70  人工池
  第4层 模糊兜底(TF-IDF + 图片pHash)   置信度 <0.60  人工池

输入:
  output/竞品_归一化后.xlsx
  output/自家_归一化后.xlsx
输出:
  output/对齐_第1层_条码.xlsx
  output/对齐_第2层_强组合.xlsx
  output/对齐_第3层_弱组合_人工池.xlsx
  output/对齐_第4层_模糊_人工池.xlsx
  output/对齐_人工确认池.xlsx
  output/对齐_汇总.xlsx
"""
import os
import re
import sys
import io
import time
import hashlib

_HERE_DIR = os.path.dirname(os.path.abspath(__file__))
if _HERE_DIR not in sys.path:
    sys.path.insert(0, _HERE_DIR)
import text_norm as TN

import json
from datetime import datetime

import pandas as pd
import numpy as np
import requests
from PIL import Image
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.metrics.pairwise import cosine_similarity

# 竞店标准库中文表头 -> 英文内部名（读标准库链路时转回）
from import_adapter import CN_STD
from summary_library import (DEFAULT_SHOP as LIB_DEFAULT_SHOP,
                             LIBRARY_FILE, load_library, save_library, upsert_rows)

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMP_NORM = os.path.join(BASE, "output", "竞品_归一化后.xlsx")
OWN_NORM = os.path.join(BASE, "output", "自家_归一化后.xlsx")
OUT_DIR = os.path.join(BASE, "output")
os.makedirs(OUT_DIR, exist_ok=True)

# 置信度常量
CONF_BARCODE = 0.95
CONF_STRONG = 0.90
CONF_WEAK = 0.70
CONF_FUZZY_CAP = 0.59      # 第4层置信度统一 <0.60
FUZZY_THRESHOLD = 0.50     # 第4层融合分数阈值（用户标准：宁缺毋滥，不做硬凑匹配）
IMAGE_TIMEOUT = 5          # 图片下载超时(秒)
PHASH_THRESHOLD = 8        # pHash 汉明距离阈值（≤8 视为图片相似，放宽自≤5）
URL_MATCH_BONUS = 0.10     # 图片url同源时的置信度加成
BRAND_PENALTY = 0.5        # 品牌硬惩罚：双方品牌明确且不一致 -> 相似度/置信度×0.5（跨品牌误配防线）

# 规格拼接字段顺序
SPEC_FIELDS = ["颜色档位", "开关档位", "功率", "容量", "尺寸", "电压"]


# ===========================================================================
# 数据加载
# ===========================================================================
def load_comp_norm(path):
    """读取竞品归一化表并做列名兼容（中英文列名统一为中文内部名）。
    条码/ID 列强制 str 防精度丢失。load_data 与 export_overview 共用。"""
    dfc = pd.read_excel(
        path,
        dtype={"条形码": str, "商品ID": str, "Sku_id": str,
               "barcode": str, "comp_id": str, "sku_id": str},
    )
    # 标准库英文列名 -> 中文标准列名（对齐旧链路内部约定）
    dfc = dfc.rename(columns={k: v for k, v in CN_STD.items() if k in dfc.columns})
    COMP_RENAME = {
        "comp_id": "商品ID", "sku_id": "Sku_id", "name": "商品名字",
        "barcode": "条形码", "spec": "规格", "image_url": "图片url",
        "category1": "一级类", "category2": "二级类", "category3": "美团三级分类",
        "sales": "月售", "price": "原价", "discount_price": "折后价",
        "god_price": "神价", "stock": "库存",
    }
    dfc = dfc.rename(columns={k: v for k, v in COMP_RENAME.items() if k in dfc.columns})
    # 兜底：若列被读成非 str，再强制转 str
    for c in ["条形码", "商品ID", "Sku_id"]:
        if c in dfc.columns:
            dfc[c] = dfc[c].astype(str)
    for c in dfc.columns:
        if dfc[c].dtype == object:
            dfc[c] = dfc[c].fillna("")
    return dfc


def load_data():
    """读取归一化后两表。条码/ID 列强制 str 防止 float 丢精度。
    竞品表兼容两种列名：旧单店链路（中文列）与竞店标准库链路（英文列）。"""
    dfc = load_comp_norm(COMP_NORM)
    # 学习闭环：已启用品牌补全（人工确认反哺）填充竞品缺失品牌/型号
    try:
        from learn import apply_brand_fill
        n_fill = apply_brand_fill(dfc)
        if n_fill:
            print(f"[学习] 品牌补全填充 {n_fill} 行竞品缺失品牌/型号")
    except Exception:
        pass
    dfo = pd.read_excel(
        OWN_NORM,
        dtype={"商品条码": str, "SPUID": str, "SKUID": str},
    )
    for c in ["商品条码", "SPUID", "SKUID"]:
        if c in dfo.columns:
            dfo[c] = dfo[c].astype(str)

    # 空值统一（竞品表已在 load_comp_norm 内处理）
    for c in dfo.columns:
        if dfo[c].dtype == object:
            dfo[c] = dfo[c].fillna("")

    # 竞品主键 Sku_id，自家主键 SKUID（落空时退化到行号）
    dfc["__cid"] = dfc["Sku_id"].where(dfc["Sku_id"].astype(str).str.strip() != "", dfc["商品ID"])
    dfo["__oid"] = dfo["SKUID"].where(dfo["SKUID"].astype(str).str.strip() != "", dfo["SPUID"])
    return dfc, dfo


# ===========================================================================
# 条码归一
# ===========================================================================
def norm_barcode(b):
    """条码归一（2026-08-13 统一口径，与 build_comp_library.clean_barcode 一致）：
    提取纯数字；13 位保留原样（不 lstrip 零——0 开头是合法 EAN/UPC）；
    12 位 UPC-A 补前导 0 转 EAN-13；其余长度（含 19 位美团内码）返回 ''。
    空值返回 ''。"""
    if b is None:
        return ""
    if isinstance(b, float):
        if b != b:              # NaN
            return ""
        if b.is_integer():
            b = int(b)
    s = str(b).strip()
    if s == "" or s.lower() in ("nan", "none"):
        return ""
    digits = re.sub(r"\D", "", s)
    if len(digits) == 12:
        digits = "0" + digits   # UPC-A -> EAN-13
    return digits if len(digits) == 13 else ""


def build_spec_key(row):
    """规格拼接键：颜色档位+开关档位+功率+容量+尺寸+电压，空值跳过。"""
    parts = []
    for f in SPEC_FIELDS:
        v = row.get(f, "")
        if isinstance(v, float):
            v = "" if np.isnan(v) else str(v)
        v = str(v).strip()
        if v and v.lower() != "nan":
            parts.append(v)
    return "|".join(parts)


def norm_text(s):
    """文本归一：小写、去非字母数字汉字，便于 TF-IDF。"""
    if not isinstance(s, str):
        s = "" if (s is None or (isinstance(s, float) and np.isnan(s))) else str(s)
    s = s.lower()
    s = re.sub(r"[\s\-_/\\]+", "", s)
    return s


def _clean(v):
    """统一清洗单元格值：None/NaN/'nan'/'None' -> ''，其余转 str。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if np.isnan(v) else str(v)
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _brand_conflict(b1, b2):
    """品牌硬惩罚判定：双方品牌均明确（非空且非'无'）且不一致 -> True。
    '无'=无品牌（项目约定），不参与冲突判定；空品牌（数据缺失）也不惩罚。"""
    b1 = _clean(b1).upper()
    b2 = _clean(b2).upper()
    if not b1 or not b2 or b1 == "无" or b2 == "无":
        return False
    return b1 != b2


def spec_jaccard(s1, s2):
    """规格键的 token Jaccard 相似度。token 按 | 和 / 切分。"""
    t1 = set(re.split(r"[|/]", s1)) - {""}
    t2 = set(re.split(r"[|/]", s2)) - {""}
    if not t1 and not t2:
        return 1.0
    if not t1 or not t2:
        return 0.0
    return len(t1 & t2) / len(t1 | t2)


# ===========================================================================
# 感知哈希 pHash（PIL + hashlib 手写，不依赖 imagehash）
# ===========================================================================
def phash(img, hash_size=8, highfreq_factor=4):
    """对 PIL.Image 计算感知哈希，返回 16 位十六进制字符串。"""
    img_size = hash_size * highfreq_factor
    image = img.convert("L").resize((img_size, img_size), Image.LANCZOS)
    pixels = np.asarray(image, dtype=np.float64)
    dct = _dct2(pixels)
    dctlow = dct[:hash_size, :hash_size]
    med = np.median(dctlow.flatten()[1:]) if dctlow.size > 1 else 0.0
    bits = (dctlow.flatten() > med)
    # bits -> hex
    h = 0
    for b in bits:
        h = (h << 1) | int(b)
    return format(h, "0{}x".format(hash_size * hash_size // 4))


def _dct2(matrix):
    """简易 2D DCT-II（用于 pHash）。"""
    m = np.asarray(matrix, dtype=np.float64)
    n = m.shape[0]
    # 1D DCT 沿 axis=0
    def dct1d(v):
        N = len(v)
        out = np.zeros(N)
        for k in range(N):
            s = 0.0
            for t in range(N):
                s += v[t] * np.cos(np.pi * (2 * t + 1) * k / (2 * N))
            out[k] = s
        return out
    rows = np.array([dct1d(m[i, :]) for i in range(n)])
    cols = np.array([dct1d(rows[:, j]) for j in range(n)]).T
    return cols


def hamming_hex(h1, h2):
    """两个十六进制字符串的汉明距离。"""
    n = min(len(h1), len(h2))
    d = 0
    for i in range(n):
        x = int(h1[i], 16) ^ int(h2[i], 16)
        d += bin(x).count("1")
    d += abs(len(h1) - len(h2)) * 4
    return d


def download_image_phash(url, cache):
    """下载图片并返回 pHash。失败返回 ''。带缓存与超时。"""
    if not isinstance(url, str) or not url.startswith("http"):
        return ""
    if url in cache:
        return cache[url]
    ph = ""
    try:
        r = requests.get(url, timeout=IMAGE_TIMEOUT, stream=True)
        r.raise_for_status()
        img = Image.open(io.BytesIO(r.content))
        ph = phash(img)
    except Exception:
        ph = ""
    cache[url] = ph
    return ph


def image_similarity(url_c, url_o, cache):
    """返回图片相似度 0~1。任一无图/失败返回 None。"""
    pc = download_image_phash(url_c, cache)
    po = download_image_phash(url_o, cache)
    if not pc or not po:
        return None
    d = hamming_hex(pc, po)
    # hash_size=8 -> 64 bits
    return 1.0 - d / 64.0


def check_image_url_match(comp_url, own_url):
    """检查竞品图片url与自家图片url是否同源（同域名+同文件名）。
    提取域名+文件名（去掉query参数）。
    返回 (是否同源: bool, 状态标记: "是"/"否"/"无url")。
    任一为空/非法 -> (False, "无url")。
    """
    from urllib.parse import urlparse

    def _extract(url):
        if not isinstance(url, str):
            return None
        s = url.strip()
        if not s or not s.startswith("http"):
            return None
        try:
            parsed = urlparse(s)
            domain = (parsed.netloc or "").lower()
            path = parsed.path or ""
            # 文件名：取 path 最后一段，去 query/fragment
            filename = path.rsplit("/", 1)[-1] if path else ""
            filename = filename.split("?")[0].split("#")[0]
            if not domain or not filename:
                return None
            return (domain, filename)
        except Exception:
            return None

    c = _extract(comp_url)
    o = _extract(own_url)
    if c is None or o is None:
        return False, "无url"
    if c[0] == o[0] and c[1] == o[1]:
        return True, "是"
    return False, "否"


def compare_images_phash(comp_phash, own_phash, threshold=8):
    """比较两个 pHash 十六进制字符串。
    返回 (是否相似: bool, 汉明距离: int/None)。
    距离≤threshold -> (True, distance)
    距离>threshold 或任一为空 -> (False, distance)
    """
    if not comp_phash or not own_phash:
        return False, None
    d = hamming_hex(comp_phash, own_phash)
    return (d <= threshold, d)


# ===========================================================================
# 输出行构造
# ===========================================================================
COMP_KEEP = ["商品ID", "Sku_id", "商品名字", "条形码", "规格", "图片url",
             "品牌（抽取）", "型号", "功率", "容量", "尺寸", "电压",
             "颜色档位", "开关档位", "类目三级_归一"]
OWN_KEEP = ["SPUID", "SKUID", "商品名称", "商品条码", "规格名称",
            "品牌（抽取）", "型号", "功率", "容量", "尺寸", "电压",
            "颜色档位", "开关档位", "类目三级_归一", "美团闪购渠道售价"]


def _get(row, cols, src_map):
    """从原始行取列值，统一列名到输出列名。"""
    out = {}
    for c in cols:
        out[src_map.get(c, c)] = row.get(c, "")
    return out


def make_match_row(crow, orow, conf, basis, layer):
    """构造一条匹配输出行。所有值经 _clean 清洗，避免 NaN 残留。"""
    row = {}
    # 竞品侧
    row["竞品品ID"] = _clean(crow.get("商品ID", ""))
    row["竞品Sku_id"] = _clean(crow.get("Sku_id", ""))
    row["竞品商品名"] = _clean(crow.get("商品名字", "")) or _clean(crow.get("清洗后商品名", ""))
    row["竞品条形码"] = _clean(crow.get("条形码", ""))
    row["竞品品牌"] = _clean(crow.get("品牌（抽取）", ""))
    row["竞品型号"] = _clean(crow.get("型号", ""))
    row["竞品类目"] = _clean(crow.get("类目三级_归一", ""))
    row["竞品规格键"] = build_spec_key(crow)
    # 自家侧
    row["自家SPUID"] = _clean(orow.get("SPUID", ""))
    row["自家SKUID"] = _clean(orow.get("SKUID", ""))
    row["自家商品名"] = _clean(orow.get("商品名称", "")) or _clean(orow.get("清洗后商品名", ""))
    row["自家商品条码"] = _clean(orow.get("商品条码", ""))
    row["自家品牌"] = _clean(orow.get("品牌（抽取）", ""))
    row["自家型号"] = _clean(orow.get("型号", ""))
    row["自家类目"] = _clean(orow.get("类目三级_归一", ""))
    row["自家规格键"] = build_spec_key(orow)
    # 匹配元信息
    row["匹配层"] = layer
    row["匹配依据"] = basis
    row["置信度"] = conf
    return row


# ===========================================================================
# 价格字段处理与校验
# ===========================================================================
def _to_float(v):
    """单元格转 float。None/NaN/空/非数字 -> None。
    兼容中文前缀/后缀格式：'月售2'/'12人想买'/'2万'/'1.5k'/'100+'/'1,200'
    （通用正则提取首数字——不同店/不同期美团导出 sales 格式不一，见 2026-08-18 教训）。"""
    import re as _re
    if v is None:
        return None
    if isinstance(v, float):
        return None if np.isnan(v) else v
    s = str(v).strip()
    if s == "" or s.lower() in ("nan", "none", "-"):
        return None
    try:
        f = float(s)
        return None if np.isnan(f) else f
    except (ValueError, TypeError):
        pass
    m = _re.search(r"(\d[\d,]*(?:\.\d+)?)\s*([万亿千kKmM])?", s)
    if not m:
        return None
    f = float(m.group(1).replace(",", ""))
    u = m.group(2)
    if u in ("万", "m", "M"):
        f *= 10000
    elif u in ("千", "k", "K"):
        f *= 1000
    return f


def build_price_lookup(dfc, dfo):
    """构造价格查找表。
    comp: Sku_id -> (神价, 折后价, 原价)
    own : SKUID  -> 美团闪购渠道售价
    """
    comp = {}
    for _, r in dfc.iterrows():
        sid = _clean(r.get("Sku_id"))
        if sid:
            comp[sid] = (r.get("神价"), r.get("折后价"), r.get("原价"))
    own = {}
    for _, r in dfo.iterrows():
        sid = _clean(r.get("SKUID"))
        if sid:
            own[sid] = r.get("美团闪购渠道售价")
    return comp, own


def get_comp_price(row, comp_lookup):
    """竞品到手价: 神价 > 折后价 > 原价（取第一个有值且>0的）。"""
    sid = _clean(row.get("竞品Sku_id"))
    if not sid or sid not in comp_lookup:
        return None
    shen, zhe, yuan = comp_lookup[sid]
    for v in (shen, zhe, yuan):
        f = _to_float(v)
        if f is not None and f > 0:
            return f
    return None


def get_own_price(row, own_lookup):
    """自家售价: 美团闪购渠道售价。"""
    sid = _clean(row.get("自家SKUID"))
    if not sid or sid not in own_lookup:
        return None
    return _to_float(own_lookup[sid])


# ===========================================================================
# 冲突消解（同一竞品Sku_id命中多个自家品时，保留置信度最高的）
# ===========================================================================
CONFLICT_LOG_COLUMNS = [
    "竞品品ID", "竞品Sku_id", "竞品商品名",
    "被丢弃的自家SKUID", "被丢弃的匹配层", "被丢弃的置信度",
    "保留的自家SKUID", "保留的匹配层", "保留的置信度",
    "丢弃原因",
]


def resolve_conflicts(all_matches_df):
    """对同一竞品Sku_id的多匹配做消解，保留置信度最高的一条。
    Sku_id 为空的行回退按 竞品品ID 分组（保留旧行为）。
    返回 (消解后 DataFrame, 冲突日志 DataFrame)。

    粒度说明：按规格（Sku_id）而非商品（品ID）去重——同一竞品商品下
    多个规格各配各的自家品属合法一对多，不消解；只有同一个竞品规格
    命中多个自家品时才保留置信度最高者。不同竞品命中同一自家品不消解。
    """
    if len(all_matches_df) == 0:
        return all_matches_df, pd.DataFrame(columns=CONFLICT_LOG_COLUMNS)

    # 分组键：优先 竞品Sku_id，落空回退 竞品品ID
    df = all_matches_df.copy()
    df["__dedup_key"] = df["竞品Sku_id"].astype(str).str.strip()
    empty_sku = df["__dedup_key"] == ""
    df.loc[empty_sku, "__dedup_key"] = df.loc[empty_sku, "竞品品ID"].astype(str).str.strip()

    has_id = df["__dedup_key"] != ""
    sub = df[has_id].copy()
    rest = df[~has_id].copy()

    conflict_log = []
    kept_frames = []

    # 按 竞品Sku_id（或回退的品ID）分组
    for cid, grp in sub.groupby("__dedup_key", dropna=False):
        if len(grp) <= 1:
            kept_frames.append(grp)
            continue
        # 多匹配：按置信度降序，保留 top1
        grp_sorted = grp.sort_values("置信度", ascending=False)
        keep_row = grp_sorted.iloc[0]
        kept_frames.append(grp_sorted.head(1))
        # 其余记入冲突日志
        for _, drop_row in grp_sorted.iloc[1:].iterrows():
            try:
                keep_layer = int(keep_row.get("匹配层"))
            except (ValueError, TypeError):
                keep_layer = keep_row.get("匹配层")
            try:
                drop_layer = int(drop_row.get("匹配层"))
            except (ValueError, TypeError):
                drop_layer = drop_row.get("匹配层")
            conflict_log.append({
                "竞品品ID": _clean(drop_row.get("竞品品ID")),
                "竞品Sku_id": _clean(drop_row.get("竞品Sku_id")),
                "竞品商品名": _clean(drop_row.get("竞品商品名")),
                "被丢弃的自家SKUID": _clean(drop_row.get("自家SKUID")),
                "被丢弃的匹配层": drop_layer,
                "被丢弃的置信度": drop_row.get("置信度"),
                "保留的自家SKUID": _clean(keep_row.get("自家SKUID")),
                "保留的匹配层": keep_layer,
                "保留的置信度": keep_row.get("置信度"),
                "丢弃原因": (
                    f"同一竞品Sku_id命中多个自家品，保留置信度更高者"
                    f"(保留层{keep_layer} vs 被丢弃层{drop_layer})"
                ),
            })

    if kept_frames:
        resolved_sub = pd.concat(kept_frames, ignore_index=True)
    else:
        resolved_sub = pd.DataFrame(columns=sub.columns)
    resolved = pd.concat([resolved_sub, rest], ignore_index=True)
    resolved = resolved.drop(columns=["__dedup_key"])
    conflict_df = pd.DataFrame(conflict_log, columns=CONFLICT_LOG_COLUMNS)
    return resolved, conflict_df


# ===========================================================================
# 价格合理性校验（仅对第3/4层，价差率超±50%标记异常）
# ===========================================================================
PRICE_DIFF_THRESHOLD = 0.50  # 价差率超 ±50% 视为异常


def validate_prices(df, comp_lookup, own_lookup):
    """为已匹配结果计算价差率并标记价格校验状态。
    第1/2层(自动确认) -> 状态="未校验"
    第3/4层           -> 状态="正常"/"异常"/"未校验"(价格缺失)
    返回添加了 价差率、价格校验状态 列的 DataFrame。
    """
    if len(df) == 0:
        df = df.copy()
        df["价差率"] = []
        df["价格校验状态"] = []
        return df

    df = df.copy()
    rates = []
    statuses = []
    for _, row in df.iterrows():
        try:
            layer = int(row.get("匹配层"))
        except (ValueError, TypeError):
            layer = 0
        cp = get_comp_price(row, comp_lookup)
        op = get_own_price(row, own_lookup)
        if cp is None or op is None or op == 0:
            rate = None
        else:
            rate = (cp - op) / op
        rates.append(round(rate, 4) if rate is not None else "")
        # 第1/2层不校验
        if layer in (1, 2):
            statuses.append("未校验")
        else:
            if rate is None:
                statuses.append("未校验")
            elif abs(rate) > PRICE_DIFF_THRESHOLD:
                statuses.append("异常")
            else:
                statuses.append("正常")
    df["价差率"] = rates
    df["价格校验状态"] = statuses
    return df


# ===========================================================================
# 第1层 条码精确匹配
# ===========================================================================
def layer1_barcode(dfc, dfo):
    """条码归一后精确匹配。返回 (匹配行列表, 已匹配竞品行索引集合, 已匹配自家__oid集合)。"""
    dfc = dfc.copy()
    dfo = dfo.copy()
    dfc["__bc"] = dfc["条形码"].apply(norm_barcode)
    dfo["__bc"] = dfo["商品条码"].apply(norm_barcode)

    # 自家: 条码 -> 代表 SKU（同条码多 SKU 取月销最高者为代表，1:1 收敛）。
    # 原实现按 __oid 排序取首个（无业务意义）；月销高的规格通常是主打规格，
    # 匹配后人工确认页展示的代表 SKU 更合理。
    def _bc_sales(r):
        try:
            v = str(r.get("月销量") or r.get("月售") or 0).replace(",", "").strip()
            return float(v) if v and v.lower() not in ("nan", "none") else 0.0
        except (ValueError, TypeError):
            return 0.0

    dfo_bc = dfo[dfo["__bc"] != ""].copy()
    dfo_bc["__bc_sales"] = dfo_bc.apply(_bc_sales, axis=1)
    own_bc = (dfo_bc.sort_values("__bc_sales", ascending=False)
              .drop_duplicates("__bc", keep="first")
              .set_index("__bc"))

    matches = []
    used_cidx = set()
    used_oid = set()
    for i, crow in dfc.iterrows():
        bc = crow["__bc"]
        if bc == "" or bc not in own_bc.index:
            continue
        orow = own_bc.loc[bc]
        oid = orow["__oid"]
        if oid in used_oid:
            continue
        matches.append(make_match_row(crow, orow, CONF_BARCODE, "条码", 1))
        used_cidx.add(i)
        used_oid.add(oid)
    return matches, used_cidx, used_oid


# ===========================================================================
# 第2层 强组合键匹配（类目+品牌+型号+规格）
# ===========================================================================
def layer2_strong(dfc, dfo, used_cidx, used_oid):
    """强组合键匹配。先按 类目+品牌 分组，组内 型号+规格 全命中。"""
    dfc = dfc.copy()
    dfo = dfo.copy()
    dfc["__spec"] = dfc.apply(build_spec_key, axis=1)
    dfo["__spec"] = dfo.apply(build_spec_key, axis=1)

    # 自家可用候选（排除已被第1层匹配的自家）
    own_avail = dfo[~dfo["__oid"].isin(used_oid)].copy()

    matches = []
    # 分组键：类目三级_归一 + 品牌（抽取）
    dfc_rem = dfc[~dfc.index.isin(used_cidx)].copy()
    group_key = ["类目三级_归一", "品牌（抽取）"]

    for (cat, brand), gc in dfc_rem.groupby(group_key, dropna=False):
        if cat == "" and brand == "":
            continue
        go = own_avail[
            (own_avail["类目三级_归一"] == cat)
            & (own_avail["品牌（抽取）"] == brand)
        ]
        if go.empty:
            continue
        # 自家按 (型号, 规格) 索引
        go_idx = {}
        for j, orow in go.iterrows():
            key = (str(orow["型号"]).strip(), str(orow["__spec"]).strip())
            go_idx.setdefault(key, []).append((j, orow))
        for i, crow in gc.iterrows():
            ck = (str(crow["型号"]).strip(), str(crow["__spec"]).strip())
            if crow["型号"] == "" or crow["__spec"] == "":
                continue  # 强组合要求四要素全命中
            cands = go_idx.get(ck, [])
            for j, orow in cands:
                oid = orow["__oid"]
                if oid in used_oid:
                    continue
                matches.append(
                    make_match_row(crow, orow, CONF_STRONG,
                                   "类目+品牌+型号+规格", 2)
                )
                used_cidx.add(i)
                used_oid.add(oid)
                break
    return matches, used_cidx, used_oid


# ===========================================================================
# 第3层 弱组合键匹配（类目+品牌+型号，规格不一致）
# ===========================================================================
def layer3_weak(dfc, dfo, used_cidx, used_oid):
    """同系列不同规格：型号相同但规格键不同 -> 人工池。
    为避免组合爆炸，每个竞品行只取规格最接近的 1 个自家候选（Jaccard top1）。"""
    dfc = dfc.copy()
    dfo = dfo.copy()
    dfc["__spec"] = dfc.apply(build_spec_key, axis=1)
    dfo["__spec"] = dfo.apply(build_spec_key, axis=1)

    # 自家候选：排除已被第1+2层确认匹配的（避免候选池重复已确认项）
    own_avail = dfo[~dfo["__oid"].isin(used_oid)].copy()
    dfc_rem = dfc[~dfc.index.isin(used_cidx)].copy()

    candidates = []
    group_key = ["类目三级_归一", "品牌（抽取）"]
    for (cat, brand), gc in dfc_rem.groupby(group_key, dropna=False):
        if cat == "" and brand == "":
            continue
        go = own_avail[
            (own_avail["类目三级_归一"] == cat)
            & (own_avail["品牌（抽取）"] == brand)
        ]
        if go.empty:
            continue
        # 自家按型号索引
        go_model = {}
        for j, orow in go.iterrows():
            m = _clean(orow["型号"])
            if m:
                go_model.setdefault(m, []).append(orow)
        for i, crow in gc.iterrows():
            cm = _clean(crow["型号"])
            if cm == "":
                continue
            cands = go_model.get(cm, [])
            if not cands:
                continue
            c_spec = _clean(crow["__spec"])
            # 仅保留规格不同的候选
            diff = [orow for orow in cands if _clean(orow["__spec"]) != c_spec]
            if not diff:
                continue
            # 取规格最接近的 top1（Jaccard）
            best = max(diff, key=lambda r: spec_jaccard(c_spec, _clean(r["__spec"])))
            sim = spec_jaccard(c_spec, _clean(best["__spec"]))
            row = make_match_row(crow, best, CONF_WEAK,
                                 "类目+品牌+型号(规格不一致)", 3)
            row["备注"] = "同系列不同规格"
            row["规格相似度"] = round(sim, 4)
            candidates.append(row)
            # 跨层去重：该竞品行已产候选，关键字层/第4层不再重复处理
            used_cidx.add(i)
    return candidates, used_cidx, used_oid


# ===========================================================================
# 关键字候选层（大模型词库驱动：竞品名按词库词包含匹配，人工池）
# ===========================================================================
KEYWORD_LIB_FILE = os.path.join(OUT_DIR, "自家关键词库.xlsx")


# 关键词停用词（营销/通用/特性词；与 learn.KEYWORD_STOP 保持一致）。
# 词库来源含大模型提取，可能混入"通用/加厚/新款"等纯泛词，命中即误配
# （实测"通用"命中下水管类商品），加载词库时统一过滤。
KEYWORD_STOP = {"正品", "包邮", "家用", "通用", "专用", "新款", "批发", "加厚", "高配",
                "升级", "加强", "精品", "促销", "特价", "优惠", "赠品", "免费", "上门",
                "安装", "维修", "厂家", "直销", "现货", "下单", "客服", "联系", "咨询",
                "备注", "型号", "规格", "颜色", "款式", "左右", "上下", "系列", "套装",
                "闪购", "美团", "闪电仓", "热销", "爆款", "断货", "补货", "到货"}


# ===========================================================================
# 品级语义层（2026-08-14）：分类树逐级下钻，跨品牌品对应
# ===========================================================================
MIN_SALES_FOCUS = 1    # 品对应聚焦阈值：只剔除竞品月销=0（含无销量数据）的品；
                       # 月销≥1 的全部进池，靠价值排序（月售降序）聚焦中高销量


def layer_spu_semantic(dfc, dfo, used_cidx, used_oid, min_sales=MIN_SALES_FOCUS):
    """品级语义层：竞品按「自家商品分类树」逐级下钻（类目→宽品名→细品名词），
    停在最深层；品牌不参与匹配（跨品牌同品可对应，品牌仅作对比列）；
    规格词不参与匹配（品级语义，规格自行查阅）。
    月售<min_sales 的竞品聚焦跳过（写回 used_cidx，不进后续层）。
    返回 (candidates, used_cidx)。"""
    from spu_semantic import SpuSemanticMatcher, _clean as s_clean, _num as s_num

    used_cidx = set(used_cidx)
    dfc_rem = dfc[~dfc.index.isin(used_cidx)].copy()
    if not len(dfc_rem):
        return [], used_cidx

    print("  [品级语义] 构建分类树索引...")
    matcher = SpuSemanticMatcher().build()

    # 自家 SPUID -> 组内 SKU 行列表（候选代表 SKU = 组内月销最高）
    own_by_spu = {}
    for j, orow in dfo.iterrows():
        spu = _clean(orow.get("SPUID"))
        if spu:
            own_by_spu.setdefault(spu, []).append(orow)

    sales_col = "月销量" if "月销量" in dfc.columns else ("月售" if "月售" in dfc.columns else None)
    candidates = []
    n_skip = 0
    for i, crow in dfc_rem.iterrows():
        if sales_col:
            sales = s_num(crow.get(sales_col))
        else:
            sales = min_sales  # 无月售列时不聚焦
        if sales < min_sales:
            used_cidx.add(i)   # 低价值聚焦跳过
            n_skip += 1
            continue
        name = _clean(crow.get("商品名字")) or _clean(crow.get("清洗后商品名"))
        cat = _clean(crow.get("类目三级_归一")) or _clean(crow.get("category3"))
        brand = _clean(crow.get("品牌（抽取）")) or _clean(crow.get("品牌名称"))
        res = matcher.match(name, cat, brand)
        if not res:
            continue
        top3, conf, layer_desc = res
        # 代表 SKU = 月销最高品组内月销最高 SKU
        orows = own_by_spu.get(top3[0], [])
        if not orows:
            continue
        orow = max(orows, key=lambda r: s_num(r.get("月销量") or r.get("月售")))
        row = make_match_row(crow, orow, conf, "品级语义：" + layer_desc, "品级语义(人工)")
        row["语义层级"] = layer_desc
        row["品级置信度"] = conf
        # 候选品组 JSON（top3，仿品级候选列格式）
        cand_list = []
        for gid in top3:
            g_orows = own_by_spu.get(gid, [])
            if not g_orows:
                continue
            rep = max(g_orows, key=lambda r: s_num(r.get("月销量") or r.get("月售")))
            cand_list.append({
                "品组ID": gid,
                "品名": _clean(rep.get("商品名称")) or _clean(rep.get("清洗后商品名")),
                "品牌": _clean(rep.get("品牌（抽取）")),
                "型号": _clean(rep.get("型号")),
                "规格数": len(g_orows),
                "置信度": conf,
                "SKU列表": [_clean(o.get("SKUID")) for o in g_orows][:20],
            })
        row["候选自家品"] = _json_dumps(cand_list)
        row["备注"] = "品级语义：" + layer_desc + "（跨品牌品对应，规格自行查阅）"
        candidates.append(row)
        # 跨层去重：该竞品行已产候选
        used_cidx.add(i)
    print(f"  [品级语义] 候选 {len(candidates)} 条 | 月售<{min_sales} 聚焦跳过 {n_skip} 条")
    return candidates, used_cidx


def _load_keyword_lib():
    """读取自家关键词库（状态=启用，且非停用词）。返回 {词: [(品组ID, 品名, 类目)]}。"""
    if not os.path.exists(KEYWORD_LIB_FILE):
        return {}
    try:
        df = pd.read_excel(KEYWORD_LIB_FILE, dtype=str).fillna("")
        if "关键词" not in df.columns or "品组ID" not in df.columns:
            return {}
        lib = {}
        for _, r in df.iterrows():
            if _clean(r.get("状态")) not in ("", "启用"):
                continue
            w = _clean(r.get("关键词"))
            if not w or len(w) < 2 or w in KEYWORD_STOP:
                continue
            lib.setdefault(w, []).append((_clean(r.get("品组ID")), _clean(r.get("品名")),
                                          _clean(r.get("类目"))))
        return lib
    except Exception:
        return {}


def _kw_conf(score, n_hit):
    """关键字候选置信度（2026-08-13 校准）：按加权分分段。
    加权分<0.6（次强词单词命中/弱词组合）-> 0.60 低档；
    0.6~1.0（弱组合）-> 0.63；>=1.0（强词/强组合）-> 0.63~0.69 按命中词数梯度。"""
    if score < 0.6:
        return 0.60
    if score < 1.0:
        return 0.63
    return round(min(0.69, 0.60 + 0.03 * min(n_hit, 3)), 2)


def layer_keyword(dfc, dfo, used_cidx, used_oid):
    """关键字候选：竞品名按词库词"包含匹配"，命中词关联自家品组 -> 人工池候选。
    同框：词库词带类目时要求与竞品类目一致。排序按"词指向性"（一词关联品组越少越点名）。

    收敛规则（人工池压缩，2026-08-13 改造）：
      - 每个竞品行只产 1 条候选：代表候选 = 加权分最高的品组；
        top3 品组序列化进「候选自家品」JSON（仿品级候选列），人工页仍可看到全部候选；
      - 命中的竞品行写回 used_cidx，后续层（第4层）不再重复处理（跨层去重）。"""
    lib = _load_keyword_lib()
    if not lib:
        print("  [关键字候选] 词库为空：请先运行 build_keyword_lib.py 生成自家关键词库")
        return [], used_cidx, used_oid
    # 词→指向品组数（判别力：指向越少越点名）
    word_pt = {w: len(set(g for g, _, _ in v)) for w, v in lib.items()}
    # 词按长度降序预排序（提出循环，避免每竞品行重排 O(n log n)）
    words_sorted = sorted(lib.keys(), key=len, reverse=True)

    dfc = dfc.copy()
    dfo = dfo.copy()
    own_avail = dfo[~dfo["__oid"].isin(used_oid)]
    dfc_rem = dfc[~dfc.index.isin(used_cidx)]
    # 自家按品组索引（SPUID -> SKU 行列表）
    own_by_spu = {}
    if "SPUID" in own_avail.columns:
        for j, orow in own_avail.iterrows():
            own_by_spu.setdefault(_clean(orow.get("SPUID")), []).append(orow)

    def _sn(v):
        try:
            return float(str(v).replace(",", "").strip())
        except (ValueError, TypeError):
            return 0.0

    candidates = []
    for i, crow in dfc_rem.iterrows():
        cname = _clean(crow["商品名字"]) or _clean(crow.get("清洗后商品名", ""))
        ccat = _clean(crow.get("类目三级_归一", ""))
        if not cname:
            continue
        # 词库包含匹配（词按长度降序，长词优先）
        hit_words = []
        for w in words_sorted:
            if w in cname:
                hit_words.append(w)
        if not hit_words:
            continue
        # 候选品组：类目一致优先；词命中加权（指向品组数越少权重越高）
        spu_score = {}
        spu_hits = {}
        for w in hit_words:
            wt = 1.0 / max(word_pt.get(w, 5), 1)
            for gid, gname, gcat in lib[w]:
                if gid not in own_by_spu:
                    continue
                if gcat and ccat and gcat != ccat:
                    continue  # 词库词带类目且不一致 -> 同框排除
                spu_score[gid] = spu_score.get(gid, 0.0) + wt
                spu_hits.setdefault(gid, []).append(w)
        # 加权门槛：加权和 >= 0.5 才给候选（强词 1/指向1=1.0 达标；
        # 泛词如"插排"指向多品组权重<0.5 被过滤，防止单泛词误配）
        spu_score = {g: s for g, s in spu_score.items() if s >= 0.5}
        if not spu_score:
            continue
        top_spus = sorted(spu_score.items(), key=lambda kv: -kv[1])[:3]
        # 候选品组明细（JSON 供人工确认页展示；结构仿品级候选列）
        cand_list = []
        best_gid, best_conf = None, 0.0
        for gid, score in top_spus:
            orows = own_by_spu[gid]
            # 候选 SKU：品组内月销最高的一个（代表该品）
            best = max(orows, key=lambda r: _sn(r.get("月销量") or r.get("月售")))
            hit = sorted(set(spu_hits[gid]), key=len, reverse=True)[:4]
            conf = _kw_conf(score, len(hit))
            cand_list.append({
                "品组ID": gid,
                "品名": _clean(best.get("商品名称")) or _clean(best.get("清洗后商品名")),
                "品牌": _clean(best.get("品牌（抽取）")),
                "型号": _clean(best.get("型号")),
                "规格数": "",
                "置信度": conf,
                "SKU列表": [_clean(o.get("SKUID")) for o in orows][:20],
            })
            if conf > best_conf:
                best_conf, best_gid = conf, gid
        if best_gid is None:
            continue
        # 每竞品行只产 1 条：代表候选 = 置信度最高的品组
        best = max(own_by_spu[best_gid],
                   key=lambda r: _sn(r.get("月销量") or r.get("月售")))
        hit = sorted(set(spu_hits[best_gid]), key=len, reverse=True)[:4]
        row = make_match_row(crow, best, best_conf,
                             "关键字候选：命中[" + ",".join(hit) + "]",
                             "关键字候选(人工)")
        row["备注"] = "关键字候选(词库)"
        row["命中词数"] = len(hit)
        # 加权分（内部列：供阈值校准/质量分析用；= 代表品组的加权命中分）
        row["加权分"] = round(spu_score[best_gid], 4)
        row["候选自家品"] = _json_dumps(cand_list)
        row["品级置信度"] = best_conf
        candidates.append(row)
        # 跨层去重：该竞品行已产候选，第4层不再重复处理
        used_cidx.add(i)
    return candidates, used_cidx, used_oid


# ===========================================================================
# 向量候选层（P2.4 本地小模型兜底：异名同物语义召回）
# ===========================================================================
_BASE_DIR_ALIGN = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
_VEC_CACHE_DIR = os.path.join(_BASE_DIR_ALIGN, "output", "_向量缓存")
MODEL_DIR_ALIGN = os.path.join(_BASE_DIR_ALIGN, "scripts", "models", "bge-small-zh-v1.5")
_VEC_CONF_MIN = 0.62          # 基础门槛（类目锚点可用时；同物实测 0.81+，灰区 0.55-0.62 误配多，首版从严）
_VEC_CONF_STRONG = 0.68       # 无类目锚点时的强门槛
_CAT_TH_CACHE = {"key": None, "map": {}}


def _cat_thresholds():
    """类目分位门槛（M5·2.4d）：读 output/_train_v4/threshold_calibration.json 校准表。

    启用条件（宁缺毋滥，四道闸）：
      1) 该类目正对 ≥30（样本足够）
      2) 该类目难负 ≥100（分位估计可信）
      3) sug_th = max(0.62, 难负p90) > 0.62（确有收紧必要）
      4) 正对中位 ≥ sug_th + 0.05（类内可分性足够——否则收紧会误杀真对，
         如 家用五金工具 pos_med 0.652 < sug_th 0.675，收紧将流失 58% 真对，不启用）
    当前数据（2026-09-10 校准）下 0 类达标 → 全局维持 0.62；未来重新校准后自动生效。"""
    path = os.path.join(_BASE_DIR_ALIGN, "output", "_train_v4", "threshold_calibration.json")
    try:
        key = (path, int(os.path.getmtime(path)) if os.path.exists(path) else 0)
    except OSError:
        key = (path, 0)
    if _CAT_TH_CACHE.get("key") == key:
        return _CAT_TH_CACHE["map"]
    mp = {}
    if os.path.exists(path):
        try:
            with open(path, encoding="utf-8") as _tf:
                data = json.load(_tf) or {}
            for r in data.get("cat_rows", []):
                n_pos = int(r.get("n_pos") or 0)
                n_neg = int(r.get("n_neg") or 0)
                sug = r.get("sug_th") or 0
                pmed = r.get("pos_med") or 0
                if n_pos >= 30 and n_neg >= 100 and sug > 0.62 and pmed >= sug + 0.05:
                    mp[str(r["cat"])] = float(sug)
        except Exception:
            mp = {}
    _CAT_TH_CACHE["key"] = key
    _CAT_TH_CACHE["map"] = mp
    return mp


def _vec_clean(name, brand=""):
    """品名预处理：剔品牌与营销词，**保留规格数字**（规格是同品/非同品判别核心）。
    实现统一在 scripts/text_norm.py，与训练数据生成共用同一份代码。"""
    return TN.clean(name, brand)


def _own_group_vectors(dfo):
    """自家品组编码（缓存：output/_向量缓存/own_spu_*.npz）。
    缓存键 = 自家表指纹 + **模型身份** + **清洗口径版本** + 品名文本指纹。
    品组=SPUID 分组；品名=组内有效月销最高 SKU；品牌按该 SKU 剔除。
    返回 [(SPUID, 品组名, 类目, 代表行)] 列表 + 向量矩阵；模型缺失返回 (None, None)。"""
    import vec_encoder as VE
    if not VE.is_ready():
        return None, None
    try:
        fp = f"{os.path.getmtime(OWN_NORM)}-{os.path.getsize(OWN_NORM)}"
    except OSError:
        fp = "unknown"
    import hashlib as _hl
    # 模型身份与清洗口径必须进键：否则换模型/改清洗后仍复用旧向量（静默失效）
    tag = "own_spu_" + _hl.md5(
        f"{fp}|{VE.model_id()}|{TN.VERSION}".encode()).hexdigest()[:10]
    os.makedirs(_VEC_CACHE_DIR, exist_ok=True)
    npz_path = os.path.join(_VEC_CACHE_DIR, tag + ".npz")
    meta_path = os.path.join(_VEC_CACHE_DIR, tag + ".json")
    groups = []
    if "SPUID" in dfo.columns:
        import sales_text as _ST
        tmp = dfo.copy()
        tmp["__s"] = tmp["月销量"].apply(_to_float).fillna(0.0)
        tmp = tmp.drop_duplicates(subset=["SKUID"], keep="first")
        for gid, g in tmp.groupby(tmp["SPUID"].astype(str).str.strip()):
            g = g.sort_values("__s", ascending=False)
            best = g.iloc[0]
            name = _clean(best.get("商品名称")) or _clean(best.get("清洗后商品名"))
            if not name:
                continue
            groups.append((_clean(best.get("SPUID")), name,
                           _clean(best.get("类目三级_归一")), best))
    if not groups:
        return None, None
    texts = [_vec_clean(g[1], _clean(g[3].get("品牌（抽取）"))) for g in groups]
    text_fp = _hl.md5("\u0001".join(texts).encode()).hexdigest()[:10]
    if os.path.exists(npz_path) and os.path.exists(meta_path):
        try:
            import json as _json
            meta = _json.load(open(meta_path, encoding="utf-8"))
            if meta.get("n") == len(groups) and meta.get("text_fp") == text_fp:
                import numpy as _np
                return groups, _np.load(npz_path)["v"]
        except Exception:
            pass
    print(f"  [向量缓存] 自家品组首次编码 {len(groups)} 组...")
    import numpy as _np
    vecs = VE.encode(texts)
    _np.savez_compressed(npz_path, v=vecs)
    import json as _json
    with open(meta_path, "w", encoding="utf-8") as fp:
        _json.dump({"n": len(groups), "text_fp": text_fp, "model": VE.model_id(),
                    "clean_version": TN.VERSION, "at": time.strftime("%Y-%m-%d %H:%M")},
                   fp, ensure_ascii=False)
    return groups, vecs


def layer_vector(dfc, dfo, used_cidx, used_oid):
    """第 3.6 层·向量候选（人工池）：异名同物语义召回，绝不自动确认。

    规则（与关键字层同款收敛）：
      - 自家按品组（SPUID 分组）编码一次（缓存 output/_向量缓存/）；
      - 竞品（有效销量≥1）批量编码 → 与品组向量内积 → 类目硬约束 + 门槛过滤；
        双门槛：类目锚点可用（竞品类目与品组类目一致）= 0.62；无锚点 = 0.68 强门槛；
      - 每竞品行只产 1 条代表候选（top3 进「候选自家品」JSON，仿关键字层）；
      - 命中写回 used_cidx（第4层不再重复处理，跨层去重）。

    启用门控（2026-09-08 校准结论）：原版 BGE 黄金集召回够用（recall@5=71.6%）
    但**难负分辨不足**（难负 cos 中位 0.790 ≈ 正对 0.793，99.3% 超门槛）——
    未微调就开会把灰区噪音灌进人工池（违背减负原则）。故：
      默认关闭；仅当 models/bge-small-zh-v1.5/FINETUNED 标记存在或 env VECTOR_LAYER=1 才启用。
      微调 v1 落地（训练样本 320 正 + 134 难负已就绪）后放标记即开。"""
    used_cidx = set(used_cidx)
    used_oid = set(used_oid)
    import vec_encoder as VE
    marker = os.path.join(_BASE_DIR_ALIGN, "scripts", "models", "FINETUNED")
    if not (os.path.exists(marker) or os.environ.get("VECTOR_LAYER") == "1"):
        print("  [向量候选] 未启用（原版 BGE 灰区噪音大，待微调后放 FINETUNED 标记开启；"
              "测试可用 env VECTOR_LAYER=1）")
        return [], used_cidx, used_oid
    groups, gvecs = _own_group_vectors(dfo)
    if groups is None:
        print("  [向量候选] 模型未就绪（scripts/models/bge-small-zh-v1.5/），本层跳过")
        return [], used_cidx, used_oid
    import numpy as _np
    import sales_text as _ST

    # 自家品组 → 组内行索引（供代表 SKU 选取）
    by_gid = {}
    for k, g in enumerate(groups):
        by_gid[g[0]] = k

    dfc_rem = dfc[~dfc.index.isin(used_cidx)].copy()
    sales_col = "月销量" if "月销量" in dfc_rem.columns else "月售"
    keep_idx = [i for i, r in dfc_rem.iterrows() if _ST.eff(r.get(sales_col)) >= MIN_SALES_FOCUS]
    if not keep_idx:
        return [], used_cidx, used_oid
    dfc_rem = dfc_rem.loc[keep_idx]
    texts = [_vec_clean(_clean(r.get("商品名字")) or _clean(r.get("清洗后商品名", "")),
                        _clean(r.get("品牌（抽取）")))
             for _, r in dfc_rem.iterrows()]
    # 竞品向量缓存（M5·2.4b）：键=输出目录+清洗后文本指纹+模型身份+清洗版本。
    # 竞品名不变的重跑直接读缓存（对齐从 100-220s 降回 ~30-60s/店）；
    # 任一要素变化自动作废重算，与自家品组缓存同一套安全语义。
    import hashlib as _hl2
    c_fp = _hl2.md5("\u0001".join(texts).encode()).hexdigest()[:10]
    c_tag = "comp_" + _hl2.md5(
        f"{OUT_DIR}|{len(texts)}|{c_fp}|{VE.model_id()}|{TN.VERSION}".encode()
    ).hexdigest()[:10]
    c_npz = os.path.join(_VEC_CACHE_DIR, c_tag + ".npz")
    c_meta = os.path.join(_VEC_CACHE_DIR, c_tag + ".json")
    cvecs = None
    if os.path.exists(c_npz) and os.path.exists(c_meta):
        try:
            import json as _cj
            cm = _cj.load(open(c_meta, encoding="utf-8"))
            if cm.get("n") == len(texts) and cm.get("text_fp") == c_fp:
                import numpy as _npc
                cvecs = _np.load(c_npz)["v"]
        except Exception:
            cvecs = None
    if cvecs is None:
        cvecs = VE.encode(texts)
        try:
            _np.savez_compressed(c_npz, v=cvecs)
            import json as _cj2
            with open(c_meta, "w", encoding="utf-8") as _fm:
                _cj2.dump({"n": len(texts), "text_fp": c_fp, "model": VE.model_id(),
                           "clean_version": TN.VERSION,
                           "at": time.strftime("%Y-%m-%d %H:%M")}, _fm, ensure_ascii=False)
        except Exception:
            pass

    def _sn(v):
        try:
            return float(str(v).replace(",", "").strip())
        except (ValueError, TypeError):
            return 0.0

    candidates = []
    for k, (i, crow) in enumerate(dfc_rem.iterrows()):
        sims = gvecs @ cvecs[k]
        ccat = _clean(crow.get("类目三级_归一", ""))
        picked, best_k, best_conf = [], None, 0.0
        for kk in sims.argsort()[::-1][:8]:
            cos = float(sims[kk])
            gid, gname, gcat, _brow = groups[kk]
            if ccat and gcat and gcat != ccat:
                continue                      # 类目硬约束：不一致不进候选
            if ccat and gcat:
                th = _cat_thresholds().get(gcat) or _VEC_CONF_MIN
            else:
                th = _VEC_CONF_STRONG         # 无类目锚点：强门槛
            if cos < th:
                break                         # 相似度降序，低于门槛即止
            conf = round(min(0.60, 0.50 + (cos - 0.62) * 0.5), 2)
            picked.append({"品组ID": gid, "品名": gname, "类目": gcat,
                           "相似度": round(cos, 3), "置信度": conf})
            if best_k is None:
                best_k, best_conf = kk, conf
            if len(picked) >= 3:
                break
        if best_k is None:
            continue
        gid, _gname, _gcat, best_row = groups[best_k]
        row = make_match_row(crow, best_row, best_conf,
                             f"向量候选：语义相似 {float(sims[best_k]):.2f}",
                             "向量候选(人工)")
        row["备注"] = "向量候选(本地小模型)"
        row["相似度"] = round(float(sims[best_k]), 3)
        row["候选自家品"] = _json_dumps(picked)
        row["品级置信度"] = best_conf
        candidates.append(row)
        used_cidx.add(i)                       # 跨层去重：第4层不再重复处理
    print(f"  [向量候选] 编码竞品 {len(dfc_rem)} 行，候选 {len(candidates)} 条")
    return candidates, used_cidx, used_oid


# ===========================================================================
# 第4层 模糊兜底（TF-IDF + 图片pHash）
# ===========================================================================
def layer4_fuzzy(dfc, dfo, used_cidx, used_oid):
    """对前3层未命中竞品做模糊匹配。
    有类目竞品按类目分组 TF-IDF；无类目但有品牌的竞品按品牌分组兜底
    （类目缺失/未映射的竞品也能进人工池，不再静默跳过）。
    融合图片 pHash。
    返回 (candidates, used_cidx)：命中的竞品行写回 used_cidx（跨层去重 + 统计口径）。"""
    used_cidx = set(used_cidx)
    dfc_rem = dfc[~dfc.index.isin(used_cidx)].copy()
    own_avail = dfo[~dfo["__oid"].isin(used_oid)].copy()

    # 自家图片 url 列：bind_own_images.py 写入「商品图片url」，旧链路可能有「图片url」。
    # 统一优先读「商品图片url」（bind 成果，8716 行有图），兼容旧列名。
    own_url_col = ("商品图片url" if "商品图片url" in dfo.columns
                   else ("图片url" if "图片url" in dfo.columns else None))
    own_has_url = own_url_col is not None
    img_cache = {}

    candidates = []
    n_done = 0
    # 层内自家去重：一个自家 SKU 只给一个竞品行做代表候选（避免池内自家侧重复）
    layer_used_oid = set()

    def _match_group(gc, go, tag):
        """组内 TF-IDF 模糊匹配，命中追加到 candidates。返回本组命中数。"""
        nonlocal n_done
        if gc.empty or go.empty:
            return 0
        # 文本：竞品用 商品名字，自家用 商品名称
        c_text = (gc["商品名字"].fillna("").astype(str)
                  + " " + gc["品牌（抽取）"].fillna("").astype(str)
                  + " " + gc["型号"].fillna("").astype(str)).apply(norm_text)
        o_text = (go["商品名称"].fillna("").astype(str)
                  + " " + go["品牌（抽取）"].fillna("").astype(str)
                  + " " + go["型号"].fillna("").astype(str)).apply(norm_text)
        corpus = list(c_text) + list(o_text)
        # 过滤空文档
        if not any(corpus):
            return 0
        try:
            vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
            mat = vec.fit_transform(corpus)
        except ValueError:
            return 0
        c_mat = mat[: len(c_text)]
        o_mat = mat[len(c_text):]
        # 稀疏 top-k 检索（2026-08-13）：逐行稀疏点积取 top5，替代稠密 cosine_similarity。
        # 稠密 (nc×no) 在大类目组（竞品1万×自家6000）约 480MB 有 OOM 风险；
        # 逐行 c_mat[a] @ o_mat.T 只物化单行（no 个 float），内存与耗时均大幅下降。
        o_mat_t = o_mat.T.tocsr()          # V x no
        gc_idx = gc.index.tolist()
        go_idx = go.index.tolist()
        hit = 0
        for a in range(c_mat.shape[0]):
            crow = dfc_rem.loc[gc_idx[a]]
            row_sims = np.asarray((c_mat[a] @ o_mat_t).todense()).ravel()
            # 自家侧去重：优先取未占用的最相似自家（top5 内第一个未占用）；
            # 全部占用则本竞品无候选（静默跳过，不进池）
            best = None
            for b in np.argsort(row_sims)[::-1][:5]:
                if own_avail.loc[go_idx[b], "__oid"] not in layer_used_oid:
                    best = int(b)
                    break
            if best is None:
                continue
            orow = own_avail.loc[go_idx[best]]
            text_sim = float(row_sims[best])
            # 图片url同源检查（同域名+同文件名=强匹配信号）
            comp_img_url = _clean(crow.get("图片url", ""))
            own_img_url = _clean(orow.get(own_url_col, "")) if own_url_col else ""
            url_match, url_match_status = check_image_url_match(comp_img_url, own_img_url)
            # 图片pHash相似度（双端均有url时计算，受PHASH_THRESHOLD约束）
            img_sim = None
            phash_distance = None
            if own_has_url and comp_img_url and own_img_url:
                pc = download_image_phash(comp_img_url, img_cache)
                po = download_image_phash(own_img_url, img_cache)
                if pc and po:
                    similar, dist = compare_images_phash(pc, po, PHASH_THRESHOLD)
                    phash_distance = dist
                    if similar:
                        img_sim = 1.0 - dist / 64.0
                    # 图片不相似：img_sim 保持 None，退化为纯文本（不拉低融合分）
            # 融合：图片可用则 text*0.6+img*0.4，否则退化为 text。
            # 先算融合、最后统一过阈值：图片证据可兜底文本弱匹配（0.49文本+1.0图可过阈值）
            if img_sim is not None:
                fusion = text_sim * 0.6 + img_sim * 0.4
                basis = "模糊(文本+图片)"
            else:
                fusion = text_sim
                basis = "模糊(仅文本,图片缺失/不相似)"
            # 品牌硬惩罚：双方品牌明确且不一致 -> 融合分×0.5（跨品牌误配防线）
            if _brand_conflict(crow.get("品牌（抽取）"), orow.get("品牌（抽取）")):
                fusion *= BRAND_PENALTY
                basis += "+品牌不一致惩罚"
            # 图片url同源加成（辅助证据，置信度+0.1）
            if url_match:
                fusion += URL_MATCH_BONUS
                basis += "+图url同源"
            if fusion < FUZZY_THRESHOLD:
                continue
            conf = min(fusion, CONF_FUZZY_CAP)
            row = make_match_row(crow, orow, round(conf, 4), basis, 4)
            row["文本相似度"] = round(text_sim, 4)
            row["图片相似度"] = "" if img_sim is None else round(img_sim, 4)
            row["融合分数"] = round(fusion, 4)
            row["图url同源"] = url_match_status
            row["pHash距离"] = "" if phash_distance is None else phash_distance
            row["备注"] = "强制人工确认" + (f"（{tag}）" if tag else "")
            candidates.append(row)
            hit += 1
            n_done += 1
            # 跨层去重：该竞品行已产候选，写回供统计与后续排除
            used_cidx.add(gc_idx[a])
            # 层内去重：该自家 SKU 已被占用
            layer_used_oid.add(orow["__oid"])
            if n_done % 500 == 0:
                print(f"  第4层已生成候选 {n_done} 条")
        return hit

    # 1) 有类目竞品：按类目分组
    cats = [c for c in dfc_rem["类目三级_归一"].dropna().unique() if c]
    for cat in cats:
        gc = dfc_rem[dfc_rem["类目三级_归一"] == cat]
        go = own_avail[own_avail["类目三级_归一"] == cat]
        _match_group(gc, go, cat)

    # 2) 无类目竞品（类目缺失/未映射）：按品牌分组兜底，无品牌的跳过
    dfc_rem["__cat"] = dfc_rem["类目三级_归一"].fillna("").astype(str).str.strip()
    no_cat = dfc_rem[dfc_rem["__cat"] == ""]
    if len(no_cat):
        own_no_brand = own_avail[own_avail["品牌（抽取）"].fillna("").astype(str).str.strip() != ""]
        for brand, gc in no_cat.groupby(no_cat["品牌（抽取）"].fillna("").astype(str).str.strip()):
            if brand == "":
                continue
            go = own_no_brand[own_no_brand["品牌（抽取）"].fillna("").astype(str).str.strip() == brand]
            _match_group(gc, go, f"未分类/{brand}")

    return candidates, used_cidx


# ===========================================================================
# SPU 聚合对齐（先品对应再细化规格）
# 在4层SKU级匹配之前执行：把同三级类目+品牌+型号的SKU聚合成一个SPU组，
# 先把"品"对应上，再在组内细化规格匹配。
# ===========================================================================
CONF_SPU = 0.90          # SPU级匹配置信度（类目+品牌+型号全命中）
CONF_SKU_IN_SPU = 0.95   # SPU组内SKU规格一致匹配置信度


def build_spu_groups(df, side):
    """构建SPU组：同三级类目+品牌+型号的SKU聚合为一个SPU。
    品牌或型号为空的不参与SPU聚合（无法确定同品）。
    返回DataFrame，每行一个SPU组，列：
      SPU组ID / 三级类目 / 品牌 / 型号 / SKU数量 / SKU列表 / 商品名 / 来源
    """
    df = df.copy()
    if side == "comp":
        sku_col, fallback_col, name_col, spu_prefix = "Sku_id", "商品ID", "商品名字", "SPU_C"
    else:
        sku_col, fallback_col, name_col, spu_prefix = "SKUID", "SPUID", "商品名称", "SPU_O"

    # SKU ID 列：缺失时退化（与 __cid/__oid 逻辑保持一致，便于回填）
    df["__sid"] = df[sku_col].where(
        df[sku_col].astype(str).str.strip() != "", df[fallback_col]
    ).astype(str)

    # 清洗聚合键
    df["__cat"] = df["类目三级_归一"].apply(_clean)
    df["__brand"] = df["品牌（抽取）"].apply(_clean)
    df["__model"] = df["型号"].apply(_clean)

    # 过滤空品牌/空型号/空类目（无法确定同品）
    valid = df[
        (df["__brand"] != "") & (df["__model"] != "") & (df["__cat"] != "")
    ].copy()

    groups = []
    for (cat, brand, model), g in valid.groupby(
        ["__cat", "__brand", "__model"], dropna=False
    ):
        sku_list = [s for s in g["__sid"].astype(str).tolist() if s and s.lower() != "nan"]
        if not sku_list:
            continue
        # 取第一个非空商品名作为代表
        rep_name = ""
        for v in g[name_col].tolist():
            n = _clean(v)
            if n:
                rep_name = n
                break
        groups.append({
            "三级类目": cat,
            "品牌": brand,
            "型号": model,
            "SKU数量": len(sku_list),
            "SKU列表": ",".join(sku_list),
            "商品名": rep_name,
            "来源": side,
        })

    spu_df = pd.DataFrame(groups)
    if not len(spu_df):
        # 兜底：品牌/型号/类目全空（如整店类目缺失）时返回带列的空表，避免下游排序报错
        return pd.DataFrame(columns=["三级类目", "品牌", "型号", "SKU数量",
                                     "SKU列表", "商品名", "来源", "SPU组ID"])
    spu_df = spu_df.sort_values(["三级类目", "品牌", "型号"]).reset_index(drop=True)
    spu_df.insert(0, "SPU组ID", [f"{spu_prefix}{i+1:04d}" for i in range(len(spu_df))])
    return spu_df


def align_spu_groups(comp_spu, own_spu):
    """SPU组对齐：三级类目+品牌+型号完全一致。
    返回匹配的SPU对列表，每对为一字典，包含：
      竞品SPU组ID / 自家SPU组ID / 三级类目 / 品牌 / 型号 /
      竞品SKU数 / 自家SKU数 / 匹配状态="SPU对齐" / 置信度=0.90 /
      竞品SKU列表 / 自家SKU列表
    """
    matches = []
    if not len(comp_spu) or not len(own_spu):
        return matches

    # 自家按(类目, 品牌, 型号)索引
    own_idx = {}
    for _, orow in own_spu.iterrows():
        key = (orow["三级类目"], orow["品牌"], orow["型号"])
        own_idx.setdefault(key, []).append(orow)

    for _, crow in comp_spu.iterrows():
        key = (crow["三级类目"], crow["品牌"], crow["型号"])
        for orow in own_idx.get(key, []):
            matches.append({
                "竞品SPU组ID": crow["SPU组ID"],
                "自家SPU组ID": orow["SPU组ID"],
                "三级类目": crow["三级类目"],
                "品牌": crow["品牌"],
                "型号": crow["型号"],
                "竞品SKU数": int(crow["SKU数量"]),
                "自家SKU数": int(orow["SKU数量"]),
                "匹配状态": "SPU对齐",
                "置信度": CONF_SPU,
                "竞品SKU列表": crow["SKU列表"],
                "自家SKU列表": orow["SKU列表"],
            })
    return matches


def refine_sku_within_spu(comp_df, own_df, spu_pair, used_cidx=None):
    """对已对齐的SPU对，在组内做规格匹配。
    同SPU下，规格完全一致=SKU级匹配（"SKU对齐"）；
    否则该SPU对进入"同品缺规格"建议。

    返回 (sku_aligned_rows, missing_spec_suggestion):
      sku_aligned_rows: list[dict]，SKU级匹配行（规格一致），标记"SKU对齐"，
                        每行附 __cidx/__oidx（竞品/自家 DataFrame 行号，供外层标记已用）
      missing_spec_suggestion: dict 或 None，缺规格建议（竞品有自家无的规格）

    注：刻意不传入全局 used_oid——同品组内一个自家SKU可对应多个竞品SKU
    （一对多，如条码相同的多规格竞品），组内去重由局部 used_oid 完成；
    对第3/4层的排除由外层把本层占用的 __oid 写回全局 used_oid 实现。
    """
    used_cidx = used_cidx or set()
    sku_rows = []
    comp_skus = [s for s in str(spu_pair["竞品SKU列表"]).split(",") if s]
    own_skus = [s for s in str(spu_pair["自家SKU列表"]).split(",") if s]

    # comp_df/own_df 已带 __cid/__oid，与 build_spu_groups 的 __sid 逻辑一致
    comp_sub = comp_df[comp_df["__cid"].astype(str).isin(comp_skus)].copy()
    own_sub = own_df[own_df["__oid"].astype(str).isin(own_skus)].copy()

    if comp_sub.empty or own_sub.empty:
        return sku_rows, None

    comp_sub["__spec"] = comp_sub.apply(build_spec_key, axis=1)
    own_sub["__spec"] = own_sub.apply(build_spec_key, axis=1)

    # 自家按规格键索引
    own_spec_idx = {}
    for _, orow in own_sub.iterrows():
        spec = _clean(orow["__spec"])
        own_spec_idx.setdefault(spec, []).append(orow)

    used_oid = set()
    for _, crow in comp_sub.iterrows():
        if crow.name in used_cidx:
            continue  # 已被上层匹配占用
        c_spec = _clean(crow["__spec"])
        if c_spec == "":
            continue  # 无规格不参与SKU级匹配
        for orow in own_spec_idx.get(c_spec, []):
            oid = orow["__oid"]
            if oid in used_oid:
                continue
            row = make_match_row(
                crow, orow, CONF_SKU_IN_SPU, "SPU内SKU规格一致", "SPU"
            )
            row["竞品SPU组ID"] = spu_pair["竞品SPU组ID"]
            row["自家SPU组ID"] = spu_pair["自家SPU组ID"]
            row["匹配状态"] = "SKU对齐"
            row["__cidx"] = crow.name
            row["__oidx"] = orow.name
            row["__oid"] = orow["__oid"]
            sku_rows.append(row)
            used_oid.add(oid)
            break

    # 缺规格建议：竞品有规格但自家无对应的
    own_specs = {s for s in own_spec_idx.keys() if s != ""}
    comp_specs = {_clean(s) for s in comp_sub["__spec"].tolist()} - {""}
    missing = comp_specs - own_specs
    suggestion = None
    if missing:
        suggestion = {
            "SPU组ID": spu_pair["竞品SPU组ID"],
            "自家SPU组ID": spu_pair["自家SPU组ID"],
            "三级类目": spu_pair["三级类目"],
            "品牌": spu_pair["品牌"],
            "型号": spu_pair["型号"],
            "竞品SKU列表": spu_pair["竞品SKU列表"],
            "自家SKU列表": spu_pair["自家SKU列表"],
            "建议内容": "补入规格：" + ", ".join(sorted(missing)),
        }
    return sku_rows, suggestion


# ===========================================================================
# 品级细化层（第2层失败后、第3层之前）：先按品收敛，再在品内对规格
# ===========================================================================
CONF_SPU_SPEC = 0.95      # 品内规格一致（自动确认）
CONF_SPU_FALLBACK_CAP = 0.69  # 缺失兜底候选（名称相似度，强制人工）
MIN_FALLBACK_SIM = 0.25   # 缺失兜底最低文本相似度（用户标准：宁缺毋滥，不做硬凑匹配）
MIN_FALLBACK_CONF = 0.55  # 缺失兜底最低置信度（低于不出候选，留给模糊层）


def build_spu_refine_layer(dfc, dfo, comp_spu, own_spu, spu_pairs,
                           used_cidx, used_oid):
    """品级细化：对第1/2层未命中的竞品，先做品级匹配收敛候选范围。

    1) 严格组（类目+品牌+型号齐全）：已对齐品组内做规格匹配
       - 规格键一致 -> 自动确认行（匹配层=品内规格，置信度 0.95）
       - 同组规格不一致（自家缺该规格）-> 归入缺规格建议（品组级），
         该竞品SKU标记已用（品已对上，规格缺口走建议流程，不再进第3/4层）
    2) 缺失兜底（品牌或型号缺失、未入严格组）：按类目+品牌+名称相似度
       找候选自家品组 Top3 -> 品级候选行（匹配层=品级候选(人工)，置信度<0.7）

    返回 (auto_rows, cand_rows, used_cidx, used_oid)
    """
    auto_rows, cand_rows = [], []
    used_cidx, used_oid = set(used_cidx), set(used_oid)
    missing_spec_rows = []

    # 竞品 SKU -> 严格SPU组（供缺失兜底判断“未入组”用）
    comp_sku_to_spu = {}
    for _, r in comp_spu.iterrows():
        for sku in str(r["SKU列表"]).split(","):
            if sku:
                comp_sku_to_spu[sku] = r["SPU组ID"]

    # 自家品组 -> 组内代表行（dfo 行号 + 代表商品名）
    own_spu_rep = {}
    for _, orow in own_spu.iterrows():
        skus = [s for s in str(orow["SKU列表"]).split(",") if s]
        rep_idx = None
        sub = dfo[dfo["__oid"].astype(str).isin(skus)]
        if len(sub):
            rep_idx = sub.index[0]
        own_spu_rep[orow["SPU组ID"]] = {
            "品组ID": orow["SPU组ID"],
            "品名": orow["商品名"],
            "品牌": orow["品牌"],
            "型号": orow["型号"],
            "规格数": int(orow["SKU数量"]),
            "SKU列表": skus,
            "rep_idx": rep_idx,
        }

    # ---- 1) 严格组内细化 ----
    for pair in spu_pairs:
        sku_rows, suggestion = refine_sku_within_spu(dfc, dfo, pair, used_cidx)
        own_group = own_spu_rep.get(pair["自家SPU组ID"])
        for row in sku_rows:
            row["匹配层"] = "品内规格"
            row["品级置信度"] = CONF_SPU_SPEC
            row["候选自家品"] = ""
            auto_rows.append(row)
            used_cidx.add(row["__cidx"])
            # 写入 __oid（SKUID 字符串）而非 __oidx（整数行号），
            # 与第1/2层口径一致，保证第3/4层 isin 排除生效
            used_oid.add(row["__oid"])
        if suggestion is not None:
            missing_spec_rows.append(suggestion)
        # 同组规格不一致的竞品 SKU：品已对上，规格缺口走缺规格建议，
        # 标记已用（不进第3/4层，避免大范围重复候选）
        comp_skus = [s for s in str(pair["竞品SKU列表"]).split(",") if s]
        comp_sub = dfc[dfc["__cid"].astype(str).isin(comp_skus)]
        for _, crow in comp_sub.iterrows():
            if crow.name not in used_cidx:
                used_cidx.add(crow.name)

    # ---- 2) 缺失兜底（未入严格组的竞品 SKU）----
    comp_rem = dfc[~dfc.index.isin(used_cidx)].copy()
    comp_rem = comp_rem[~comp_rem["__cid"].astype(str).isin(comp_sku_to_spu)].copy()
    if len(comp_rem) and len(own_spu):
        cats = [c for c in comp_rem["类目三级_归一"].dropna().unique() if c]
        for cat in cats:
            gc = comp_rem[comp_rem["类目三级_归一"] == cat]
            # 候选自家品组：类目一致
            go_spu = own_spu[own_spu["三级类目"] == cat]
            if gc.empty or go_spu.empty:
                continue
            c_text = (gc["商品名字"].fillna("").astype(str)
                      + " " + gc["品牌（抽取）"].fillna("").astype(str)
                      + " " + gc["型号"].fillna("").astype(str)).apply(norm_text)
            o_text = (go_spu["商品名"].fillna("").astype(str)
                      + " " + go_spu["品牌"].fillna("").astype(str)
                      + " " + go_spu["型号"].fillna("").astype(str)).apply(norm_text)
            corpus = list(c_text) + list(o_text)
            if not any(corpus):
                continue
            try:
                vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 4))
                mat = vec.fit_transform(corpus)
            except ValueError:
                continue
            c_mat = mat[: len(c_text)]
            o_mat = mat[len(c_text):]
            sims = cosine_similarity(c_mat, o_mat)
            gc_idx = gc.index.tolist()
            go_spu_idx = go_spu.index.tolist()
            c_brands = gc["品牌（抽取）"].fillna("").astype(str).tolist()
            for a in range(sims.shape[0]):
                crow = dfc.loc[gc_idx[a]]
                row_sims = sims[a]
                top = np.argsort(row_sims)[::-1][:3]
                cand_list = []
                best_conf = 0.0
                for b in top:
                    text_sim = float(row_sims[b])
                    if text_sim < MIN_FALLBACK_SIM:
                        continue
                    spu_row = go_spu.loc[go_spu_idx[b]]
                    gid = spu_row["SPU组ID"]
                    rep = own_spu_rep.get(gid)
                    if rep is None or rep["rep_idx"] is None:
                        continue
                    brand_hit = 1.0 if (_clean(c_brands[a]) and
                                        _clean(c_brands[a]) == _clean(spu_row["品牌"])) else 0.0
                    conf = round(min(0.5 + text_sim * 0.35 + brand_hit * 0.1,
                                     CONF_SPU_FALLBACK_CAP), 4)
                    # 品牌硬惩罚：双方品牌明确且不一致 -> 置信度×0.5
                    # （0.69封顶×0.5=0.345 < 0.55门槛，跨品牌候选基本被拦截）
                    if _brand_conflict(c_brands[a], spu_row["品牌"]):
                        conf = round(conf * BRAND_PENALTY, 4)
                    if conf < MIN_FALLBACK_CONF:
                        continue
                    cand_list.append({
                        "品组ID": gid, "品名": rep["品名"],
                        "品牌": rep["品牌"], "型号": rep["型号"],
                        "规格数": rep["规格数"], "置信度": conf,
                        "SKU列表": rep["SKU列表"][:20],
                    })
                    if conf > best_conf:
                        best_conf = conf
                if not cand_list or best_conf < MIN_FALLBACK_CONF:
                    continue
                orow = dfo.loc[own_spu_rep[cand_list[0]["品组ID"]]["rep_idx"]]
                row = make_match_row(crow, orow, best_conf,
                                     "品级候选(名称相似)", "品级候选(人工)")
                row["竞品SPU组ID"] = comp_sku_to_spu.get(_clean(crow["__cid"]), "")
                row["自家SPU组ID"] = cand_list[0]["品组ID"]
                row["品级置信度"] = best_conf
                row["候选自家品"] = _json_dumps(cand_list)
                row["备注"] = "品级候选：品牌/型号缺失，按名称相似兜底，人工确认"
                cand_rows.append(row)
                used_cidx.add(crow.name)

    return auto_rows, cand_rows, missing_spec_rows, used_cidx, used_oid


def _json_dumps(obj):
    """候选自家品 JSON 序列化（确保中文正常输出）。"""
    import json
    return json.dumps(obj, ensure_ascii=False)


# ===========================================================================
# 两层对齐汇总库合并（output/对齐汇总库.xlsx：SPU匹配层 + SKU匹配层）
# ===========================================================================
def auto_confirm_vector(review, shop, prog_path=None):
    """向量高置信自动确认（M5·2.4a 试点：参考级 + 可撤销）。

    四守卫（全过才自动）：
      1) 匹配层=向量候选(人工) 且 相似度≥0.80 —— 校准依据：难负误报率 0.77%
         （14095 难负 / output/_train_v4/threshold_calibration.json，2026-09-10）；
      2) 类目锚点：竞品类目与自家类目一致且非空；
      3) 规格无冲突（text_norm.spec_conflict 为空——残余误报 69% 是规格冲突型）；
      4) 价格校验非异常（价差率超±50% 的不自动）。
    写入：汇总库 SPU 匹配层（确认来源=向量自动，对应级别=参考）+ progress（spu_only）。
    保护：progress 中该 Sku 已有 人工确认 或 无对应标记 → 跳过（人与否决优先）；
          撤销走面板 /api/confirm_undo（删汇总库行+progress 记录）。
    开关：环境变量 VEC_AUTO=0 可禁用。返回写入条数。"""
    if review is None or not len(review) or "匹配层" not in review.columns:
        return 0
    if os.environ.get("VEC_AUTO") == "0":
        print("  [向量自动确认] 已禁用（VEC_AUTO=0）")
        return 0
    from summary_library import (CONFIRM_VEC_AUTO, LEVEL_REF, upsert_auto)

    prog_path = prog_path or os.path.join(BASE, "scripts", "progress.json")
    progress = {}
    if os.path.exists(prog_path):
        try:
            with open(prog_path, encoding="utf-8") as _pf:
                progress = json.load(_pf) or {}
        except Exception:
            progress = {}
    confirmed_map = progress.setdefault("confirmed", {}).setdefault(shop, {})
    no_match_map = (progress.get("no_match") or {}).get(shop) or {}

    sim = pd.to_numeric(review["相似度"], errors="coerce").fillna(0.0)
    cand_idx = review.index[(review["匹配层"].astype(str) == "向量候选(人工)")
                            & (sim >= 0.80)]
    if not len(cand_idx):
        print("  [向量自动确认] 无 sim≥0.80 候选")
        return 0

    spu_df, sku_df = load_library()
    n_write = 0
    n_skip = 0
    for i in cand_idx:
        r = review.loc[i]
        ccat, ocat = _clean(r.get("竞品类目")), _clean(r.get("自家类目"))
        if not (ccat and ocat and ccat == ocat):
            continue
        if _clean(r.get("价格校验状态")) == "异常":
            continue
        gid = _clean(r.get("自家SPU组ID")) or _clean(r.get("自家SPUID"))
        pid = _clean(r.get("竞品品ID"))
        sku = _clean(r.get("竞品Sku_id"))
        if not (gid and pid and sku) or sku in confirmed_map or sku in no_match_map:
            n_skip += 1
            continue
        if TN.spec_conflict(r.get("竞品商品名"), r.get("自家商品名")):
            n_skip += 1
            continue
        spu_row = {
            "竞品品ID": pid,
            "竞品商品名": _clean(r.get("竞品商品名")),
            "竞品品牌": _clean(r.get("竞品品牌")),
            "竞品型号": _clean(r.get("竞品型号")),
            "竞品类目": ccat,
            "自家品组ID": gid,
            "自家品名": _clean(r.get("自家商品名")),
            "置信度": _clean(r.get("相似度")),
            "对应级别": LEVEL_REF,
            "来源": "向量",
        }
        spu_df, ok = upsert_auto(spu_df, "竞品品ID", shop, spu_row, CONFIRM_VEC_AUTO)
        if not ok:
            n_skip += 1
            continue
        confirmed_map[sku] = {
            "sku_id": _clean(r.get("自家SKUID")),
            "spu_only": True,
            "source": "向量高置信(自动)",
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        review.at[i, "建议动作"] = "已自动确认（可撤销）"
        review.at[i, "确认结果"] = "向量自动"
        n_write += 1
    if n_write:
        save_library(spu_df, sku_df)
        with open(prog_path, "w", encoding="utf-8") as _wf:
            json.dump(progress, _wf, ensure_ascii=False, indent=2)
    print(f"  [向量自动确认] sim≥0.80 四守卫通过 {n_write} 条"
          f"（跳过 {n_skip} 条）→ SPU层参考级（可撤销）")
    return n_write


def merge_to_summary_library(dfc, summary, spu_pairs, own_spu, shop=None):
    """把本次跑批结果合并进两层汇总库。
    SKU层收 对齐_汇总（机器自动行）；SPU层收 品对品（按竞品品ID展开）。
    人工确认行数据列不被覆盖；店内编号不在本次竞品表 -> 状态=已下架。"""
    if shop is None:
        base = os.path.basename(os.path.normpath(OUT_DIR))
        shop = base[len("试跑_"):] if base.startswith("试跑_") else LIB_DEFAULT_SHOP

    active_sku = {s for s in dfc["Sku_id"].astype(str).str.strip() if s}
    active_pid = {s for s in dfc["商品ID"].astype(str).str.strip() if s}

    # ---- SKU层行（对齐_汇总：第1/2层+品内规格）----
    sku_rows = []
    for _, r in summary.iterrows():
        sku_rows.append({
            "竞品品ID": _clean(r.get("竞品品ID")),
            "竞品Sku_id": _clean(r.get("竞品Sku_id")),
            "竞品商品名": _clean(r.get("竞品商品名")),
            "竞品条形码": _clean(r.get("竞品条形码")),
            "自家SPUID": _clean(r.get("自家SPUID")),
            "自家SKUID": _clean(r.get("自家SKUID")),
            "自家商品名": _clean(r.get("自家商品名")),
            "匹配层": _clean(r.get("匹配层")),
            "匹配依据": _clean(r.get("匹配依据")),
            "置信度": _clean(r.get("置信度")),
        })

    # ---- SPU层行（品对品，按竞品品ID展开）----
    sku2pid = dict(zip(dfc["Sku_id"].astype(str).str.strip(),
                       dfc["商品ID"].astype(str).str.strip()))
    sku2name = dict(zip(dfc["Sku_id"].astype(str).str.strip(),
                        dfc["商品名字"].astype(str)))
    own_name_by_gid = (dict(zip(own_spu["SPU组ID"], own_spu["商品名"]))
                       if len(own_spu) else {})
    spu_rows, seen = [], set()
    for pair in spu_pairs:
        gid_own = pair["自家SPU组ID"]
        for sku in str(pair["竞品SKU列表"]).split(","):
            sku = sku.strip()
            pid = sku2pid.get(sku, "")
            if not pid or (gid_own, pid) in seen:
                continue
            seen.add((gid_own, pid))
            spu_rows.append({
                "竞品品ID": pid,
                "竞品商品名": _clean(sku2name.get(sku, "")),
                "竞品品牌": _clean(pair.get("品牌")),
                "竞品型号": _clean(pair.get("型号")),
                "竞品类目": _clean(pair.get("三级类目")),
                "自家品组ID": gid_own,
                "自家品名": _clean(own_name_by_gid.get(gid_own, "")),
                "置信度": _clean(pair.get("置信度")),
            })

    spu_df, sku_df = load_library()
    spu_df = upsert_rows(spu_df, spu_rows, "竞品品ID", shop, active_pid)
    sku_df = upsert_rows(sku_df, sku_rows, "竞品Sku_id", shop, active_sku)
    save_library(spu_df, sku_df)
    print(f"  [汇总库] {shop}: SPU层键 {len(spu_rows)} 个, SKU层键 {len(sku_rows)} 个 "
          f"-> {LIBRARY_FILE}（库内 SPU层 {len(spu_df)} 行 / SKU层 {len(sku_df)} 行）")


# ===========================================================================
# 主流程
# ===========================================================================
def main():
    import argparse
    parser = argparse.ArgumentParser(description="阶段2 跨店商品对齐引擎")
    parser.add_argument("--comp", help="竞品归一化表路径（默认 output/竞品_归一化后.xlsx）")
    parser.add_argument("--own", help="自家归一化表路径（默认 output/自家_归一化后.xlsx）")
    parser.add_argument("--outdir", help="输出目录（默认 output/，多店试跑时各店独立目录）")
    args = parser.parse_args()

    global COMP_NORM, OWN_NORM, OUT_DIR
    if args.comp:
        COMP_NORM = args.comp
    if args.own:
        OWN_NORM = args.own
    if args.outdir:
        OUT_DIR = args.outdir
        os.makedirs(OUT_DIR, exist_ok=True)

    print("=" * 60)
    print("阶段2 跨店商品对齐引擎")
    print("=" * 60)
    print("加载数据...")
    dfc, dfo = load_data()
    print(f"竞品: {dfc.shape}  自家: {dfo.shape}")
    print(f"竞品条形码非空: {int(dfc['条形码'].apply(norm_barcode).astype(bool).sum())}")
    print(f"自家商品条码非空: {int(dfo['商品条码'].apply(norm_barcode).astype(bool).sum())}")

    # ---- SPU 聚合对齐（先品对应再细化规格，4层SKU级匹配之前）----
    print("\n[SPU聚合] 构建SPU组(三级类目+品牌+型号)...")
    t0 = time.time()
    comp_spu = build_spu_groups(dfc, "comp")
    own_spu = build_spu_groups(dfo, "own")
    print(f"  竞品SPU组: {len(comp_spu)}  自家SPU组: {len(own_spu)}  "
          f"耗时 {time.time()-t0:.1f}s")

    print("\n[SPU对齐] 三级类目+品牌+型号对齐...")
    t0 = time.time()
    spu_pairs = align_spu_groups(comp_spu, own_spu)
    spu_pairs_df = pd.DataFrame(spu_pairs)
    print(f"  SPU对齐: {len(spu_pairs)} 对  耗时 {time.time()-t0:.1f}s")

    print("\n[SPU内SKU细化] 组内规格匹配（并入品级细化，第2层后执行）...")

    # ---- 第1层 ----
    print("\n[第1层] 条码精确匹配...")
    t0 = time.time()
    m1, used_cidx, used_oid = layer1_barcode(dfc, dfo)
    print(f"  匹配 {len(m1)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 第2层 ----
    print("\n[第2层] 强组合键匹配(类目+品牌+型号+规格)...")
    t0 = time.time()
    m2, used_cidx, used_oid = layer2_strong(dfc, dfo, set(used_cidx), set(used_oid))
    print(f"  匹配 {len(m2)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 品级细化（第2层失败后、第3层之前）：先按品收敛，再品内对规格 ----
    print("\n[品级细化] 品内规格匹配（先按品收敛候选范围）...")
    t0 = time.time()
    m_spu_auto, m_spu_cand, missing_spec_rows, used_cidx, used_oid = \
        build_spu_refine_layer(dfc, dfo, comp_spu, own_spu, spu_pairs,
                               used_cidx, used_oid)
    spu_refine_df = pd.DataFrame(m_spu_auto + m_spu_cand)
    missing_spec_df = pd.DataFrame(missing_spec_rows, columns=[
        "SPU组ID", "自家SPU组ID", "三级类目", "品牌", "型号",
        "竞品SKU列表", "自家SKU列表", "建议内容",
    ])
    print(f"  品内规格自动确认: {len(m_spu_auto)} 条 | 品级候选(人工): {len(m_spu_cand)} 条 | "
          f"缺规格建议: {len(missing_spec_rows)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 输出 SPU 级文件 ----
    spu_out_path = os.path.join(OUT_DIR, "对齐_SPU级.xlsx")
    with pd.ExcelWriter(spu_out_path, engine="openpyxl") as writer:
        spu_pairs_df.to_excel(writer, sheet_name="SPU组对齐", index=False)
        spu_refine_df.to_excel(writer, sheet_name="SPU内SKU匹配", index=False)
    spu_refine_path = os.path.join(OUT_DIR, "对齐_SPU细化.xlsx")
    with pd.ExcelWriter(spu_refine_path, engine="openpyxl") as writer:
        spu_pairs_df.to_excel(writer, sheet_name="品对品", index=False)
        spu_refine_df.to_excel(writer, sheet_name="品内规格", index=False)
    missing_spec_path = os.path.join(OUT_DIR, "对齐_SPU级_缺规格建议.xlsx")
    missing_spec_df.to_excel(missing_spec_path, index=False)
    print(f"  -> {spu_out_path}")
    print(f"  -> {spu_refine_path}")
    print(f"  -> {missing_spec_path}")

    # ---- 品级语义层（分类树逐级下钻，跨品牌品对应；月售≥6 聚焦）----
    print("\n[品级语义层] 分类树下钻（跨品牌品对应，月售≥6 聚焦）-> 人工池...")
    t0 = time.time()
    m_sem, used_cidx = layer_spu_semantic(dfc, dfo, set(used_cidx), set(used_oid))
    print(f"  候选 {len(m_sem)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 第3层 ----
    print("\n[第3层] 弱组合键匹配(同系列不同规格) -> 人工池...")
    t0 = time.time()
    m3, used_cidx, used_oid = layer3_weak(dfc, dfo, set(used_cidx), set(used_oid))
    print(f"  候选 {len(m3)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 关键字候选层（词库驱动：竞品名按自家关键词库包含匹配）----
    print("\n[关键字候选层] 大模型词库匹配 -> 人工池...")
    t0 = time.time()
    m35, used_cidx, used_oid = layer_keyword(dfc, dfo, set(used_cidx), set(used_oid))
    print(f"  候选 {len(m35)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 向量候选层（P2.4 本地小模型：异名同物语义召回；模型缺失自动跳过）----
    print("\n[向量候选层] BGE-small-zh 语义兜底 -> 人工池...")
    t0 = time.time()
    m36, used_cidx, used_oid = layer_vector(dfc, dfo, set(used_cidx), set(used_oid))
    print(f"  候选 {len(m36)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 第4层（M5·2.4e 简化：无类目竞品占比 <10% 时跳过）----
    # 依据：向量层（3.6）已接管语义兜底，现役 5 店无类目占比均为 0%、层4产出恒为 0；
    # 层4（TF-IDF+图片pHash）成本高，仅保留给"无类目店"（类目缺失的店仍需模糊兜底）。
    no_cat_ratio = 1.0
    if "类目三级_归一" in dfc.columns:
        no_cat_ratio = float((dfc["类目三级_归一"].astype(str).str.strip() == "").mean())
    if no_cat_ratio < 0.10:
        print(f"\n[第4层] 跳过（无类目竞品占比 {no_cat_ratio:.0%} < 10%，向量层已接管；"
              "层4 仅保留给无类目店）")
        m4 = []
    else:
        print("\n[第4层] 模糊兜底(TF-IDF + 图片pHash) -> 人工池...")
        t0 = time.time()
        m4, used_cidx = layer4_fuzzy(dfc, dfo, set(used_cidx), set(used_oid))
        print(f"  候选 {len(m4)} 条  耗时 {time.time()-t0:.1f}s")

    # ---- 输出（各层原始结果先落盘）----
    print("\n写出各层原始结果文件...")
    df1 = pd.DataFrame(m1)
    df2 = pd.DataFrame(m2)
    df_sem = pd.DataFrame(m_sem)
    df3 = pd.DataFrame(m3)
    df35 = pd.DataFrame(m35)
    df4 = pd.DataFrame(m4)

    df1.to_excel(os.path.join(OUT_DIR, "对齐_第1层_条码.xlsx"), index=False)
    df2.to_excel(os.path.join(OUT_DIR, "对齐_第2层_强组合.xlsx"), index=False)
    df_sem.to_excel(os.path.join(OUT_DIR, "对齐_第2.5层_品级语义_人工池.xlsx"), index=False)
    df3.to_excel(os.path.join(OUT_DIR, "对齐_第3层_弱组合_人工池.xlsx"), index=False)
    df35.to_excel(os.path.join(OUT_DIR, "对齐_第3.5层_关键字候选_人工池.xlsx"), index=False)
    pd.DataFrame(m36).to_excel(os.path.join(OUT_DIR, "对齐_第3.6层_向量候选_人工池.xlsx"), index=False)
    df4.to_excel(os.path.join(OUT_DIR, "对齐_第4层_模糊_人工池.xlsx"), index=False)

    # ---- 冲突消解（同一竞品品ID命中多个自家品，保留置信度最高）----
    print("\n[冲突消解] 同一竞品品ID多匹配去重...")
    parts_all = [d for d in [df1, df2, df3, df4] if len(d)]
    all_matches = pd.concat(parts_all, ignore_index=True) if parts_all else pd.DataFrame()
    if len(all_matches):
        resolved, conflict_log = resolve_conflicts(all_matches)
    else:
        resolved = all_matches
        conflict_log = pd.DataFrame(columns=CONFLICT_LOG_COLUMNS)
    n_conflict = len(conflict_log)
    print(f"  消解前 {len(all_matches)} 条 -> 消解后 {len(resolved)} 条 "
          f"(冲突日志 {n_conflict} 条)")
    conflict_log_path = os.path.join(OUT_DIR, "对齐_冲突日志.xlsx")
    conflict_log.to_excel(conflict_log_path, index=False)

    # 按匹配层拆分消解后结果
    if len(resolved) and "匹配层" in resolved.columns:
        resolved["匹配层"] = pd.to_numeric(resolved["匹配层"], errors="coerce").fillna(0).astype(int)
        df1r = resolved[resolved["匹配层"] == 1].copy()
        df2r = resolved[resolved["匹配层"] == 2].copy()
        df3r = resolved[resolved["匹配层"] == 3].copy()
        df4r = resolved[resolved["匹配层"] == 4].copy()
    else:
        df1r = df2r = df3r = df4r = pd.DataFrame()

    # ---- 价格合理性校验 ----
    print("\n[价格校验] 第3/4层价差率超±50%降级到人工池...")
    comp_lookup, own_lookup = build_price_lookup(dfc, dfo)

    # ---- 关联 SPU 组 ID（竞品Sku_id -> 竞品SPU组ID, 自家SKUID -> 自家SPU组ID）----
    print("\n[SPU关联] 为汇总表/人工池关联SPU组ID...")
    comp_sku_to_spu = {}
    for _, r in comp_spu.iterrows():
        for sku in str(r["SKU列表"]).split(","):
            if sku:
                comp_sku_to_spu[sku] = r["SPU组ID"]
    own_sku_to_spu = {}
    for _, r in own_spu.iterrows():
        for sku in str(r["SKU列表"]).split(","):
            if sku:
                own_sku_to_spu[sku] = r["SPU组ID"]

    def _map_comp_spu(row):
        return comp_sku_to_spu.get(_clean(row.get("竞品Sku_id")), "")

    def _map_own_spu(row):
        return own_sku_to_spu.get(_clean(row.get("自家SKUID")), "")

    # 汇总表 = 第1+2层 + 品内规格自动确认（自动确认，不降级，状态=未校验）
    parts_sum = [d for d in [df1r, df2r] if len(d)]
    if len(m_spu_auto):
        parts_sum.append(pd.DataFrame(m_spu_auto))
    summary = pd.concat(parts_sum, ignore_index=True) if parts_sum else pd.DataFrame()
    summary = validate_prices(summary, comp_lookup, own_lookup)
    # 关联 SPU 组 ID
    if len(summary):
        summary["竞品SPU组ID"] = summary.apply(_map_comp_spu, axis=1)
        summary["自家SPU组ID"] = summary.apply(_map_own_spu, axis=1)
    else:
        summary["竞品SPU组ID"] = []
        summary["自家SPU组ID"] = []
    summary_path = os.path.join(OUT_DIR, "对齐_汇总.xlsx")
    summary.to_excel(summary_path, index=False)

    # 人工确认池 = 品级语义 + 第3+4层 + 关键字候选 + 向量候选 + 品级候选（做价格校验，异常的标记降级原因）
    parts_rev = [d for d in [df_sem, df3r, df4r] if len(d)]
    if len(m35):
        parts_rev.append(pd.DataFrame(m35))
    if len(m36):
        parts_rev.append(pd.DataFrame(m36))
    if len(m_spu_cand):
        parts_rev.append(pd.DataFrame(m_spu_cand))
    review = pd.concat(parts_rev, ignore_index=True) if parts_rev else pd.DataFrame()
    # 兜底去重：同一竞品 Sku_id 跨层只保留置信度最高的一条
    # （各候选层已写回 used_cidx，正常不重复；此为防御性兜底，防旧产物/边缘情况）
    if len(review) and "竞品Sku_id" in review.columns:
        review["__conf_num"] = pd.to_numeric(review["置信度"], errors="coerce").fillna(0.0)
        review = (review.sort_values("__conf_num", ascending=False)
                        .drop_duplicates("竞品Sku_id", keep="first")
                        .drop(columns=["__conf_num"]))
    # 聚焦过滤：按**有效销量 = 成交 + λ×收藏**（λ 见 sales_text）剔除没量的品。
    # 收藏不剔除、只降权：35 人想买折成 7，仍够进池；只有零星收藏的才掉出去。
    if len(review) and "竞品Sku_id" in review.columns:
        import sales_text as _STx
        comp_sales = {}
        sales_col = "月销量" if "月销量" in dfc.columns else ("月售" if "月售" in dfc.columns else None)
        if sales_col:
            for _, r in dfc.iterrows():
                sid = _clean(r.get("Sku_id"))
                if sid:
                    comp_sales[sid] = _STx.eff(r.get(sales_col))
        if comp_sales:
            n_before = len(review)
            review["__sales"] = review["竞品Sku_id"].map(comp_sales).fillna(0.0)
            review = review[review["__sales"] >= MIN_SALES_FOCUS].drop(columns=["__sales"])
            print(f"  聚焦过滤：有效销量<1 移出人工池 {n_before - len(review)} 条"
                  f"（有效销量=成交+λ×收藏，保留≥{MIN_SALES_FOCUS}）")
    review = validate_prices(review, comp_lookup, own_lookup)
    if len(review):
        review["建议动作"] = "确认"
        review["确认结果"] = ""           # 人工填写: 确认 / 拒绝 / 标记错配
        review["人工备注"] = ""
        review["降级原因"] = ""
        if "候选自家品" not in review.columns:
            review["候选自家品"] = ""
        if "品级置信度" not in review.columns:
            review["品级置信度"] = ""
        review["候选自家品"] = review["候选自家品"].fillna("")
        review["品级置信度"] = review["品级置信度"].fillna("")
        review["竞品SPU组ID"] = review.apply(_map_comp_spu, axis=1)
        review["自家SPU组ID"] = review.apply(_map_own_spu, axis=1)
        mask = review["价格校验状态"] == "异常"
        review.loc[mask, "降级原因"] = "价格异常待确认"
        n_downgrade = int(mask.sum())
        # 向量高置信自动确认（M5·2.4a 试点）：池落盘前写入，标记随池文件一起落盘
        _base_now = os.path.basename(os.path.normpath(OUT_DIR))
        _shop_now = _base_now[len("试跑_"):] if _base_now.startswith("试跑_") else LIB_DEFAULT_SHOP
        n_vec_auto = auto_confirm_vector(review, _shop_now)
    else:
        n_downgrade = 0
        n_vec_auto = 0
    review_path = os.path.join(OUT_DIR, "对齐_人工确认池.xlsx")
    review.to_excel(review_path, index=False)
    print(f"  人工池 {len(review)} 条，价格异常降级 {n_downgrade} 条"
          + (f"，向量自动确认 {n_vec_auto} 条" if n_vec_auto else ""))

    # ---- 统计 ----
    n_comp = len(dfc)
    auto_raw = len(df1) + len(df2)               # 4层原始自动确认数
    auto = len(summary)                          # 消解后自动确认数
    pool = len(review)
    print("\n" + "=" * 60)
    print("对齐总览统计")
    print("=" * 60)
    print(f"竞品总行数            : {n_comp}")
    print(f"[SPU层] 竞品SPU组      : {len(comp_spu)}")
    print(f"[SPU层] 自家SPU组      : {len(own_spu)}")
    print(f"[SPU层] SPU对齐数      : {len(spu_pairs)}")
    print(f"[品级细化] 品内规格自动确认: {len(m_spu_auto)}")
    print(f"[品级细化] 品级候选(人工)  : {len(m_spu_cand)}")
    print(f"[品级细化] 缺规格建议数    : {len(missing_spec_rows)}")
    print(f"第1层 条码匹配(原始)  : {len(df1)}")
    print(f"第2层 强组合匹配(原始): {len(df2)}")
    print(f"第3层 弱组合候选(原始): {len(df3)}")
    print(f"第3.5层 关键字候选(原始): {len(df35)}")
    print(f"第3.6层 向量候选(原始) : {len(m36)}")
    print(f"第4层 模糊候选(原始)  : {len(df4)}")
    print(f"冲突消解丢弃         : {n_conflict} 条")
    print(f"自动确认(消解后)      : {auto}  (原始 {auto_raw})")
    print(f"自动匹配率(对竞品行)  : {auto/n_comp*100:.2f}%  ({auto}/{n_comp})")
    print(f"人工确认池规模        : {pool}")
    print(f"  其中价格异常降级    : {n_downgrade} 条")
    print(f"前4层未命中竞品行     : {n_comp - len(used_cidx)}")
    print(f"\n输出文件:")
    for f in ["对齐_SPU级.xlsx", "对齐_SPU级_缺规格建议.xlsx",
              "对齐_第1层_条码.xlsx", "对齐_第2层_强组合.xlsx",
              "对齐_第3层_弱组合_人工池.xlsx", "对齐_第4层_模糊_人工池.xlsx",
              "对齐_冲突日志.xlsx", "对齐_人工确认池.xlsx", "对齐_汇总.xlsx"]:
        p = os.path.join(OUT_DIR, f)
        print(f"  {p}  {'存在' if os.path.exists(p) else '缺失'}")

    # ---- 抽检 ----
    print("\n=== 抽检 ===")
    # 0) SPU 层抽检
    if len(spu_pairs_df):
        r = spu_pairs_df.iloc[0]
        print(f"[SPU对齐抽检] {r.get('竞品SPU组ID')} <-> {r.get('自家SPU组ID')} "
              f"类目={r.get('三级类目')} 品牌={r.get('品牌')} 型号={r.get('型号')} "
              f"竞品SKU数={r.get('竞品SKU数')} 自家SKU数={r.get('自家SKU数')}")
    if len(spu_refine_df):
        r = spu_refine_df.iloc[0]
        print(f"[SPU内SKU对齐抽检] 竞品Sku_id={r.get('竞品Sku_id')} "
              f"-> 自家SKUID={r.get('自家SKUID')} "
              f"竞品规格={r.get('竞品规格键')} 自家规格={r.get('自家规格键')}")
    if len(missing_spec_df):
        r = missing_spec_df.iloc[0]
        print(f"[缺规格建议抽检] SPU组ID={r.get('SPU组ID')} "
              f"品牌={r.get('品牌')} 型号={r.get('型号')} "
              f"建议内容={r.get('建议内容')}")
    # 1) 条码匹配抽检
    if len(df1):
        s = df1.iloc[0]
        print(f"[第1层抽检] 竞品Sku_id={s.get('竞品Sku_id')} 条码={s.get('竞品条形码')} "
              f"-> 自家SKUID={s.get('自家SKUID')} 条码={s.get('自家商品条码')} "
              f"依据={s.get('匹配依据')} 置信={s.get('置信度')}")
    # 2) 公牛 G32 抽检
    g32 = dfc[(dfc["品牌（抽取）"] == "公牛") & (dfc["型号"].str.upper() == "G32")]
    print(f"[公牛G32] 竞品命中 {len(g32)} 行")
    if len(g32):
        gids = set(g32["Sku_id"].astype(str).tolist())
        hit2 = df2[df2["竞品Sku_id"].astype(str).isin(gids)] if len(df2) else df2
        print(f"  其中第2层匹配 {len(hit2)} 条")
        if len(hit2):
            r = hit2.iloc[0]
            print(f"  样例: 竞品Sku_id={r.get('竞品Sku_id')}({r.get('竞品商品名','')[:30]}) "
                  f"-> 自家SKUID={r.get('自家SKUID')}({r.get('自家商品名','')[:30]}) "
                  f"依据={r.get('匹配依据')} 竞品规格={r.get('竞品规格键')} 自家规格={r.get('自家规格键')}")
        # G32 在第3层候选
        hit3 = df3[df3["竞品Sku_id"].astype(str).isin(gids)] if len(df3) else df3
        print(f"  其中第3层候选 {len(hit3)} 条")
        # G32 SPU 对齐抽检
        if len(spu_pairs_df):
            g32_spu = spu_pairs_df[
                (spu_pairs_df["品牌"] == "公牛")
                & (spu_pairs_df["型号"].astype(str).str.upper() == "G32")
            ]
            print(f"  其中SPU对齐 {len(g32_spu)} 对")
            if len(g32_spu):
                r = g32_spu.iloc[0]
                print(f"  SPU样例: {r.get('竞品SPU组ID')} <-> {r.get('自家SPU组ID')} "
                      f"竞品SKU数={r.get('竞品SKU数')} 自家SKU数={r.get('自家SKU数')}")
    # 3) 第3层同系列不同规格
    if len(df3):
        r = df3.iloc[0]
        print(f"[第3层抽检] 竞品型号={r.get('竞品型号')} 规格={r.get('竞品规格键')} "
              f"-> 自家规格={r.get('自家规格键')} 置信={r.get('置信度')}")
    # 4) 第4层样例
    if len(df4):
        r = df4.iloc[0]
        print(f"[第4层抽检] 文本={r.get('文本相似度')} 图片={r.get('图片相似度')} "
              f"融合={r.get('融合分数')} 依据={r.get('匹配依据')}")
        # 图url同源 / pHash距离 统计
        if "图url同源" in df4.columns:
            url_vc = df4["图url同源"].fillna("无url").value_counts().to_dict()
            n_url_match = int(url_vc.get("是", 0))
            print(f"[第4层图url同源] {url_vc}  同源命中 {n_url_match} 条")
        if "pHash距离" in df4.columns:
            phash_series = df4["pHash距离"]
            phash_valid = phash_series[phash_series.apply(lambda x: isinstance(x, (int, float)))]
            if len(phash_valid):
                print(f"[第4层pHash距离] 计算样本 {len(phash_valid)} 条, "
                      f"min={phash_valid.min()} max={phash_valid.max()} "
                      f"mean={phash_valid.mean():.2f}")

    # 总览统计 xlsx
    export_overview(dfc)

    # ---- 两层汇总库合并 ----
    print("\n[汇总库] 合并本次结果到 对齐汇总库.xlsx ...")
    merge_to_summary_library(dfc, summary, spu_pairs, own_spu)
    print("\n完成。")


# ===========================================================================
# 总览统计输出（可独立运行，避免重跑全量对齐）
# ===========================================================================
def export_overview(dfc=None):
    """输出 <OUT_DIR>/对齐_总览统计.xlsx，含两个 sheet：
       - 匹配统计：层级/匹配数/置信度区间/是否自动确认（含品级细化口径）
       - 未匹配竞品清单：未被任何环节覆盖的竞品行
    统计口径以 OUT_DIR 下真实产出文件为准：
      自动确认 = 对齐_汇总（第1+2层+品内规格）
      人工池   = 对齐_人工确认池（第3+4层+品级候选）
      已覆盖   = 汇总 + 人工池 + 已对齐品组的全部竞品SKU（品对上、规格缺口走建议）
    可独立运行（python align.py --overview [--comp ... --outdir ...]）。
    """
    print("\n" + "=" * 60)
    print("生成对齐总览统计 xlsx")
    print("=" * 60)

    def _read(path, **kw):
        if not os.path.exists(path):
            return pd.DataFrame()
        try:
            return pd.read_excel(path, **kw)
        except Exception:
            return pd.DataFrame()

    # 读取4层输出文件
    df1 = _read(os.path.join(OUT_DIR, "对齐_第1层_条码.xlsx"))
    df2 = _read(os.path.join(OUT_DIR, "对齐_第2层_强组合.xlsx"))
    df3 = _read(os.path.join(OUT_DIR, "对齐_第3层_弱组合_人工池.xlsx"))
    df4 = _read(os.path.join(OUT_DIR, "对齐_第4层_模糊_人工池.xlsx"))
    # 品级细化产出
    summary_f = _read(os.path.join(OUT_DIR, "对齐_汇总.xlsx"), dtype=str)
    pool_f = _read(os.path.join(OUT_DIR, "对齐_人工确认池.xlsx"), dtype=str)
    spu_refine_f = _read(os.path.join(OUT_DIR, "对齐_SPU细化.xlsx"),
                         sheet_name="品对品", dtype=str)

    # 读取竞品归一化表（中英文列名兼容）
    if dfc is None:
        dfc = load_comp_norm(COMP_NORM)

    n_comp = len(dfc)
    n_l1, n_l2, n_l3, n_l4 = len(df1), len(df2), len(df3), len(df4)
    # 品内规格自动确认 = 汇总 - 第1层 - 第2层（汇总=1+2层+品内规格）
    n_spu_spec = max(0, len(summary_f) - n_l1 - n_l2) if len(summary_f) else 0
    # 品级候选 = 人工池 - 第3层 - 第4层
    n_spu_cand = max(0, len(pool_f) - n_l3 - n_l4) if len(pool_f) else 0
    n_auto = len(summary_f)      # 自动确认合计（含品内规格）
    n_pool = len(pool_f)         # 人工池合计（含品级候选）
    n_unmatched = max(0, n_comp - n_auto - n_pool)

    # ---- sheet1: 匹配统计 ----
    stats = pd.DataFrame([
        {"层级": "第1层 条码精确匹配", "匹配数": n_l1,
         "置信度区间": f"{CONF_BARCODE:.2f}", "是否自动确认": "是"},
        {"层级": "第2层 强组合键匹配", "匹配数": n_l2,
         "置信度区间": f"{CONF_STRONG:.2f}", "是否自动确认": "是"},
        {"层级": "品内规格(品级细化)", "匹配数": n_spu_spec,
         "置信度区间": f"{CONF_SPU_SPEC:.2f}", "是否自动确认": "是"},
        {"层级": "第3层 弱组合(人工池)", "匹配数": n_l3,
         "置信度区间": f"{CONF_WEAK:.2f}", "是否自动确认": "否(人工池)"},
        {"层级": "第4层 模糊兜底(人工池)", "匹配数": n_l4,
         "置信度区间": f"<{CONF_FUZZY_CAP:.2f}", "是否自动确认": "否(人工池)"},
        {"层级": "品级候选(人工池)", "匹配数": n_spu_cand,
         "置信度区间": f"<{CONF_SPU_FALLBACK_CAP:.2f}", "是否自动确认": "否(人工池)"},
        {"层级": "自动确认匹配合计", "匹配数": n_auto,
         "置信度区间": "-", "是否自动确认": "是"},
        {"层级": "人工确认池合计", "匹配数": n_pool,
         "置信度区间": "-", "是否自动确认": "否"},
        {"层级": "未覆盖竞品行(未自动确认且未进人工池)", "匹配数": n_unmatched,
         "置信度区间": "-", "是否自动确认": "-"},
        {"层级": "竞品总行数", "匹配数": n_comp,
         "置信度区间": "-", "是否自动确认": "-"},
    ])

    # ---- sheet2: 未匹配竞品清单 ----
    # 已覆盖 = 汇总 + 人工池 的竞品Sku_id + 已对齐品组的竞品SKU列表
    covered_skus = set()
    for dfx in (summary_f, pool_f):
        if len(dfx) and "竞品Sku_id" in dfx.columns:
            covered_skus.update(
                s for s in dfx["竞品Sku_id"].astype(str).str.strip() if s)
    if len(spu_refine_f) and "竞品SKU列表" in spu_refine_f.columns:
        for lst in spu_refine_f["竞品SKU列表"].astype(str):
            covered_skus.update(s.strip() for s in lst.split(",") if s.strip())
    dfc["_sku_str"] = dfc["Sku_id"].astype(str).str.strip()
    unmatched = dfc[~dfc["_sku_str"].isin(covered_skus)].copy()

    unmatched_list = pd.DataFrame({
        "商品ID": unmatched["商品ID"].astype(str),
        "商品名": unmatched.get("商品名字", unmatched.get("清洗后商品名", "")).fillna(""),
        "类目": unmatched["类目三级_归一"].fillna("") if "类目三级_归一" in unmatched.columns else "",
        "品牌": unmatched["品牌（抽取）"].fillna("") if "品牌（抽取）" in unmatched.columns else "",
        "型号": unmatched["型号"].fillna("") if "型号" in unmatched.columns else "",
    })

    out_path = os.path.join(OUT_DIR, "对齐_总览统计.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        stats.to_excel(writer, sheet_name="匹配统计", index=False)
        unmatched_list.to_excel(writer, sheet_name="未匹配竞品清单", index=False)

    print(f"  第1层: {n_l1}  第2层: {n_l2}  品内规格: {n_spu_spec}")
    print(f"  第3层: {n_l3}  第4层: {n_l4}  品级候选: {n_spu_cand}")
    print(f"  自动确认: {n_auto}  人工池: {n_pool}  未覆盖: {n_unmatched}  竞品总: {n_comp}")
    print(f"  -> {out_path}")
    print(f"     sheet1 匹配统计: {len(stats)} 行")
    print(f"     sheet2 未匹配竞品清单: {len(unmatched_list)} 行")


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--overview":
        # 仅生成总览统计，不重跑全量对齐（支持 --comp/--outdir 指定目录）
        import argparse
        _p = argparse.ArgumentParser(description="仅生成对齐总览统计")
        _p.add_argument("--overview", action="store_true")
        _p.add_argument("--comp", help="竞品归一化表路径（默认 output/竞品_归一化后.xlsx）")
        _p.add_argument("--outdir", help="输出目录（默认 output/）")
        _a, _ = _p.parse_known_args()
        if _a.comp:
            COMP_NORM = _a.comp
        if _a.outdir:
            OUT_DIR = _a.outdir
        export_overview()
    else:
        main()

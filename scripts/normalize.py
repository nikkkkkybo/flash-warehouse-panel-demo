# -*- coding: utf-8 -*-
"""
阶段1 字段归一化模块 (Task 1.1-1.5)
对竞品表、自家表做归一化，输出到 output/ 目录。

Task 1.1 商品名清洗器
Task 1.2 品牌抽取器
Task 1.3 型号抽取器
Task 1.4 规格参数抽取
Task 1.5 类目对齐

输出:
  output/竞品_归一化后.xlsx
  output/自家_归一化后.xlsx
  output/类目映射表.xlsx
"""
import os
import re
import sys
import difflib
import pandas as pd
import numpy as np

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from brand_extractor_v2 import extract_brand_v2
# 竞店标准库中文表头 -> 英文内部名（读标准库链路时转回）
from import_adapter import CN_STD

COMP_PATH = os.path.join(BASE, "竞品_示例竞店B商品库_美团导出.xlsx")
OWN_PATH = os.path.join(BASE, "自家商品库_标准版本.xlsx")
OUT_DIR = os.path.join(BASE, "output")
os.makedirs(OUT_DIR, exist_ok=True)

# 两表列名映射（统一逻辑用）
COMP = {
    "name": "商品名字",
    "spec": "规格",
    "brand": None,          # 竞品无品牌列
    "cat3": "美团后台三级类",
}
OWN = {
    "name": "商品名称",
    "spec": "规格名称",
    "brand": "品牌名称",
    "cat3": "美团闪购类目三级",
}

# ===========================================================================
# Task 1.1 商品名清洗器
# ===========================================================================
PROMO_WORDS = [
    "特价", "现货", "限时", "秒杀", "包邮", "促销", "热卖", "爆款", "清仓",
    "亏本", "正品", "行货", "厂家直销", "直供", "新品", "热销", "特惠",
    "优惠", "打折", "满减", "赠品", "免邮", "顺丰包邮", " economic", "亏本清仓",
    "送电池", "送礼品", "送配件", "送", "免费", "限时秒杀", "今日特价",
]
# 已知电器/数码缩写（统一大写）
ACRONYMS = ["LED", "USB", "PVC", "RGB", "DC", "AC", "PD", "QC", "HDMI",
            "WIFI", "Type-C", "TypeC", "NPC", "UVC", "COB", "SMD"]


def remove_brackets(s: str) -> str:
    """去除中英文括号 / 方括号 / 【】 内的内容（支持混合半全角括号与嵌套）。"""
    prev = None
    while prev != s:
        prev = s
        # 混合半全角括号对：开 ( 或 （ ，闭 ) 或 ）
        s = re.sub(r"[(（][^()（）]*[)）]", "", s)
        s = re.sub(r"【[^】]*】", "", s)
        s = re.sub(r"\[[^\]]*\]", "", s)
    return s


def unify_case(s: str) -> str:
    """统一大小写：型号（字母在数字前）大写，单位（数字后跟 w/v/l/a/mm/cm/ml）小写。"""
    # 1) 字母串紧邻数字之前 -> 大写（型号前缀，如 par30->PAR30, e27->E27, JLW2401 不变）
    s = re.sub(r"[a-zA-Z]+(?=\d)", lambda m: m.group(0).upper(), s)
    # 2) 数字后跟已知单位 -> 单位小写（35W->35w, 250V->250v, 1.6L->1.6l, 5mm->5mm）
    s = re.sub(
        r"(\d)\s*(mm|cm|ml|w|v|l|a)(?![a-zA-Z])",
        lambda m: m.group(1) + m.group(2).lower(),
        s, flags=re.IGNORECASE,
    )
    # 3) 已知缩写统一大写
    for ac in ACRONYMS:
        s = re.sub(re.compile(re.escape(ac), re.IGNORECASE), ac, s)
    return s


def clean_product_name(name) -> str:
    """Task 1.1 清洗商品名。返回清洗后字符串。"""
    if not isinstance(name, str):
        name = "" if (name is None or (isinstance(name, float) and np.isnan(name))) else str(name)
    s = name.strip()
    if not s:
        return ""
    # 去括号内容
    s = remove_brackets(s)
    # 去促销词（长词优先）
    for w in sorted(PROMO_WORDS, key=len, reverse=True):
        s = s.replace(w, "")
    # 去重复词（按空格/斜杠切分，保留首次出现）
    tokens = [t for t in re.split(r"[\s/]+", s) if t]
    seen, out = set(), []
    for t in tokens:
        if t not in seen:
            seen.add(t)
            out.append(t)
    s = " ".join(out)
    # 统一大小写
    s = unify_case(s)
    # 清理多余空格与残留标点
    s = re.sub(r"\s+", " ", s).strip()
    return s


# ===========================================================================
# Task 1.2 品牌抽取器
# ===========================================================================
# 品牌种子词典（任务给定 + 五金/电器/日用常见品牌扩充）
BRAND_SEED = [
    "公牛", "松下", "九阳", "路易华登", "正泰", "德力西", "飞雕", "施耐德",
    "西门子", "罗格朗", "小米", "华为", "飞利浦", "美的", "格力", "海尔",
    "欧普照明", "欧普", "雷士", "佛山照明", "三雄极光", "鸿雁", "突破",
    "ABB", "西蒙", "卡西欧", "双鹿", "南孚", "超霸", "金普达", "扬子", "康佳",
    "日丰", "得力", "晨光", "德佑", "舒客", "佳速冲", "苏泊尔", "爱仕达",
    "荣事达", "小熊", "北鼎", "摩飞", "米家", "红心", "钱龙", "利德",
    "TCL", "奥克斯", "绿联", "倍思", "安克", "倍量", "久量", "宜家",
    "德铂", "韩派", "奥妙", "彩虹", "飞科", "奔腾", "万和", "万家乐",
    "方太", "老板", "华帝", "帅康", "樱花", "史密斯", "神火", "清风",
    "心相印", "维达", "植护", "洁柔", "立白", "蓝月亮", "超能", "威露士",
    "舒肤佳", "冷酸灵", "黑人", "高露洁", "佳洁士", "云南白药", "隆力奇",
    "奇强", "植护", "清风", "小米", "红米", "罗马仕", "倍斯特", "罗马仕",
    "钱龙", "利德", "鸿雁", "突破电气", "公牛", "子弹头", "航嘉", "长城",
    "先科", "新飞", "容声", "美菱", "志高", "格兰仕", "海信", "创维",
    "康宝", "华凌", "奥马", "吉德", "奇迪", "东芝", "三洋", "夏普",
]
# 描述性词，正则兜底时排除（避免误判为品牌）
NON_BRAND_WORDS = {
    "食品级", "多功能", "不锈钢", "家用", "商用", "一次性", "通用", "大功率",
    "加厚", "防滑", "高温", "高档", "耐用", "经济", "实用", "升级", "新款",
    "绝缘", "防水", "防烫", "防爆", "防爆", "国标", "纯铜", "全铜", "环保",
    "便携", "迷你", "智能", "高清", "无线", "有线", "可充电", "充电",
}
# 产品品类词（正则兜底时若候选包含这些词则拒绝，避免把品类当品牌）
PRODUCT_NOUNS = {
    "开关", "插座", "浴霸", "灯泡", "电池", "面板", "插头", "螺丝", "膨胀",
    "吊环", "灯带", "灯条", "射灯", "轨道", "船型", "翘板", "按钮", "接头",
    "接线", "端子", "继电器", "接触器", "断路器", "漏电", "配电", "插排",
    "锅具", "筷子", "餐具", "水龙头", "门锁", "挂钩", "粘钩", "电线", "电缆",
}


def normalize_existing_brand(b) -> str:
    """自家品牌名归一：'公牛（BULL）' -> '公牛'；'其他品牌'/空 -> ''。"""
    if not isinstance(b, str):
        return ""
    s = b.strip()
    if s in ("", "其他品牌", "nan", "NaN"):
        return ""
    # 去掉中英文括号及内容
    s = re.sub(r"[(（][^()（）]*[)）]", "", s).strip()
    return s


def build_brand_set(dfo) -> set:
    """数据驱动构建品牌词典：自家表归一化后高频品牌 + 种子词典。"""
    brands = set(BRAND_SEED)
    if "品牌名称" in dfo.columns:
        raw = dfo["品牌名称"].dropna().astype(str)
        norm = raw.apply(normalize_existing_brand)
        norm = norm[norm != ""]
        vc = norm.value_counts()
        # 出现 >=2 次的视为可靠品牌
        for b, c in vc.items():
            if c >= 2 and b:
                brands.add(b)
    return brands


def extract_brand(cleaned_name, brand_set, existing_brand=""):
    """
    Task 1.2 品牌抽取。
    返回 (品牌, 置信度)。
    - 自家原品牌非空(非'其他品牌') -> 用归一化原值，置信度按是否命中词典判定
    - 否则从清洗后商品名抽：词典最长命中 / 正则兜底 / 未命中
    置信度取值：词典命中 / 正则命中 / 未命中
    """
    # 自家原值优先
    ex = normalize_existing_brand(existing_brand)
    if ex:
        if ex in brand_set:
            return ex, "词典命中"
        # 原值不在词典但仍保留（少见，因词典已纳入自家品牌）
        return ex, "词典命中"

    text = cleaned_name or ""
    if not text:
        return "", "未命中"
    # 词典最长优先匹配
    for b in sorted(brand_set, key=len, reverse=True):
        if b and b in text:
            return b, "词典命中"
    # 正则兜底 1：开头连续大写英文（>=2）
    m = re.match(r"^\s*([A-Z]{2,}[A-Za-z]*)", text)
    if m:
        return m.group(1), "正则命中"
    # 正则兜底 2：开头 2-4 个汉字 + 空格（真实品牌后常跟空格），排除描述词与品类词
    m2 = re.match(r"^([\u4e00-\u9fa5]{2,4})(?=\s)", text)
    if m2:
        cand = m2.group(1)
        if cand not in NON_BRAND_WORDS and not any(n in cand for n in PRODUCT_NOUNS):
            return cand, "正则命中"
    return "", "未命中"


# ===========================================================================
# Task 1.3 型号抽取器
# ===========================================================================
# 字母(可含连字符)+数字 组合；前置负向断言(?<!\d)：型号首字母前不能是数字
#   —— 避免 "12V24V" 中的 "V24V" 被误判为型号（V 实为电压单位后缀）
MODEL_PAT = re.compile(r"(?<!\d)[A-Za-z]+-?[A-Za-z0-9]*\d+[A-Za-z0-9]*")
# 数字+型（如 86型/118型）
MODEL_DIGIT_PAT = re.compile(r"\d+型")


def extract_model(name) -> str:
    """Task 1.3 型号抽取。字母+数字组合优先（大写），其次 数字+型。取首个命中。
    使用原始商品名（保留括号内信息，如 '（CR2450纽扣电池）'）。"""
    text = name or ""
    if not isinstance(text, str) or not text:
        return ""
    # 优先字母型型号
    for m in MODEL_PAT.finditer(text):
        tok = m.group(0)
        # 必须同时含字母与数字
        if re.search(r"[A-Za-z]", tok) and re.search(r"\d", tok):
            return tok.upper()
    # 其次 数字+型
    m = MODEL_DIGIT_PAT.search(text)
    if m:
        return m.group(0)
    return ""


# ===========================================================================
# Task 1.4 规格参数抽取
# ===========================================================================
COLOR_TEMP_MAP = [
    ("暖白光", "暖白光"),
    ("中性光", "中性光"),
    ("白光", "白光"),
    ("暖光", "暖光"),
    ("暖白", "暖白光"),
    ("暖黄光", "暖光"),
    ("黄光", "暖光"),
    ("正白", "白光"),
    ("冷白", "中性光"),
]


def _first_num(pattern, text, flags=0):
    m = re.search(pattern, text, flags)
    if m:
        return m.group(0)
    return ""


def extract_specs(name, spec_field=""):
    """
    Task 1.4 规格参数抽取。
    从 商品名(原始)+规格字段 抽取 功率/容量/尺寸/电压/颜色档位/开关档位。
    返回 dict。
    """
    text = (name if isinstance(name, str) else "") + " " + (spec_field if isinstance(spec_field, str) else "")
    # 功率 W：数字 + w/W（后不跟字母，避免 watt 等）
    power = _first_num(r"\d+(?:\.\d+)?\s*[wW](?![a-zA-Z])", text).lower()
    # 电压 V：数字 + v/V（后不跟字母）
    voltage = _first_num(r"\d+(?:\.\d+)?\s*[vV](?![a-zA-Z])", text).lower()
    # 容量 ml/L：优先 ml，再 L（L 后不跟字母，避免 led 干扰；且需前有数字）
    cap = ""
    m_ml = re.search(r"\d+(?:\.\d+)?\s*ml(?![a-zA-Z])", text, re.IGNORECASE)
    if m_ml:
        cap = m_ml.group(0).lower()
    else:
        m_l = re.search(r"\d+(?:\.\d+)?\s*l(?![a-zA-Z])", text, re.IGNORECASE)
        if m_l:
            cap = m_l.group(0).lower()
    # 尺寸 cm/mm
    size = _first_num(r"\d+(?:\.\d+)?\s*(?:cm|mm)(?![a-zA-Z])", text, re.IGNORECASE)
    if size:
        size = size.lower()
    # 颜色档位
    color_temp = ""
    for kw, norm in COLOR_TEMP_MAP:
        if kw in text:
            color_temp = norm
            break
    # 开关档位：(一/二/三/四/五/六/多)开 + (单/双)控，各取首个
    parts = []
    mk = re.search(r"(一|二|三|四|五|六|多)开", text)
    if mk:
        parts.append(mk.group(0))
    mc = re.search(r"(单|双)控", text)
    if mc:
        parts.append(mc.group(0))
    switch = "/".join(parts) if parts else ""
    return {
        "功率": power,
        "容量": cap,
        "尺寸": size,
        "电压": voltage,
        "颜色档位": color_temp,
        "开关档位": switch,
    }


# ===========================================================================
# Task 1.6 单位归一化
# ===========================================================================
def _fmt_num(x):
    """数值格式化：整数显示为整数；小数去掉无用尾零。"""
    if abs(x - round(x)) < 1e-9:
        return str(int(round(x)))
    # 最多6位小数，去尾零
    s = f"{x:.6f}".rstrip("0").rstrip(".")
    return s


def normalize_units(text):
    """
    Task 1.6 单位归一化。
    容量：ml 保留；L → ×1000 → ml；g 保留；kg → ×1000 → g
    长度：m 保留；cm → /100 → m；mm → /1000 → m；米 → m
    功率：W 保留（统一大写）；瓦 → W
    电压：V 保留（统一大写）；伏 → V
    只处理能明确识别的"数值+单位"模式，不破坏其他文本。
    """
    if not isinstance(text, str) or not text:
        return ""
    s = text
    # 1) kg → g（×1000）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[kK][gG](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1)) * 1000) + "g", s)
    # 2) L/l → ml（×1000）；后不跟字母避免 LED/锂 等干扰
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[lL](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1)) * 1000) + "ml", s)
    # 3) mm → m（/1000）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[mM][mM](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1)) / 1000) + "m", s)
    # 4) cm → m（/100）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[cC][mM](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1)) / 100) + "m", s)
    # 5) 米 → m（中文单位）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*米",
               lambda m: _fmt_num(float(m.group(1))) + "m", s)
    # 6) m/M → m（保留，统一小写；后不跟字母，避免误匹配 ml/mm/cm）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[mM](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1))) + "m", s)
    # 7) 瓦 → W
    s = re.sub(r"(\d+(?:\.\d+)?)\s*瓦",
               lambda m: _fmt_num(float(m.group(1))) + "W", s)
    # 8) w/W → W（统一大写）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[wW](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1))) + "W", s)
    # 9) 伏 → V
    s = re.sub(r"(\d+(?:\.\d+)?)\s*伏",
               lambda m: _fmt_num(float(m.group(1))) + "V", s)
    # 10) v/V → V（统一大写）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[vV](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1))) + "V", s)
    # 11) g/G → g（保留，统一小写；后不跟字母避免 gram 等）
    s = re.sub(r"(\d+(?:\.\d+)?)\s*[gG](?![a-zA-Z])",
               lambda m: _fmt_num(float(m.group(1))) + "g", s)
    return s


# ===========================================================================
# Task 1.7 术语归一化
# ===========================================================================
SYNONYMS_PATH = os.path.join(BASE, "config", "synonyms.yaml")
_SYNONYMS_CACHE = None


def _load_synonyms():
    """加载 config/synonyms.yaml，返回 {变体: 标准术语} dict。带缓存。"""
    global _SYNONYMS_CACHE
    if _SYNONYMS_CACHE is not None:
        return _SYNONYMS_CACHE
    import yaml
    syn = {}
    if os.path.exists(SYNONYMS_PATH):
        with open(SYNONYMS_PATH, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        for k, v in data.items():
            if isinstance(k, str) and isinstance(v, str) and k and v:
                syn[k] = v
    _SYNONYMS_CACHE = syn
    return syn


def normalize_terms(text):
    """
    Task 1.7 术语归一化：将变体术语替换为标准术语。
    读取 config/synonyms.yaml；按变体长度降序替换（长串优先，避免子串误替）。
    """
    if not isinstance(text, str) or not text:
        return ""
    syn = _load_synonyms()
    if not syn:
        return text
    # 长串优先，避免短串先替换破坏长串
    for variant in sorted(syn.keys(), key=len, reverse=True):
        if variant in text:
            standard = syn[variant]
            if standard != variant:
                text = text.replace(variant, standard)
    return text


# ===========================================================================
# Task 1.5 类目对齐
# ===========================================================================
def align_categories(dfc, dfo, comp_cat_col, own_cat_col):
    """
    Task 1.5 类目对齐。
    返回 (mapping_df, comp_unified_series, own_unified_series)。
    mapping_df: 三级类目 / 竞品计数 / 自家计数 / 是否一致 / 建议映射
    """
    cc = dfc[comp_cat_col].fillna("").astype(str)
    oc = dfo[own_cat_col].fillna("").astype(str)
    cset = set(x for x in cc.unique() if x)
    oset = set(x for x in oc.unique() if x)
    all_cats = sorted(cset | oset)

    cc_vc = cc.value_counts().to_dict()
    oc_vc = oc.value_counts().to_dict()

    # 统一映射：source -> unified
    unified = {}
    both = cset & oset
    only_c = cset - oset
    only_o = oset - cset
    # 共有 -> 自身
    for c in both:
        unified[c] = c
    # 仅竞品 -> 模糊匹配自家
    oset_list = sorted(oset)
    for c in only_c:
        best = difflib.get_close_matches(c, oset_list, n=1, cutoff=0.6)
        unified[c] = best[0] if best else c
    # 仅自家 -> 模糊匹配竞品
    cset_list = sorted(cset)
    for o in only_o:
        best = difflib.get_close_matches(o, cset_list, n=1, cutoff=0.6)
        unified[o] = best[0] if best else o

    # 构建映射表
    rows = []
    for cat in all_cats:
        in_c = cat in cset
        in_o = cat in oset
        if in_c and in_o:
            flag = "一致"
        elif in_c:
            flag = "仅竞品"
        else:
            flag = "仅自家"
        sug = unified.get(cat, cat)
        if flag != "一致" and sug == cat:
            sug_disp = "需人工映射"
        else:
            sug_disp = sug
        rows.append({
            "三级类目": cat,
            "竞品计数": int(cc_vc.get(cat, 0)),
            "自家计数": int(oc_vc.get(cat, 0)),
            "是否一致": flag,
            "建议映射": sug_disp,
        })
    mapping_df = pd.DataFrame(rows)
    comp_unified = cc.map(unified).fillna(cc)
    own_unified = oc.map(unified).fillna(oc)
    return mapping_df, comp_unified, own_unified


# ===========================================================================
# 主流程
# ===========================================================================
def normalize_table(df, cfg, brand_set, is_own):
    """对单表执行归一化，返回追加列后的 DataFrame。"""
    name_col = cfg["name"]
    spec_col = cfg["spec"]
    brand_col = cfg["brand"]
    cat_col = cfg["cat3"]

    out = df.copy()
    # 规格字段补空字符串
    spec_series = df[spec_col].fillna("") if spec_col in df.columns else pd.Series([""] * len(df))

    # Task 1.1 清洗商品名
    cleaned = df[name_col].apply(lambda x: clean_product_name(x))
    out["原始商品名"] = df[name_col]
    out["清洗后商品名"] = cleaned

    # Task 1.2 品牌
    if is_own and brand_col in df.columns:
        existing = df[brand_col]
    else:
        existing = pd.Series([""] * len(df))
    brand_res = [extract_brand_v2(n, e) for n, e in zip(df[name_col], existing)]
    out["品牌（抽取）"] = [r[0] for r in brand_res]
    out["品牌置信度"] = [r[1] for r in brand_res]

    # Task 1.3 型号（用原始商品名，保留括号内型号信息，如 '（CR2450纽扣电池）'）
    out["型号"] = df[name_col].apply(lambda x: extract_model(x))

    # Task 1.4 规格参数（用原始商品名+规格字段，保留括号内规格信息）
    specs = [extract_specs(n, s) for n, s in zip(df[name_col], spec_series)]
    for k in ["功率", "容量", "尺寸", "电压", "颜色档位", "开关档位"]:
        out[k] = [s[k] for s in specs]

    # Task 1.6 单位归一化：对功率/容量/尺寸/电压字段做单位换算与统一
    for k in ["功率", "容量", "尺寸", "电压"]:
        out[k] = out[k].apply(normalize_units)

    # Task 1.7 术语归一化：对清洗后商品名做术语替换（变体 → 标准术语）
    out["清洗后商品名"] = out["清洗后商品名"].apply(normalize_terms)

    return out, cat_col


def load_column_config(profile_name):
    """
    从 yaml 加载竞品列配置，返回 COMP 兼容的字典。
    配置文件: config/import_profiles/<profile_name>.yaml
    配置中的标准字段映射到 COMP 内部键:
      name -> name, spec -> spec, brand -> brand, category3 -> cat3
    其它字段忽略；缺失项保留 COMP 默认值。
    """
    import yaml
    profile_path = os.path.join(BASE, "config", "import_profiles", f"{profile_name}.yaml")
    if not os.path.exists(profile_path):
        raise FileNotFoundError(f"配置不存在: {profile_path}")
    with open(profile_path, "r", encoding="utf-8") as f:
        data = yaml.safe_load(f)
    mapping = (data or {}).get("mapping", {}) or {}
    field_to_comp = {"name": "name", "spec": "spec", "brand": "brand", "category3": "cat3"}
    comp = dict(COMP)
    for std_field, comp_key in field_to_comp.items():
        col = mapping.get(std_field)
        if col:
            comp[comp_key] = col
    return comp


def _self_test():
    """归一化补强自检：单位归一 + 术语归一。"""
    print("\n=== 归一化补强自检 ===")
    unit_cases = [
        ("550ml", "550ml"),
        ("0.55L", "550ml"),
        ("100cm", "1m"),
        ("12瓦", "12W"),
        ("12W", "12W"),
        ("220V", "220V"),
        ("1米", "1m"),
        ("0.5kg", "500g"),
    ]
    term_cases = [
        ("墙壁开关", "开关插座"),
        ("86面板", "86型"),
    ]
    ok = True
    for inp, exp in unit_cases:
        got = normalize_units(inp)
        flag = "OK" if got == exp else "FAIL"
        if got != exp:
            ok = False
        print(f"  normalize_units({inp!r}) = {got!r} (期望 {exp!r}) [{flag}]")
    for inp, exp in term_cases:
        got = normalize_terms(inp)
        flag = "OK" if got == exp else "FAIL"
        if got != exp:
            ok = False
        print(f"  normalize_terms({inp!r}) = {got!r} (期望 {exp!r}) [{flag}]")
    print(f"自检结果: {'全部通过' if ok else '存在失败'}\n")
    return ok


def main():
    import argparse
    parser = argparse.ArgumentParser(description="阶段1 字段归一化")
    parser.add_argument("--profile",
                        help="竞品来源名（从 config/import_profiles/<来源>.yaml 加载列配置覆盖 COMP 默认）")
    parser.add_argument("--config", help="直接指定 yaml 配置路径覆盖 COMP 默认")
    parser.add_argument("--comp-path", help="竞品输入表路径（默认 竞品_示例竞店B商品库_美团导出.xlsx）")
    parser.add_argument("--own-path", help="自家输入表路径（默认 自家商品库_标准版本.xlsx）")
    parser.add_argument("--comp-out", help="竞品归一化输出路径（默认 output/竞品_归一化后.xlsx）")
    parser.add_argument("--own-out", help="自家归一化输出路径（默认 output/自家_归一化后.xlsx）")
    parser.add_argument("--map-out", help="类目映射表输出路径（默认 output/类目映射表.xlsx）")
    parser.add_argument("--skip-own", action="store_true",
                        help="跳过自家表归一化（从 --own-out 或默认路径读取已有归一化结果）")
    parser.add_argument("--apply-map",
                        help="跳过类目对齐，改用指定类目映射表（output/类目映射表.xlsx）对竞品应用统一映射")
    args = parser.parse_args()

    # 归一化补强自检
    _self_test()

    comp_path = args.comp_path or COMP_PATH
    own_path = args.own_path or OWN_PATH
    comp_out = args.comp_out or os.path.join(OUT_DIR, "竞品_归一化后.xlsx")
    own_out = args.own_out or os.path.join(OUT_DIR, "自家_归一化后.xlsx")
    map_out = args.map_out or os.path.join(OUT_DIR, "类目映射表.xlsx")

    print("加载数据...")
    # ID 列强制按字符串读取，避免19位大整数被 float64 截断丢精度
    comp_dtype = {"类型ID": str, "商品ID": str, "Sku_id": str, "条形码": str,
                  "comp_id": str, "sku_id": str, "barcode": str}
    own_dtype = {"SPUID": str, "SKUID": str, "商品条码": str}
    dfc = pd.read_excel(comp_path, dtype=comp_dtype)
    # 兼容竞店标准库中文表头（转回内部英文列名）
    dfc = dfc.rename(columns={k: v for k, v in CN_STD.items() if k in dfc.columns})
    print(f"竞品: {dfc.shape}  <- {comp_path}")
    if args.skip_own:
        if not os.path.exists(own_out):
            raise FileNotFoundError(f"--skip-own 需要已有归一化结果: {own_out}")
        dfo = pd.read_excel(own_out, dtype=own_dtype)
        print(f"自家(复用): {dfo.shape}  <- {own_out}")
    else:
        dfo = pd.read_excel(own_path, dtype=own_dtype)
        print(f"自家: {dfo.shape}  <- {own_path}")

    # 决定竞品列配置：--config > --profile > 默认 COMP（兼容无参数调用）
    if args.config:
        import yaml
        with open(args.config, "r", encoding="utf-8") as f:
            data = yaml.safe_load(f)
        mapping = (data or {}).get("mapping", {}) or {}
        comp_cfg = dict(COMP)
        for std, key in (("name", "name"), ("spec", "spec"),
                         ("brand", "brand"), ("category3", "cat3")):
            col = mapping.get(std)
            if col:
                comp_cfg[key] = col
        print(f"使用配置: {args.config}")
    elif args.profile:
        comp_cfg = load_column_config(args.profile)
        print(f"使用 profile: {args.profile}")
    else:
        comp_cfg = COMP
        print("未指定配置，使用默认硬编码 COMP")
    print(f"竞品列配置: {comp_cfg}")

    # 构建品牌词典（数据驱动）
    brand_set = build_brand_set(dfo)
    print(f"品牌词典大小: {len(brand_set)}")

    # 归一化两表
    print("归一化竞品表...")
    dfc_n, comp_cat_col = normalize_table(dfc, comp_cfg, brand_set, is_own=False)
    if args.skip_own:
        print("跳过自家表归一化（复用已有结果）")
        dfo_n = dfo
    else:
        print("归一化自家表...")
        dfo_n, own_cat_col = normalize_table(dfo, OWN, brand_set, is_own=True)

    # Task 1.5 类目对齐（--apply-map 时改用统一映射表，保证多店口径一致）
    if args.apply_map:
        print(f"类目对齐: 使用统一映射表 {args.apply_map}（跳过模糊匹配）")
        import yaml
        mapping_df = pd.read_excel(args.apply_map)
        unified = {}
        for _, r in mapping_df.iterrows():
            sug = str(r.get("建议映射", ""))
            if sug and sug != "需人工映射":
                unified[str(r["三级类目"])] = sug
        cc = dfc_n[comp_cat_col].fillna("").astype(str)
        dfc_n["类目三级_归一"] = cc.map(unified).fillna(cc)
        if not args.skip_own:
            oc = dfo_n[own_cat_col].fillna("").astype(str)
            dfo_n["类目三级_归一"] = oc.map(unified).fillna(oc)
    else:
        print("类目对齐...")
        mapping_df, comp_uni, own_uni = align_categories(dfc, dfo, comp_cat_col, own_cat_col)
        dfc_n["类目三级_归一"] = comp_uni.values
        dfo_n["类目三级_归一"] = own_uni.values

    # 输出
    for p in [comp_out, own_out, map_out]:
        os.makedirs(os.path.dirname(p), exist_ok=True)
    print("写出文件...")
    dfc_n.to_excel(comp_out, index=False)
    if not args.skip_own:
        dfo_n.to_excel(own_out, index=False)
    if not args.apply_map:
        mapping_df.to_excel(map_out, index=False)
    print(f"-> {comp_out} ({dfc_n.shape})")
    if not args.skip_own:
        print(f"-> {own_out} ({dfo_n.shape})")
    if not args.apply_map:
        print(f"-> {map_out} ({mapping_df.shape})")

    # 简要统计
    print("\n=== 抽检统计 ===")
    print(f"竞品 品牌置信度分布: {dict(dfc_n['品牌置信度'].value_counts())}")
    print(f"自家 品牌置信度分布: {dict(dfo_n['品牌置信度'].value_counts())}")
    print(f"竞品 型号非空: {int(dfc_n['型号'].astype(bool).sum())}/{len(dfc_n)}")
    print(f"自家 型号非空: {int(dfo_n['型号'].astype(bool).sum())}/{len(dfo_n)}")
    print(f"竞品 颜色档位非空: {int(dfc_n['颜色档位'].astype(bool).sum())}")
    print(f"类目映射表: 共{len(mapping_df)}个三级类目, "
          f"一致{int((mapping_df['是否一致']=='一致').sum())}, "
          f"仅竞品{int((mapping_df['是否一致']=='仅竞品').sum())}, "
          f"仅自家{int((mapping_df['是否一致']=='仅自家').sum())}")


if __name__ == "__main__":
    main()

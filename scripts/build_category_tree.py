# -*- coding: utf-8 -*-
"""
自家商品分类树候选表生成（第1步：自动提取，供人工确认）

产出 output/自家商品分类树_候选确认.xlsx，三个 sheet：
  - 品名词候选：关键词库启用词 + 高频2-4字词段，按覆盖品组数降序
    （覆盖≥2品组的词是确认重点：可能是宽品名/泛词/属性词混入；
      覆盖1品组的点名词默认就是该品的品名）
  - 属性词候选：材质/特性/颜色/单位词表命中统计（供快速归类）
  - 说明：确认方法

确认方式（用户）：在「品名词候选」表按需填写 确认 列：
  品名词 / 父节点（宽品名，下面还有细分） / 属性（不参与匹配） / 忽略
  低频词可用 Excel 筛选后批量填写。
"""
import os
import re
from collections import defaultdict

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")
KW_FILE = os.path.join(BASE, "output", "自家关键词库.xlsx")
OUT_FILE = os.path.join(BASE, "output", "自家商品分类树_候选确认.xlsx")

# ---- 属性词表（自动识别用；最终归属以人工确认为准）----
UNITS = ["个", "只", "支", "套", "包", "袋", "盒", "瓶", "罐", "桶", "根", "条",
         "米", "卷", "张", "片", "块", "把", "副", "双", "对", "组", "箱", "台",
         "件", "批", "枚", "颗", "粒", "节", "提", "串", "捆", "面", "顶", "本"]
MATERIALS = ["不锈钢", "铝合金", "黄铜", "紫铜", "纯铜", "锌合金", "碳钢", "镀锌",
             "塑料", "尼龙", "橡胶", "硅胶", "陶瓷", "玻璃", "铸铁", "ABS", "PP",
             "PVC", "亚克力", "树脂", "皮质", "帆布"]
FEATURES = ["防水", "防滑", "加厚", "便携", "折叠", "静音", "耐磨", "耐高温",
            "绝缘", "阻燃", "防爆", "可调", "通用", "家用", "商用", "专业",
            "工业级", "环保", "抗压", "耐摔", "强力", "加粗", "加长", "加大",
            "迷你", "超大", "超强", "加硬", "防漏", "密封", "省力", "自动",
            "国标", "优选", "免打孔", "免钉", "即热", "简易", "无线", "有线",
            "电动", "手动", "新款", "升级", "加强", "特价", "促销", "家用"]
COLORS = ["白色", "黑色", "红色", "蓝色", "绿色", "黄色", "透明", "金色", "银色",
          "灰色", "彩色", "粉色", "紫色", "橙色", "棕色", "米色", "深色", "浅色",
          "白光", "暖白", "暖光", "中性光", "冷白", "自然光", "冷光"]
STOP = {"正品", "包邮", "家用", "通用", "专用", "新款", "批发", "加厚", "高配",
        "升级", "加强", "精品", "促销", "特价", "优惠", "赠品", "免费", "上门",
        "安装", "维修", "厂家", "直销", "现货", "下单", "客服", "联系", "咨询",
        "备注", "型号", "规格", "颜色", "款式", "左右", "上下", "系列", "套装",
        "闪购", "美团", "闪电仓", "热销", "爆款", "断货", "补货", "到货",
        "可选", "多规格", "随机", "包邮", "适用", "适配", "兼容", "支持", "含",
        "带", "送", "新", "好", "大", "小", "型"}

ATTR_ALL = UNITS + MATERIALS + FEATURES + COLORS


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    try:
        f = float(str(v).replace(",", "").strip())
        return f if f == f else 0.0
    except (ValueError, TypeError):
        return 0.0


def clean_name_for_words(name, brand=""):
    """清洗品名供词段提取：去括号/品牌/属性词/规格数字，剩中文核心段（空格分隔）。
    注意：单字单位词（个/套/袋/包…）只剥离品名结尾，不做全局替换——
    否则会破坏"垃圾袋/书包"类品名中的单位字。"""
    s = str(name)
    s = re.sub(r"[（(【\[][^）)】\]]*[）)】\]]", "", s)
    if brand:
        s = s.replace(str(brand), "")
    # 多字属性词全局替换（材质词除外：塑料/尼龙等是"塑料袋/尼龙扎带"的词根，替换会破坏品名）
    for w in FEATURES + COLORS + list(STOP):
        s = s.replace(w, "")
    # 单字单位词只剥离结尾（品名尾部计量单位）
    s = re.sub(r"[个只支套包袋盒瓶罐桶根条米卷张片块把副双对组箱台件批枚颗粒节提串捆面顶本]+$", "", s)
    s = re.sub(r"\d+\.?\d*[a-zA-Z%℃度VvWwMmKkGg]*", "", s)
    s = re.sub(r"\d+", "", s)
    s = re.sub(r"[^\u4e00-\u9fff]+", " ", s)
    return s


def load_spu_groups():
    """读自家表按 SPUID 分组：品组名=月销最高SKU名，月销=组内合计。返回 list[dict]。"""
    own = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    if "SPUID" not in own.columns:
        raise RuntimeError("自家表缺 SPUID 列")
    groups = []
    for spuid, g in own.groupby("SPUID"):
        spuid = _clean(spuid)
        if not spuid:
            continue
        g2 = g.copy()
        g2["__s"] = g2["月销量"].apply(_num)
        total_sales = float(g2["__s"].sum())
        rep = g2.loc[g2["__s"].idxmax()] if len(g2) else g2.iloc[0]
        groups.append({
            "品组ID": spuid,
            "品名": _clean(rep.get("商品名称")) or _clean(rep.get("清洗后商品名")),
            "类目": _clean(rep.get("类目三级_归一")) or _clean(rep.get("美团闪购类目三级")),
            "品牌": _clean(rep.get("品牌（抽取）")) or _clean(rep.get("品牌名称")),
            "月销": total_sales,
            "SKU数": len(g2),
        })
    return groups


def main():
    own = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    if "SPUID" not in own.columns:
        raise RuntimeError("自家表缺 SPUID 列")

    # ---- 品牌集合（自动识别品牌词用）：自家品牌列 + 高频常见品牌 ----
    brand_set = set()
    for c in ("品牌（抽取）", "品牌名称"):
        if c in own.columns:
            brand_set.update(str(v).strip() for v in own[c].dropna()
                             if str(v).strip() and str(v).lower() not in ("nan", "其他品牌"))
    COMMON_BRANDS = ["华为", "苹果", "小米", "公牛", "得力", "南孚", "飞利浦", "欧普",
                     "佛山照明", "德力西", "正泰", "施耐德", "西门子", "3M", "日丰",
                     "九阳", "美的", "格力", "海尔", "苏泊尔", "维达", "洁柔", "心相印",
                     "蓝月亮", "立白", "超能", "威露士", "舒肤佳", "长城", "博世",
                     "世达", "史丹利", "锐尔", "拓森", "联塑", "伟星", "公元", "中财"]
    brand_set.update(COMMON_BRANDS)

    def auto_type(w):
        """自动预分类：品牌/材质/特性/颜色/单位/停用 -> 直接建议，其余需人工确认。"""
        if w in brand_set:
            return "品牌词（建议忽略）"
        if w in MATERIALS:
            return "材质词（建议属性）"
        if w in FEATURES:
            return "特性词（建议属性）"
        if w in COLORS:
            return "颜色词（建议属性）"
        if w in UNITS:
            return "单位词（建议属性）"
        if w in STOP:
            return "停用词（建议忽略）"
        return ""

    # ---- 按 SPUID 分组（品组名=月销最高SKU名，月销=组内合计）----
    groups = load_spu_groups()
    print("自家品组数:", len(groups))
    gdf = pd.DataFrame(groups)
    gname = dict(zip(gdf["品组ID"], gdf["品名"]))
    gcat = dict(zip(gdf["品组ID"], gdf["类目"]))
    gsales = dict(zip(gdf["品组ID"], gdf["月销"]))

    # ---- 品名词候选 1：关键词库启用词 ----
    kw = pd.read_excel(KW_FILE, dtype=str).fillna("")
    kw = kw[kw["状态"].astype(str).isin(["", "启用"])]
    word_groups = defaultdict(set)      # 词 -> 品组ID集合
    for _, r in kw.iterrows():
        w = _clean(r.get("关键词"))
        gid = _clean(r.get("品组ID"))
        if w and gid:
            word_groups[w].add(gid)

    # ---- 品名词候选 2：完整中文词段补充（覆盖关键词库未覆盖的品组）----
    # 只取「完整连续段」（2-6字，非滑动窗口子串），且跨品组出现≥5次，
    # 避免"生间/多功/水器"这类切割碎片噪音。
    seg_groups = defaultdict(set)
    for g in groups:
        cleaned = clean_name_for_words(g["品名"], g["品牌"])
        for seg in cleaned.split():
            if 2 <= len(seg) <= 6 and not seg.startswith(("的", "了", "与", "和", "在")):
                seg_groups[seg].add(g["品组ID"])
    seg_cands = {w: gs for w, gs in seg_groups.items()
                 if w not in word_groups and len(gs) >= 5}

    # ---- 汇总统计 ----
    def stat_word(w, gids):
        # 实际覆盖面：自家品组名包含该词的品组（比词库稀疏关联数更真实）
        sub = gdf[gdf["品名"].astype(str).str.contains(re.escape(w), na=False)]
        n_groups = len(sub)
        n_skus = int(sub["SKU数"].sum())
        sales = float(sub["月销"].sum())
        # 类目分布
        cat_cnt = defaultdict(int)
        for c in sub["类目"].fillna("(无类目)"):
            cat_cnt[c] += 1
        cats = "; ".join("%s(%d组)" % (c, n) for c, n in
                         sorted(cat_cnt.items(), key=lambda kv: -kv[1])[:3])
        # 示例品名（月销最高品组，截断便于阅读）
        best_gid = sub.loc[sub["月销"].idxmax(), "品组ID"] if len(sub) else ""
        ex = gname.get(best_gid, "")
        return n_groups, n_skus, sales, cats, (ex[:30] + "…" if len(ex) > 30 else ex)

    rows = []
    for w, gids in word_groups.items():
        n_groups, n_skus, sales, cats, ex = stat_word(w, gids)
        at = auto_type(w)
        if at:
            sug = at
        elif n_groups >= 10:
            sug = "泛词?（确认：宽品名/属性/忽略）"
        elif n_groups >= 2:
            sug = "宽品名或细品名词（确认）"
        else:
            sug = "细品名词（默认）"
        rows.append({"候选词": w, "建议层级": sug, "覆盖品组数": n_groups,
                     "覆盖SKU数": n_skus, "月销合计": round(sales, 0),
                     "类目分布": cats, "示例品名": ex, "来源": "关键词库",
                     "确认": "", "备注": ""})
    for w, gids in seg_cands.items():
        n_groups, n_skus, sales, cats, ex = stat_word(w, gids)
        at = auto_type(w)
        sug = at if at else ("泛词?（确认）" if n_groups >= 10 else "细品名词（默认）")
        rows.append({"候选词": w, "建议层级": sug, "覆盖品组数": n_groups,
                     "覆盖SKU数": n_skus, "月销合计": round(sales, 0),
                     "类目分布": cats, "示例品名": ex, "来源": "高频词段",
                     "确认": "", "备注": ""})

    df_word = pd.DataFrame(rows).sort_values(
        ["覆盖品组数", "月销合计"], ascending=[False, False])
    df_word.insert(0, "序号", range(1, len(df_word) + 1))

    # 泛词重点确认 sheet：覆盖≥10 且非自动分类（树的骨架，需逐条定层级）
    df_key = df_word[(df_word["覆盖品组数"] >= 10)
                     & (~df_word["建议层级"].str.contains("建议"))].copy()

    # ---- 属性词候选 ----
    attr_rows = []
    for w in ATTR_ALL:
        if not w:
            continue
        gids = set()
        for g in groups:
            if w in g["品名"]:
                gids.add(g["品组ID"])
        if not gids:
            continue
        typ = ("单位" if w in UNITS else "材质" if w in MATERIALS
               else "特性" if w in FEATURES else "颜色")
        n_groups = len(gids)
        sales = float(gdf[gdf["品组ID"].isin(gids)]["月销"].sum())
        best_gid = max(gids, key=lambda gid: gsales.get(gid, 0))
        attr_rows.append({"候选词": w, "类型": typ, "覆盖品组数": n_groups,
                          "月销合计": round(sales, 0),
                          "示例品名": gname.get(best_gid, ""),
                          "确认": "", "备注": ""})
    df_attr = pd.DataFrame(attr_rows).sort_values(
        ["类型", "覆盖品组数"], ascending=[True, False])
    df_attr.insert(0, "序号", range(1, len(df_attr) + 1))

    # ---- 说明 ----
    note = pd.DataFrame([
        {"说明": "本表为「自家商品分类树」第1步：自动提取候选，请逐项/批量确认后固化。"},
        {"说明": "品名词候选：覆盖≥2品组的词是确认重点（可能为宽品名/泛词/属性词混入）；覆盖1品组的点名词默认即该品品名。"},
        {"说明": "确认列填写：品名词 / 父节点（宽品名，其下还有细分）/ 属性（不参与匹配）/ 忽略。"},
        {"说明": "属性词候选：材质/特性/颜色/单位词，命中即建议归「属性」；如需作品名词请改填。"},
        {"说明": "「类目分布」列显示该词主要归属类目，同一词跨类目时后续按「类目+品名词」建树。"},
        {"说明": "批量操作：可用 Excel 筛选（如覆盖品组数≥10）后整列批量填写确认。"},
    ])

    with pd.ExcelWriter(OUT_FILE) as w:
        df_key.to_excel(w, sheet_name="泛词重点确认", index=False)
        df_word.to_excel(w, sheet_name="品名词候选", index=False)
        df_attr.to_excel(w, sheet_name="属性词候选", index=False)
        note.to_excel(w, sheet_name="说明", index=False)

    print("产出:", OUT_FILE)
    print("  品名词候选 %d 条（关键词库 %d + 高频词段 %d）"
          % (len(df_word), len(word_groups), len(seg_cands)))
    print("  属性词候选 %d 条" % len(df_attr))
    print("  泛词重点确认（覆盖≥10，树骨架）: %d 条" % len(df_key))
    # 按建议分类统计（确认负担）
    for label in ["品牌词（建议忽略）", "材质词（建议属性）", "特性词（建议属性）",
                  "颜色词（建议属性）", "单位词（建议属性）", "停用词（建议忽略）"]:
        n = (df_word["建议层级"] == label).sum()
        print("    %s: %d 条" % (label, n))
    need = df_word[(~df_word["建议层级"].str.contains("建议"))
                   & (df_word["覆盖品组数"] >= 2)]
    print("  需人工确认（覆盖≥2品组且非自动分类）: %d 条" % len(need))
    print("  覆盖1品组的点名词（默认细品名词，无需逐条确认）: %d 条"
          % len(df_word[df_word["覆盖品组数"] == 1]))
    print("\n品名词候选 Top15（按覆盖品组数）:")
    print(df_word.head(15)[["候选词", "建议层级", "覆盖品组数", "月销合计",
                            "类目分布"]].to_string(index=False))
    print("\n属性词候选（前20）:")
    print(df_attr.head(20)[["候选词", "类型", "覆盖品组数"]].to_string(index=False))


def confirm_and_fix():
    """读取确认表「泛词重点确认」-> 固化 config/category_tree.yaml。
    解析规则：确认列为空 = 认可默认（宽品名/父节点）；
    确认列=属性/长度 = 属性词；确认列=其他词（如 充电线/补墙膏/螺丝）= 同义词归属。"""
    import yaml

    df = pd.read_excel(OUT_FILE, sheet_name="泛词重点确认", dtype=str).fillna("")
    broad, attrs, length, syn = [], [], [], {}
    for _, r in df.iterrows():
        w = _clean(r.get("候选词"))
        if not w:
            continue
        conf = _clean(r.get("确认"))
        if not conf:
            broad.append(w)
        elif conf in ("属性", "特性", "材质", "颜色"):
            attrs.append(w)
        elif conf in ("长度", "规格", "尺寸"):
            length.append(w)
        else:
            syn[w] = conf   # 同义词/归属：竞品叫法 -> 标准品名词

    cfg = {
        "version": 1,
        "confirmed_at": "2026-08-14",
        "说明": "自家商品分类树-泛词确认固化（人工确认结果）",
        "broad_words": sorted(broad),          # 宽品名（父节点）：竞品名含该词即进入该品分支
        "attribute_words": sorted(attrs + length),  # 属性词：剥离不参与品匹配
        "length_words": sorted(length),
        "word_synonyms": dict(sorted(syn.items())),  # 竞品叫法 -> 标准品名词
    }
    os.makedirs(os.path.join(BASE, "config"), exist_ok=True)
    out = os.path.join(BASE, "config", "category_tree.yaml")
    with open(out, "w", encoding="utf-8") as f:
        yaml.safe_dump(cfg, f, allow_unicode=True, sort_keys=False)

    print("固化完成:", out)
    print("  宽品名（父节点）: %d 条" % len(broad))
    print("  属性词: %d 条（其中长度类 %d）" % (len(attrs) + len(length), len(length)))
    print("  同义词/归属: %d 条" % len(syn))
    print("\n宽品名样例:", broad[:20])
    print("属性词:", attrs + length)
    print("同义词:", syn)


def build_tree_preview():
    """第2步：用已确认的宽品名/属性词/同义词生成「类目×宽品名×细品名词」树骨架。
    产出 output/自家商品分类树_骨架.xlsx（树骨架/未归组品组/说明）。"""
    import yaml

    cfg_path = os.path.join(BASE, "config", "category_tree.yaml")
    if not os.path.exists(cfg_path):
        print("请先运行 --confirm 固化确认结果")
        return
    cfg = yaml.safe_load(open(cfg_path, encoding="utf-8"))
    broad = list(cfg.get("broad_words", []))
    attrs = list(cfg.get("attribute_words", []))
    syn = cfg.get("word_synonyms", {})
    # 同义词标准名也视为宽品名（如 充电线/补墙膏/螺丝 未必在 broad_words 里）
    effective_broad = set(broad) | set(syn.values())

    # 关键词库词 -> 品组（细品名词优先来源：大模型点名词比规则剥离准）
    kw = pd.read_excel(KW_FILE, dtype=str).fillna("")
    kw = kw[kw["状态"].astype(str).isin(["", "启用"])]
    kw_by_gid = defaultdict(list)   # 品组ID -> 词列表（按长度降序）
    for _, r in kw.iterrows():
        w = _clean(r.get("关键词"))
        gid = _clean(r.get("品组ID"))
        if w and gid and len(w) >= 2:
            kw_by_gid[gid].append(w)
    for gid in kw_by_gid:
        kw_by_gid[gid] = sorted(set(kw_by_gid[gid]), key=len, reverse=True)

    # 常见螺丝/五金属性词（补充，避免"十字沉头"被当细品名词）
    EXTRA_ATTRS = ["十字", "沉头", "圆头", "扁头", "六角", "内六角", "外六角", "一字",
                   "米字", "钻尾", "自攻", "镀黑锌", "镀锌", "加长", "加硬", "细牙",
                   "粗牙", "蘑菇头", "大扁头", "平头", "尖头", "梅花", "内梅花"]

    groups = load_spu_groups()
    print("自家品组数:", len(groups))

    def match_broad(name):
        """宽品名匹配：取最长命中（含同义词标准名与同义词叫法）。返回宽品名或 ''。"""
        best, best_len = "", 0
        for w in effective_broad:
            if w in name and len(w) > best_len:
                best, best_len = w, len(w)
        for k, v in syn.items():
            # 自家品组名含竞品叫法（如"墙面修补"）也归到标准品名（补墙膏）
            if k in name and len(v) > best_len:
                best, best_len = v, len(v)
        return best

    def fine_word(name, broad_hit):
        """细品名词：优先关键词库点名词（含于品名的最长词）；无则规则剥离最长段。
        剥离规则与 clean_name_for_words 一致：多字属性词全局替换，单字单位词只剥结尾。"""
        for w in kw_by_gid.get(g["品组ID"], []):
            if w != broad_hit and w in name:
                return w
        s = str(name)
        s = re.sub(r"[（(【\[][^）)】\]]*[）)】\]]", "", s)
        if broad_hit:
            s = s.replace(broad_hit, "")
        for w in attrs + EXTRA_ATTRS + FEATURES + COLORS + list(STOP):
            s = s.replace(w, "")
        s = re.sub(r"[个只支套包袋盒瓶罐桶根条米卷张片块把副双对组箱台件批枚颗粒节提串捆面顶本]+$", "", s)
        s = re.sub(r"\d+\.?\d*[a-zA-Z%℃度VvWwMmKkGg]*", "", s)
        s = re.sub(r"\d+", "", s)
        s = re.sub(r"[^\u4e00-\u9fff]+", " ", s)
        segs = [seg for seg in s.split() if 2 <= len(seg) <= 6]
        return max(segs, key=len) if segs else ""

    # 品组 -> 树节点（类目, 品名词, 细品名词）
    detail = []
    unmatched = []
    for g in groups:
        b = match_broad(g["品名"])
        if not b:
            # 兜底1：品名提取词作"默认品名词"；兜底2：类目名即品名词（如"垃圾袋"）
            fw = fine_word(g["品名"], "")
            if fw:
                b = fw
                fw = ""
                src = "默认提取"
            elif g["类目"]:
                b = g["类目"]
                fw = ""
                src = "类目兜底"
            else:
                unmatched.append(g)
                continue
        else:
            src = "泛词确认"
            fw = fine_word(g["品名"], b)
        detail.append({
            "三级类目": g["类目"] or "(无类目)",
            "宽品名": b,
            "细品名词": fw or "（同宽品名）",
            "来源": src,
            "品组ID": g["品组ID"],
            "品组名": g["品名"],
            "月销": g["月销"],
            "SKU数": g["SKU数"],
        })

    ddf = pd.DataFrame(detail)
    agg = (ddf.groupby(["三级类目", "宽品名", "细品名词"], dropna=False)
              .agg(品组数=("品组ID", "count"),
                   SKU数=("SKU数", "sum"),
                   月销合计=("月销", "sum"))
              .reset_index())
    # 来源：组内含「泛词确认」即标泛词确认，否则默认提取
    src_map = (ddf.groupby(["三级类目", "宽品名", "细品名词"])["来源"]
                  .apply(lambda s: "泛词确认" if (s == "泛词确认").any()
                         else ("默认提取" if (s == "默认提取").any() else "类目兜底")))
    agg["来源"] = agg.apply(
        lambda r: src_map.get((r["三级类目"], r["宽品名"], r["细品名词"]), ""), axis=1)
    # 示例品名（节点内月销最高品组）
    ex_map = (ddf.sort_values("月销", ascending=False)
                 .drop_duplicates(["三级类目", "宽品名", "细品名词"])
                 .set_index(["三级类目", "宽品名", "细品名词"])["品组名"])
    agg["示例品名"] = agg.apply(
        lambda r: (str(ex_map.get((r["三级类目"], r["宽品名"], r["细品名词"])) or "")[:24]), axis=1)
    agg = agg.sort_values(["三级类目", "宽品名", "月销合计"], ascending=[True, True, False])
    agg.insert(0, "序号", range(1, len(agg) + 1))

    # 未归组品组
    udf = pd.DataFrame(unmatched)
    if len(udf):
        udf = udf.sort_values("月销", ascending=False)
        udf.insert(0, "序号", range(1, len(udf) + 1))

    # 待补宽品名：未归组品组的高频完整词段（覆盖≥3品组），供人工确认补入宽品名
    seg_cnt = defaultdict(list)
    for g in unmatched:
        cleaned = clean_name_for_words(g["品名"], g["品牌"])
        for seg in cleaned.split():
            if 2 <= len(seg) <= 6 and not seg.startswith(("的", "了", "与", "和", "在")):
                seg_cnt[seg].append(g)
    pend = []
    for seg, gs in seg_cnt.items():
        if len(gs) < 3 or seg in effective_broad:
            continue
        sales = sum(g["月销"] for g in gs)
        best = max(gs, key=lambda g: g["月销"])
        pend.append({"候选词": seg, "覆盖品组数": len(gs), "月销合计": round(sales, 0),
                     "主要类目": best["类目"], "示例品名": best["品名"][:26],
                     "确认": ""})
    pdf = pd.DataFrame(pend)
    if len(pdf):
        pdf = pdf.sort_values(["覆盖品组数", "月销合计"], ascending=[False, False])
        pdf.insert(0, "序号", range(1, len(pdf) + 1))

    note = pd.DataFrame([
        {"说明": "树骨架 = 三级类目 × 宽品名 × 细品名词（默认提取）三级结构。"},
        {"说明": "细品名词为自动剥离后的最长词段，仅作默认分组，后续可调整。"},
        {"说明": "「未归组品组」= 品组名未命中任何宽品名，需补充宽品名或归入忽略。"},
        {"说明": "同义词已生效：PD20W→充电线、墙面修补→补墙膏、拉爆→螺丝。"},
    ])

    out = os.path.join(BASE, "output", "自家商品分类树_骨架.xlsx")
    with pd.ExcelWriter(out) as w:
        agg.to_excel(w, sheet_name="树骨架", index=False)
        if len(udf):
            udf.to_excel(w, sheet_name="未归组品组", index=False)
        if len(pdf):
            pdf.to_excel(w, sheet_name="待补宽品名", index=False)
        note.to_excel(w, sheet_name="说明", index=False)

    n_cover = len(set(ddf["品组ID"]))
    print("产出:", out)
    print("  树节点数: %d（类目 %d × 宽品名 %d）" % (len(agg), ddf["三级类目"].nunique(),
                                                  ddf["宽品名"].nunique()))
    print("  已归组品组: %d / %d（覆盖率 %.0f%%）" % (n_cover, len(groups),
                                                  n_cover / len(groups) * 100))
    print("  未归组品组: %d | 待补宽品名候选: %d 条" % (len(unmatched), len(pdf)))
    print("\n待补宽品名 Top15（按覆盖品组数）:")
    if len(pdf):
        print(pdf.head(15)[["候选词", "覆盖品组数", "月销合计", "主要类目"]].to_string(index=False))
    print("\n节点 Top20（按月销）:")
    cols = ["三级类目", "宽品名", "细品名词", "品组数", "月销合计"]
    print(agg.sort_values("月销合计", ascending=False).head(20)[cols].to_string(index=False))
    if len(unmatched):
        print("\n未归组样例（月销前5）:")
        print(udf.head(5)[["品名", "类目", "月销"]].to_string(index=False))


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == "--confirm":
        confirm_and_fix()
    elif len(sys.argv) > 1 and sys.argv[1] == "--tree":
        build_tree_preview()
    else:
        main()

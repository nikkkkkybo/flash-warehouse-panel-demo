# -*- coding: utf-8 -*-
"""
品级语义匹配器（第3步：竞品逐级下钻，跨品牌品对应）

基于「自家商品分类树」（config/category_tree.yaml + 自家品组）：
  竞品名 → 类目 → 宽品名 → 细品名词 逐级下钻，停在最深层；
  品牌剔除匹配键（仅作对比列）→ 跨品牌同品可对应（南孚五号 vs 杂牌五号）；
  规格词（1.5V/5W/尺寸）不参与匹配（品级语义，规格自行查阅）。

置信度分级（强弱对应）：
  类目✓+宽品名✓+细品✓  = 0.80 强
  类目✓+宽品名✓         = 0.70
  类目✓+细品✓（宽品名未命中）= 0.60
  类目✗+宽品名✓+细品✓  = 0.55 弱（类目缺失，如骏涵店）
  类目✗+宽品名✓（仅同宽品名节点）= 0.50 弱（人工复核）
  其余                   = 无候选

类目硬约束（防跨类目脏配）：
  竞品类目缺失时，纯宽品名命中不再跨类目全量收集候选（否则"口罩"会配到
  袜子/鞋等类目下的撞词品组）；必须细品名词也命中（双命中）才给弱候选，
  或仅限自家以该宽品名直接入库的（同宽品名）占位节点。
  "(无类目)"不作为可命中类目，竞品类目为空一律按类目缺失处理。
"""
import os
import re
from collections import defaultdict

import pandas as pd
import yaml

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TREE_CFG = os.path.join(BASE, "config", "category_tree.yaml")
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")
KW_FILE = os.path.join(BASE, "output", "自家关键词库.xlsx")

# 常见螺丝/五金属性词（与 build_category_tree 一致，避免"十字沉头"当细品名）
EXTRA_ATTRS = ["十字", "沉头", "圆头", "扁头", "六角", "内六角", "外六角", "一字",
               "米字", "钻尾", "自攻", "镀黑锌", "镀锌", "加长", "加硬", "细牙",
               "粗牙", "蘑菇头", "大扁头", "平头", "尖头", "梅花", "内梅花"]
FEATURES = ["防水", "防滑", "加厚", "便携", "折叠", "静音", "耐磨", "耐高温",
            "绝缘", "阻燃", "防爆", "可调", "通用", "家用", "商用", "专业",
            "工业级", "环保", "抗压", "耐摔", "强力", "加粗", "加长", "加大",
            "迷你", "超大", "超强", "加硬", "防漏", "密封", "省力", "自动",
            "国标", "优选", "免打孔", "免钉", "即热", "简易", "无线", "有线",
            "电动", "手动", "新款", "升级", "加强", "特价", "促销"]
COLORS = ["白色", "黑色", "红色", "蓝色", "绿色", "黄色", "透明", "金色", "银色",
          "灰色", "彩色", "粉色", "紫色", "橙色", "棕色", "米色", "深色", "浅色",
          "白光", "暖白", "暖光", "中性光", "冷白", "自然光", "冷光"]
UNITS = ["个", "只", "支", "套", "包", "袋", "盒", "瓶", "罐", "桶", "根", "条",
         "米", "卷", "张", "片", "块", "把", "副", "双", "对", "组", "箱", "台",
         "件", "批", "枚", "颗", "粒", "节", "提", "串", "捆", "面", "顶", "本"]
STOP = {"正品", "包邮", "家用", "通用", "专用", "新款", "批发", "加厚", "高配",
        "升级", "加强", "精品", "促销", "特价", "优惠", "赠品", "免费", "上门",
        "安装", "维修", "厂家", "直销", "现货", "下单", "客服", "联系", "咨询",
        "备注", "型号", "规格", "颜色", "款式", "左右", "上下", "系列", "套装",
        "闪购", "美团", "闪电仓", "热销", "爆款", "断货", "补货", "到货",
        "可选", "多规格", "随机", "适用", "适配", "兼容", "支持", "含",
        "带", "送", "新", "好", "大", "小", "型"}


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    """数字解析：委托给全站唯一口径 `sales_text.num`。

    本模块自己不用它，是被 build_comparison / build_product_analysis / ai_top50 /
    align 当公共工具导入的。以前 top50.py、insight_scan.py 各有一份近似实现
    （一份不认「万/k」单位），同一列"销量文本"就会出现两套数——收敛到这里。
    「成交 / 意向」的语义区分不在数值层做，用 `sales_text.is_intent`。
    """
    import sales_text
    return sales_text.num(v)


def clean_name(name, brand=""):
    """竞品/品名清洗：去括号/品牌/属性词/规格数字，用于词段提取与匹配。
    材质词保留（"塑料/尼龙"是品名词根）。"""
    s = str(name)
    s = re.sub(r"[（(【\[][^）)】\]]*[）)】\]]", "", s)
    if brand:
        s = s.replace(str(brand), "")
    for w in FEATURES + COLORS + list(STOP):
        s = s.replace(w, "")
    s = re.sub(r"[个只支套包袋盒瓶罐桶根条米卷张片块把副双对组箱台件批枚颗粒节提串捆面顶本]+$", "", s)
    s = re.sub(r"\d+\.?\d*[a-zA-Z%℃度VvWwMmKkGg]*", "", s)
    s = re.sub(r"\d+", "", s)
    s = re.sub(r"[^\u4e00-\u9fff]+", " ", s)
    return s


def load_spu_groups():
    """自家表按 SPUID 分组：品组名=月销最高SKU名，月销=组内合计。"""
    own = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    groups = []
    for spuid, g in own.groupby("SPUID"):
        spuid = _clean(spuid)
        if not spuid:
            continue
        g2 = g.copy()
        g2["__s"] = g2["月销量"].apply(_num)
        rep = g2.loc[g2["__s"].idxmax()] if len(g2) else g2.iloc[0]
        groups.append({
            "品组ID": spuid,
            "品名": _clean(rep.get("商品名称")) or _clean(rep.get("清洗后商品名")),
            "类目": _clean(rep.get("类目三级_归一")) or _clean(rep.get("美团闪购类目三级")),
            "品牌": _clean(rep.get("品牌（抽取）")) or _clean(rep.get("品牌名称")),
            "月销": float(g2["__s"].sum()),
            "SKU数": len(g2),
        })
    return groups


class SpuSemanticMatcher:
    """树索引 + 竞品逐级下钻匹配。"""

    def __init__(self):
        self.broad = []            # 宽品名列表
        self.syn = {}              # 竞品叫法 -> 标准品名
        self.attrs = []            # 属性词（含 EXTRA_ATTRS）
        self.groups = []           # 全部自家品组
        self.gmap = {}             # 品组ID -> 品组 dict
        # 节点索引：类目 -> {宽品名 -> {细品名词 -> [品组ID...]}}
        self.cat_tree = defaultdict(lambda: defaultdict(lambda: defaultdict(list)))
        self.cats = set()          # 全部类目（自家）

    def build(self, own_df_path=None):
        cfg = yaml.safe_load(open(TREE_CFG, encoding="utf-8"))
        self.broad = list(cfg.get("broad_words", []))
        self.attrs = list(cfg.get("attribute_words", [])) + EXTRA_ATTRS
        self.syn = cfg.get("word_synonyms", {})
        eff_broad = set(self.broad) | set(self.syn.values())

        # 关键词库点名词 -> 品组（细品名词优先来源）
        kw_by_gid = defaultdict(list)
        if os.path.exists(KW_FILE):
            kw = pd.read_excel(KW_FILE, dtype=str).fillna("")
            kw = kw[kw["状态"].astype(str).isin(["", "启用"])]
            for _, r in kw.iterrows():
                w = _clean(r.get("关键词"))
                gid = _clean(r.get("品组ID"))
                if w and gid and len(w) >= 2:
                    kw_by_gid[gid].append(w)
        for gid in kw_by_gid:
            kw_by_gid[gid] = sorted(set(kw_by_gid[gid]), key=len, reverse=True)

        def match_broad(name):
            best, best_len = "", 0
            for w in eff_broad:
                if w in name and len(w) > best_len:
                    best, best_len = w, len(w)
            for k, v in self.syn.items():
                if k in name and len(v) > best_len:
                    best, best_len = v, len(v)
            return best

        def fine_word(name, broad_hit):
            for w in kw_by_gid.get(g["品组ID"], []):
                if w != broad_hit and w in name:
                    return w
            s = re.sub(r"[（(【\[][^）)】\]]*[）)】\]]", "", str(name))
            if broad_hit:
                s = s.replace(broad_hit, "")
            for w in self.attrs + FEATURES + COLORS + list(STOP):
                s = s.replace(w, "")
            s = re.sub(r"[个只支套包袋盒瓶罐桶根条米卷张片块把副双对组箱台件批枚颗粒节提串捆面顶本]+$", "", s)
            s = re.sub(r"\d+\.?\d*[a-zA-Z%℃度VvWwMmKkGg]*", "", s)
            s = re.sub(r"\d+", "", s)
            s = re.sub(r"[^\u4e00-\u9fff]+", " ", s)
            segs = [seg for seg in s.split() if 2 <= len(seg) <= 6]
            return max(segs, key=len) if segs else ""

        self.groups = load_spu_groups() if own_df_path is None else load_spu_groups()
        for g in self.groups:
            self.gmap[g["品组ID"]] = g
            # (无类目) 不作为可命中的有效类目：竞品类目缺失时一律走"类目缺失"弱对应分支
            if g["类目"]:
                self.cats.add(g["类目"])
            b = match_broad(g["品名"])
            if not b:
                fw = fine_word(g["品名"], "")
                if fw:
                    b, fw = fw, ""
                elif g["类目"]:
                    b, fw = g["类目"], ""
                else:
                    continue
            else:
                fw = fine_word(g["品名"], b)
            cat = g["类目"] or "(无类目)"
            self.cat_tree[cat][b][fw or "（同宽品名）"].append(g["品组ID"])
        return self

    # ------------------------------------------------------------------
    def match(self, comp_name, comp_cat, comp_brand=""):
        """竞品逐级下钻。返回 (候选品组ID列表top3, 置信度, 层级描述) 或 None。
        comp_cat: 竞品归一化后的三级类目（与自家口径一致）；空串=类目缺失。"""
        name = str(comp_name or "")
        if not name:
            return None
        # 宽品名命中（在含品牌原文上匹配；同义词归标准品名）
        broad_hit, best_len = "", 0
        eff_broad = set(self.broad) | set(self.syn.values())
        for w in eff_broad:
            if w in name and len(w) > best_len:
                broad_hit, best_len = w, len(w)
        for k, v in self.syn.items():
            if k in name and len(v) > best_len:
                broad_hit, best_len = v, len(v)

        cat = _clean(comp_cat) or "(无类目)"
        cat_ok = cat in self.cats

        # 细品名词候选（该宽品名节点下的细品名是否含于竞品名）
        fine_hit = ""
        if broad_hit:
            node_fines = []
            if cat_ok:
                node_fines = list(self.cat_tree.get(cat, {}).get(broad_hit, {}).keys())
            else:
                # 类目缺失：跨类目收集该宽品名下所有细品名
                for c, bt in self.cat_tree.items():
                    node_fines.extend(bt.get(broad_hit, {}).keys())
            node_fines = [f for f in node_fines if f and f != "（同宽品名）"]
            for f in sorted(node_fines, key=len, reverse=True):
                if f in name:
                    fine_hit = f
                    break

        cand_gids = []
        if cat_ok and broad_hit:
            node = self.cat_tree[cat][broad_hit]
            if fine_hit:
                cand_gids = list(node.get(fine_hit, []))
                conf, layer = 0.80, "类+宽品名+细品"
            else:
                for f, gids in node.items():
                    cand_gids.extend(gids)
                conf, layer = 0.70, "类+宽品名"
        elif cat_ok and not broad_hit and fine_hit:
            # 类目命中+细品名词命中但宽品名未收录（词库细词直达）
            for f, gids in self.cat_tree[cat].items():
                if f == fine_hit:
                    cand_gids = list(gids)
                    break
            conf, layer = 0.60, "类+细品"
        elif not cat_ok and broad_hit:
            # 类目缺失弱对应（硬约束：防跨类目脏配，如"口罩"配到袜子/鞋品组）
            if fine_hit:
                # 双命中（宽品名+细品名）：跨类目收集精确节点，置信度 0.55
                for c, bt in self.cat_tree.items():
                    gids = bt.get(broad_hit, {}).get(fine_hit, [])
                    if gids:
                        cand_gids.extend(gids)
                        break
                conf, layer = 0.55, "宽品名+细品(类目缺失)"
            else:
                # 单宽品名命中：仅收集（同宽品名）占位节点（自家品组本身以该宽品名入库），
                # 不再跨类目全量收集该宽品名下所有品组；置信度 0.50 排序靠后供人工复核
                for c, bt in self.cat_tree.items():
                    cand_gids.extend(bt.get(broad_hit, {}).get("（同宽品名）", []))
                conf, layer = 0.50, "宽品名(类目缺失)"
        else:
            return None

        if not cand_gids:
            return None
        # 去重 + 月销降序 top3（代表=月销最高）
        seen, uniq = set(), []
        for gid in cand_gids:
            if gid not in seen:
                seen.add(gid)
                uniq.append(gid)
        uniq.sort(key=lambda gid: self.gmap[gid]["月销"], reverse=True)
        top3 = uniq[:3]
        return top3, conf, layer

# -*- coding: utf-8 -*-
"""
AI 经营分析（方案A）：把本地分析数据提炼成摘要 -> 调 DeepSeek 生成经营建议报告。

职责：
  1. build_summary(shops, directions)  数据摘要构建（控制篇幅，约3000-5000字）
  2. call_deepseek(summary, directions, config)  DeepSeek 调用（OpenAI 兼容接口）
  3. generate_report(...) 全流程：摘要 -> 调用 -> 保存 output/AI经营分析报告.md

配置：config/ai_config.yaml（api_key 由用户自填，只存本地；base_url/model 可改）
"""
import os
import sys
import time
import glob
from datetime import datetime

import pandas as pd

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
CONFIG_PATH = os.path.join(BASE, "config", "ai_config.yaml")
REPORT_PATH = os.path.join(OUTPUT, "AI经营分析报告.md")
REPORT_HTML_PATH = os.path.join(OUTPUT, "AI经营分析报告.html")
DRAFT_DIR = os.path.join(OUTPUT, "_ai_reports")   # 草稿暂存目录（保留期后自动清理）
DRAFT_KEEP_DAYS = 7                                # 草稿保留天数（暂存一段时间，期间可回看/导出）

DEFAULT_CONFIG = {
    "api_key": "",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-chat",
}

# 方向 -> 中文名（与面板勾选一致）
DIRECTIONS = {
    "hot": "热销品与引进建议",
    "price": "定价与价格策略",
    "display": "商品展示与规格差距",
    "insight": "综合经营洞察",
}

SYSTEM_PROMPT = """你是一位深耕闪电仓（前置仓）与五金建材零售的资深经营顾问，长期帮助小店老板从竞品数据里找出能直接借鉴、落地执行的办法。
你的输出要求：
1. 全程用大白话，像跟开店老板聊天，不用空话套话；
2. 每条建议必须具体可执行（给到具体商品/价格数字/做法），不要泛泛而谈；
3. 用 Markdown 格式组织，按用户要求的板块输出，每板块用小标题；
4. 数据里没有的信息不要编造，宁可说"这份数据里看不出"；
5. 结合闪电仓经营特点：热销品引进、定价毛利、标题规格优化、品类结构、防守自家热销品。"""


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _read(path, **kw):
    """安全读表：不存在/解析失败返回空 DataFrame。"""
    if not os.path.exists(path):
        return pd.DataFrame()
    try:
        return pd.read_excel(path, **kw)
    except Exception:
        return pd.DataFrame()


def _to_num(v):
    try:
        f = float(str(v).replace(",", "").replace("¥", "").strip())
        return f if f == f else None
    except (ValueError, TypeError):
        return None


def build_name_map(shops):
    """商品名 -> (店铺, Sku_id) 映射（供报告内竞品名转详情链接）。"""
    name_map = {}
    for s in shops:
        try:
            comp_f = os.path.join(OUTPUT, f"试跑_{s}", "竞品_归一化后.xlsx")
            if not os.path.exists(comp_f):
                continue
            cdf = pd.read_excel(comp_f, dtype=str)
            name_col = None
            for c in ("name", "商品名字", "清洗后商品名"):
                if c in cdf.columns:
                    name_col = c
                    break
            sku_col = None
            for c in ("sku_id", "Sku_id"):
                if c in cdf.columns:
                    sku_col = c
                    break
            if not name_col or not sku_col:
                continue
            for _, r in cdf.iterrows():
                nm = _clean(r.get(name_col))
                sku = _clean(r.get(sku_col))
                if len(nm) >= 6 and sku:
                    name_map.setdefault(nm, (s, sku))
        except Exception:
            continue
    return name_map


def linkify_report(md_text, name_map):
    """把报告中提到的真实竞品名替换为详情页链接（长名优先，避免子串误替换）。"""
    if not md_text or not name_map:
        return md_text
    from urllib.parse import quote
    text = str(md_text)
    # 长名优先
    for nm, (shop, sku) in sorted(name_map.items(), key=lambda kv: -len(kv[0])):
        if nm not in text:
            continue
        url = f"/comp_detail?shop={quote(shop)}&sku={quote(sku)}"
        text = text.replace(nm, f"[{nm}]({url})")
    return text


def load_config(path=None):
    """读取 AI 配置；文件缺失时返回默认配置。
    环境变量 DEEPSEEK_API_KEY 优先于配置文件（方案A：密钥不落盘）。"""
    path = path or CONFIG_PATH
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        try:
            import yaml
            with open(path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            for k in cfg:
                if data.get(k):
                    cfg[k] = str(data[k]).strip()
        except Exception:
            pass
    _env_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if _env_key:
        cfg["api_key"] = _env_key
    return cfg


def save_config(api_key, path=None):
    """保存 api_key 到本地配置文件。
    读原始 yaml 全量结构后仅改 api_key，保留其他段
    （driver/prescreen/embeddings 等，防止回写丢配置）。"""
    path = path or CONFIG_PATH
    import yaml
    data = {}
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
        except Exception:
            data = {}
    if not isinstance(data, dict):
        data = {}
    # 顶层标量键兜底（文件缺失/损坏时保证最小结构）
    for k, v in DEFAULT_CONFIG.items():
        data.setdefault(k, v)
    data["api_key"] = (api_key or "").strip()
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)
    return load_config(path)


def _shop_price_cols(df, shops, suffix="_到手价"):
    """从价格矩阵中挑出所选店的到手价/月售列。"""
    cols = {}
    for c in df.columns:
        for s in shops:
            if c.startswith(s + suffix):
                cols.setdefault(s, {})[suffix] = c
    return cols


def build_summary(shops, directions=None):
    """构建数据摘要文本（按所选方向切数据）。返回 (text, stats)。"""
    directions = directions or list(DIRECTIONS.keys())
    parts = []
    stats = {"shops": shops, "sections": []}

    # ---- 店铺概况 ----
    summary_txt = ["店铺概况："]
    total_rows = 0
    lib = _read(os.path.join(OUTPUT, "对齐汇总库.xlsx"), sheet_name="SKU匹配层", dtype=str)
    for s in shops:
        n = 0
        cat_txt = ""
        try:
            comp_f = os.path.join(OUTPUT, f"试跑_{s}", "竞品_归一化后.xlsx")
            if os.path.exists(comp_f):
                cdf = pd.read_excel(comp_f, dtype=str)
                n = len(cdf)
                cat_col = "类目三级_归一" if "类目三级_归一" in cdf.columns else "category3"
                if cat_col in cdf.columns:
                    vc = cdf[cat_col].fillna("(未分类)").value_counts().head(5)
                    cat_txt = "；类目Top5: " + "、".join(
                        f"{k}({v})" for k, v in vc.items())
        except Exception:
            pass
        total_rows += n
        aligned_n = int((lib["竞店名称"] == s).sum()) if len(lib) else 0
        summary_txt.append(f"- {s}：竞品 {n} 条，已对应 {aligned_n} 条{cat_txt}")
    summary_txt.append(f"合计竞品 {total_rows} 条。")
    parts.append("\n".join(summary_txt))
    stats["sections"].append("概况")

    # ---- 热销品与引进建议 ----
    if "hot" in directions:
        hot = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="热销Top100", dtype=str)
        if len(hot):
            hot = hot.head(20)
            lines = ["【热销品Top20】商品名/月售合计/在售店数/价格区间："]
            stats["hot"] = []
            for _, r in hot.head(10).iterrows():
                lines.append(
                    f"- {_clean(r.get('商品名'))} | 月售{_clean(r.get('月售合计'))} | "
                    f"{_clean(r.get('在售店数'))}店在售 | 价{_clean(r.get('最低价'))}~{_clean(r.get('最高价'))}"
                    f"（类目：{_clean(r.get('类目'))}）")
                stats["hot"].append({
                    "商品名": _clean(r.get("商品名")), "月售": _clean(r.get("月售合计")),
                    "在售店数": _clean(r.get("在售店数")), "最低价": _clean(r.get("最低价")),
                    "最高价": _clean(r.get("最高价")), "类目": _clean(r.get("类目")),
                })
            parts.append("\n".join(lines))
            stats["sections"].append("热销")
        gap = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="缺口清单", dtype=str)
        if len(gap):
            gap = gap.head(20)
            lines = ["【自家缺口品Top20】竞品在卖、自家没有的："]
            stats["gap"] = []
            for _, r in gap.head(10).iterrows():
                lines.append(f"- {_clean(r.get('商品名'))} | 月售{_clean(r.get('月售'))} | {_clean(r.get('在售店数'))}店在售")
                stats["gap"].append({
                    "商品名": _clean(r.get("商品名")), "月售": _clean(r.get("月售")),
                    "在售店数": _clean(r.get("在售店数")),
                })
            parts.append("\n".join(lines))
            stats["sections"].append("缺口")
        cat_gap = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="类目缺口", dtype=str)
        if len(cat_gap):
            lines = ["【类目缺口】竞品有、自家少/无的类目："]
            stats["cat_gap"] = []
            for _, r in cat_gap.head(15).iterrows():
                lines.append(f"- {_clean(r.get('类目'))}：竞品 {_clean(r.get('竞店商品数'))} 个，自家 {_clean(r.get('自家商品数'))} 个")
                stats["cat_gap"].append({
                    "类目": _clean(r.get("类目")), "竞店商品数": _clean(r.get("竞店商品数")),
                    "自家商品数": _clean(r.get("自家商品数")),
                })
            parts.append("\n".join(lines))
            stats["sections"].append("类目缺口")

    # ---- 定价与价格策略 ----
    if "price" in directions:
        band = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="价格带", dtype=str)
        if len(band):
            lines = ["【价格带】主要类目最低/中位/最高价："]
            stats["band"] = []
            for _, r in band.head(8).iterrows():
                lines.append(
                    f"- {_clean(r.get('类目'))}：{_clean(r.get('商品数'))}个商品，"
                    f"价 {_clean(r.get('最低价'))} / 中位 {_clean(r.get('价格中位数'))} / {_clean(r.get('最高价'))}")
                stats["band"].append({
                    "类目": _clean(r.get("类目")), "商品数": _clean(r.get("商品数")),
                    "最低价": _clean(r.get("最低价")), "中位价": _clean(r.get("价格中位数")),
                    "最高价": _clean(r.get("最高价")),
                })
            parts.append("\n".join(lines))
            stats["sections"].append("价格带")
        matrix = _read(os.path.join(OUTPUT, "竞品价格矩阵.xlsx"), dtype=str)
        own = _read(os.path.join(OUTPUT, "自家_归一化后.xlsx"), dtype=str)
        if len(matrix) and len(own):
            own_price = {}
            sku_col = "SKUID" if "SKUID" in own.columns else "自家单品编号"
            price_col = "美团闪购渠道售价" if "美团闪购渠道售价" in own.columns else "售价"
            if sku_col in own.columns and price_col in own.columns:
                for _, r in own.iterrows():
                    sku = _clean(r.get(sku_col))
                    if sku:
                        own_price[sku] = _clean(r.get(price_col))
            lines = ["【自家同款比价】自家品 vs 所选店竞品到手价（价差率=竞品价-自家价)/自家价）："]
            stats["compare"] = []
            n = 0
            for _, r in matrix.iterrows():
                if n >= 20:
                    break
                own_sku = _clean(r.get("自家单品编号"))
                own_name = _clean(r.get("自家商品名"))
                oprice = own_price.get(own_sku, "")
                cprices = [_clean(r.get(c)) for c in r.index if any(c.startswith(s) and c.endswith("_到手价") for s in shops)]
                cprice = next((p for p in cprices if p), "")
                if not own_name or not cprice:
                    continue
                diff = ""
                of, cf_ = _to_num(oprice), _to_num(cprice)
                if of and cf_ and of > 0:
                    diff = f"（价差{(cf_ - of) / of * 100:+.0f}%）"
                lines.append(f"- {own_name} | 自家价 {oprice} vs 竞品 {cprice} {diff}")
                stats["compare"].append({
                    "自家商品名": own_name, "自家价": oprice, "竞品价": cprice,
                    "价差": diff.replace("（", "").replace("）", ""),
                })
                n += 1
            if n:
                parts.append("\n".join(lines))
                stats["sections"].append("比价")

    # ---- 商品展示与规格差距 ----
    if "display" in directions:
        hot = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="热销Top100", dtype=str)
        if len(hot):
            lines = ["【竞品热销标题写法】供学习标题/卖点表述："]
            for _, r in hot.head(15).iterrows():
                lines.append(f"- {_clean(r.get('商品名'))}")
            parts.append("\n".join(lines))
            stats["sections"].append("标题")
        own = _read(os.path.join(OUTPUT, "自家_归一化后.xlsx"), dtype=str)
        if len(own):
            has_img = "是否有图片" if "是否有图片" in own.columns else None
            if has_img:
                # P4.3：渠道重复行按 SKUID 去重（图率按唯一 SKU 计）
                own_u = own.drop_duplicates(subset=["SKUID"], keep="first") if "SKUID" in own.columns else own
                n_img = int((own_u[has_img].fillna("").astype(str).str.strip() != "").sum())
                parts.append(f"【自家图片】有图 {n_img}/{len(own_u)}（{n_img / max(len(own_u), 1) * 100:.0f}%），缺图会影响顾客点击。")
                stats["sections"].append("图片")

    # ---- 综合经营洞察 ----
    if "insight" in directions:
        ov = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="总览", dtype=str)
        if len(ov):
            lines = ["【本次分析总览指标】："]
            for _, r in ov.iterrows():
                lines.append(f"- {_clean(r.get('指标'))}：{_clean(r.get('数值'))}")
            parts.append("\n".join(lines))
            stats["sections"].append("总览")
        cat = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="品类结构", dtype=str)
        if len(cat):
            lines = ["【竞品品类结构】各店卖得最多的类目："]
            for _, r in cat.head(12).iterrows():
                lines.append(f"- {_clean(r.get('竞店名称'))} | {_clean(r.get('类目'))}：{_clean(r.get('商品数'))}个，月售{_clean(r.get('月售合计'))}")
            parts.append("\n".join(lines))
            stats["sections"].append("品类")
        defend = _read(os.path.join(OUTPUT, "竞争分析报告.xlsx"), sheet_name="防守清单", dtype=str)
        if len(defend):
            lines = ["【防守清单】自家热销品需盯防（对手低价/缺货）："]
            for _, r in defend.head(15).iterrows():
                lines.append(f"- {_clean(r.get('商品名称'))} | 自家月售{_clean(r.get('月销量'))} | 售价{_clean(r.get('售价'))}")
            parts.append("\n".join(lines))
            stats["sections"].append("防守")

    missing = [k for k in directions if k not in stats["sections"]]
    for k in missing:
        parts.append(f"【{DIRECTIONS.get(k, k)}】相关数据未生成（请先在选店分析页运行分析，或所选店铺暂无此类数据）。")
    return "\n\n".join(parts), stats


def call_deepseek(summary_text, directions, config, timeout=120):
    """调 DeepSeek 生成报告。返回报告 Markdown 文本。（已收口 ai_client）"""
    dir_names = [DIRECTIONS[d] for d in directions if d in DIRECTIONS]
    user_text = (
        f"请基于以下本地竞品分析数据，输出经营建议报告（板块：{'、'.join(dir_names)}）。\n"
        f"数据如下（来源：本地对齐与竞争分析结果）：\n\n{summary_text}"
    )
    from ai_client import call_llm
    return call_llm("analysis",
                    [{"role": "system", "content": SYSTEM_PROMPT},
                     {"role": "user", "content": user_text}],
                    config=config, timeout=timeout,
                    temperature=0.7, max_tokens=3000)


def md_to_html(md_text):
    """Markdown 轻量转 HTML（标题/加粗/链接/列表/分隔线/段落），供报告页面与面板展示。"""
    if not md_text:
        return ""
    import html as _html
    import re as _re
    out = []
    for line in str(md_text).split("\n"):
        t = _html.escape(line)
        # 链接 [text](url) -> <a>
        t = _re.sub(r"\[([^\]]+)\]\(([^)]+)\)",
                    r'<a href="\2" target="_blank" class="rep-link">\1</a>', t)
        # 加粗
        parts = t.split("**")
        if len(parts) > 1:
            t = "".join(f"<b>{p}</b>" if i % 2 == 1 else p for i, p in enumerate(parts))
        if t.startswith("### "):
            out.append(f'<h4 class="sec3">{t[4:]}</h4>')
        elif t.startswith("## "):
            out.append(f'<h3 class="sec2">{t[3:]}</h3>')
        elif t.startswith("# "):
            out.append(f'<h2 class="sec1">{t[2:]}</h2>')
        elif t.startswith("- "):
            out.append(f'<div class="li">· {t[2:]}</div>')
        elif t.startswith("1. ") or t.startswith("2. ") or t.startswith("3. "):
            out.append(f'<div class="li num">{t}</div>')
        elif t.strip() in ("---", "***"):
            out.append('<hr>')
        elif t.strip() == "":
            out.append('<div class="gap"></div>')
        else:
            out.append(f'<div class="p">{t}</div>')
    return "\n".join(out)


def build_html_report(report_md, shops, directions, generated_at=None):
    """把报告包成完整 HTML 页面（内联样式，可直接打开/打印）。"""
    from datetime import datetime
    generated_at = generated_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    dir_names = "、".join(DIRECTIONS[d] for d in directions if d in DIRECTIONS)
    body = md_to_html(report_md)
    return f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<title>AI 经营分析报告</title>
<style>
body{{font-family:'Microsoft YaHei',sans-serif;background:#f1f5f9;margin:0;padding:24px;color:#1e293b;}}
.wrap{{max-width:960px;margin:0 auto;background:#fff;border-radius:12px;box-shadow:0 2px 12px rgba(0,0,0,.08);padding:32px 40px;}}
h1{{font-size:22px;margin:0 0 6px;color:#0f172a;}}
.meta{{color:#64748b;font-size:13px;margin-bottom:20px;border-bottom:2px solid #2563eb;padding-bottom:14px;}}
.meta b{{color:#1e40af;}}
h2.sec1{{font-size:18px;color:#1e3a8a;margin:22px 0 10px;border-left:4px solid #2563eb;padding-left:10px;}}
h3.sec2{{font-size:16px;color:#1e40af;margin:18px 0 8px;}}
h4.sec3{{font-size:14px;color:#475569;margin:14px 0 6px;}}
.p{{margin:6px 0;line-height:1.9;font-size:14px;}}
.li{{margin:4px 0 4px 10px;line-height:1.8;font-size:14px;}}
.li.num{{color:#334155;}}
.gap{{height:6px;}}
hr{{border:none;border-top:1px solid #e2e8f0;margin:14px 0;}}
.foot{{margin-top:28px;padding-top:12px;border-top:1px solid #e2e8f0;color:#94a3b8;font-size:12px;}}
b{{color:#0f172a;}}
a.rep-link{{color:#2563eb;text-decoration:none;border-bottom:1px dashed #93c5fd;}}
a.rep-link:hover{{color:#1d4ed8;border-bottom-style:solid;}}
</style>
</head>
<body>
<div class="wrap">
<h1>🤖 AI 经营分析报告</h1>
<div class="meta">店铺：<b>{_html_escape("、".join(shops))}</b>　板块：{_html_escape(dir_names)}　生成时间：{generated_at}<br>数据来源：本地竞品对齐与竞争分析结果（人工确认进度与汇总库）</div>
{body}
<div class="foot">本报告由 DeepSeek 大模型基于本地数据摘要生成，供经营参考；关键决策请结合人工确认数据复核。</div>
</div>
</body>
</html>"""


def _html_escape(s):
    import html as _h
    return _h.escape(str(s))


def _cleanup_drafts():
    """清理超过保留期（DRAFT_KEEP_DAYS）的草稿文件，防止堆积。"""
    cutoff = time.time() - DRAFT_KEEP_DAYS * 86400
    if os.path.isdir(DRAFT_DIR):
        for f in os.listdir(DRAFT_DIR):
            p = os.path.join(DRAFT_DIR, f)
            try:
                if os.path.getmtime(p) < cutoff:
                    os.remove(p)
            except OSError:
                pass


def _latest_draft():
    """返回最新一份草稿的 (md_path, html_path)，无草稿返回 (None, None)。"""
    if not os.path.isdir(DRAFT_DIR):
        return None, None
    mds = sorted(glob.glob(os.path.join(DRAFT_DIR, "ai_*.md")))
    if not mds:
        return None, None
    md = mds[-1]
    return md, md.replace(".md", ".html")


def generate_report(shops, directions, config_path=None, out_path=None, mock=False, export=False):
    """全流程：摘要 -> 调用 -> 保存。
    export=False（默认）：写入草稿暂存目录 output/_ai_reports/（带时间戳，保留 7 天自动清理），
    供面板展示与后续导出；export=True：写正式文件（AI经营分析报告.md/.html）。
    mock=True 时用本地假响应（测试用，不消耗 API）。"""
    config_path = config_path or CONFIG_PATH
    export = bool(export)
    config = load_config(config_path)
    summary_text, stats = build_summary(shops, directions)

    if mock:
        report = (
            "# AI 经营分析（模拟数据，未调用真实接口）\n\n"
            "## 热销品与引进建议\n- 模拟：请配置真实 API key 后重新生成。\n\n"
            "## 定价与价格策略\n- 模拟：价格带数据显示……\n"
        )
    else:
        report = call_deepseek(summary_text, directions, config)

    # 报告中提到的真实竞品名 -> 详情页链接（建议可点开看竞品详情）
    name_map = build_name_map(shops)
    report = linkify_report(report, name_map)
    # 结构化表格条目补 竞品(shop,sku) 链接信息（前端表格名称可点开详情）
    for key in ("hot", "gap"):
        for item in stats.get(key, []):
            m = name_map.get(item.get("商品名"))
            if m:
                item["shop"], item["sku"] = m

    _cleanup_drafts()
    if export:
        os.makedirs(os.path.dirname(out_path or REPORT_PATH), exist_ok=True)
        out_path = out_path or REPORT_PATH
    else:
        os.makedirs(DRAFT_DIR, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_path = os.path.join(DRAFT_DIR, f"ai_{ts}.md")
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(report)
    # HTML 版报告（同内容，样式化页面）
    html_path = out_path.replace(".md", ".html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(build_html_report(report, shops, directions))
    return report, stats, out_path, html_path


def export_draft_to_report(server_prefix="http://127.0.0.1:5000"):
    """把最新一份草稿导出为正式报告文件（AI经营分析报告.md/.html）。
    HTML 内的相对链接（/comp_detail?...）补全为服务器前缀，保证文件单独打开也可跳转。
    无草稿时返回 (None, None)。"""
    draft_md, draft_html = _latest_draft()
    if not draft_md:
        return None, None
    with open(draft_md, "r", encoding="utf-8") as f:
        md = f.read()
    with open(draft_html, "r", encoding="utf-8") as f:
        html = f.read()
    # 补全相对链接
    html = html.replace('href="/comp_detail', f'href="{server_prefix}/comp_detail')
    html = html.replace('href=\'/comp_detail', f"href='{server_prefix}/comp_detail")
    with open(REPORT_PATH, "w", encoding="utf-8") as f:
        f.write(md)
    with open(REPORT_HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    return REPORT_PATH, REPORT_HTML_PATH


# =====================================================================
# 品类诊断卡片墙
# =====================================================================

CATEGORY_DIAG_SYSTEM_PROMPT = """你是一位深耕闪电仓（前置仓）与五金建材零售的资深经营顾问。
输入：品类维度的结构化数据（JSON 数组）。
输出要求：
1. 为每个品类输出一段诊断（50-80字，大白话，像跟老板聊天，用具体数字说话）；
2. 为每个品类输出 1-3 条具体可执行建议（每条含具体商品名/价格/做法，30-60字）；
3. 严格按 JSON 格式输出，不要 Markdown，不要多余文字；
4. 只基于提供的数据，不编造；数据少的品类直说"数据较少，建议补充后重新分析"。"""


def compute_health_score(card):
    """品类健康分（0-100）。
    权重：销售力30% / 品类深度25% / 价格竞争力25% / 缺口20% / 对齐加分0-10。"""
    score = 0.0

    # 1. 销售力得分（0-30）
    sales = card.get("comp_sales_total", 0)
    if sales >= 1000:
        score += 30
    elif sales >= 500:
        score += 25
    elif sales >= 200:
        score += 20
    elif sales >= 50:
        score += 15
    elif sales > 0:
        score += 10

    # 2. 品类深度得分（0-25）——竞品品数多 = 市场大
    count = card.get("comp_product_count", 0)
    if count >= 100:
        score += 25
    elif count >= 50:
        score += 20
    elif count >= 20:
        score += 15
    elif count >= 10:
        score += 10
    else:
        score += 5

    # 3. 价格竞争力得分（0-25）
    pdiff = card.get("price_diff_pct")
    if pdiff is not None:
        if pdiff <= -10:
            score += 25    # 自家便宜 10%+ → 强势
        elif pdiff <= 0:
            score += 22    # 自家更便宜
        elif pdiff <= 10:
            score += 18    # 贵 10% 以内 → 还行
        elif pdiff <= 20:
            score += 10    # 贵 10-20%
        else:
            score += 0     # 贵 20%+ → 劣势
    else:
        score += 12  # 无比价数据 → 中间分

    # 4. 缺口得分（0-20）
    if not card.get("gap_flag", False):
        score += 20
    elif card.get("comp_product_count", 0) < 10:
        score += 15
    else:
        score += 0   # 严重缺口

    # 5. 对齐加分（0-10）——有对齐数据是额外优势
    aligned = card.get("aligned_count", 0)
    if aligned >= 10:
        score += 10
    elif aligned >= 3:
        score += 7
    elif aligned >= 1:
        score += 5

    # 6. 数据可靠度折扣
    reliability_discount = {"high": 1.0, "medium": 0.9, "low": 0.75}
    score *= reliability_discount.get(card.get("data_reliability", "low"), 0.75)

    return int(round(score))


def _score_to_color(score):
    """健康分 → 红绿灯颜色。"""
    if score >= 65:
        return "green"
    elif score >= 35:
        return "yellow"
    else:
        return "red"


def _reliability(comp_count, aligned_count, shop_count):
    """数据可靠性判断。"""
    if comp_count >= 20 and shop_count >= 2 and aligned_count >= 3:
        return "high"
    elif comp_count >= 5 and (shop_count >= 1 or aligned_count >= 1):
        return "medium"
    else:
        return "low"


def build_category_card_data(shops):
    """构建品类卡片墙的完整数据。返回 {"cards": [...], "summary": {...}}。
    纯本地计算，不调用 AI。"""
    import json as _json

    # ------------------------------------------------------------------
    # Step 1: 读取各店竞品数据，按三级类目聚合
    # ------------------------------------------------------------------
    comp_frames = []
    for s in shops:
        comp_f = os.path.join(OUTPUT, f"试跑_{s}", "竞品_归一化后.xlsx")
        if not os.path.exists(comp_f):
            continue
        try:
            cdf = pd.read_excel(comp_f, dtype=str)
            cdf["_shop"] = s
            comp_frames.append(cdf)
        except Exception:
            continue

    if not comp_frames:
        return {"cards": [], "summary": {"total_categories": 0, "green_count": 0,
                "yellow_count": 0, "red_count": 0, "shops": shops,
                "total_comp_products": 0, "total_own_products": 0}}

    all_comp = pd.concat(comp_frames, ignore_index=True)
    # 统一品类列
    cat_col = "类目三级_归一" if "类目三级_归一" in all_comp.columns else "category3"
    # 填充空品类
    all_comp["_cat"] = all_comp[cat_col].fillna("").str.strip().replace("", "(未分类)")
    # 排除未分类
    all_comp_categorized = all_comp[all_comp["_cat"] != "(未分类)"]

    # 竞品按品类聚合
    comp_agg = all_comp_categorized.groupby("_cat").agg(
        comp_product_count=("_cat", "size"),
        comp_shop_count=("_shop", "nunique"),
    ).reset_index()

    # 价格统计（向量化）
    price_col = "price" if "price" in all_comp.columns else None
    if price_col:
        all_comp_categorized_copy = all_comp_categorized.copy()
        all_comp_categorized_copy["_price_num"] = pd.to_numeric(
            all_comp_categorized_copy[price_col], errors="coerce")
        price_stats = all_comp_categorized_copy.groupby("_cat")["_price_num"].agg(
            ["median", "mean", "min", "max"]
        ).reset_index()
        price_stats.columns = ["_cat", "median_price_comp", "avg_price_comp",
                               "min_price_comp", "max_price_comp"]
        comp_agg = comp_agg.merge(price_stats, on="_cat", how="left")

    # 销量统计（向量化）
    sales_col = "sales" if "sales" in all_comp.columns else None
    if sales_col:
        all_comp_categorized_copy = all_comp_categorized.copy()
        all_comp_categorized_copy["_sales_num"] = pd.to_numeric(
            all_comp_categorized_copy[sales_col], errors="coerce")
        sales_stats = all_comp_categorized_copy.groupby("_cat")["_sales_num"].agg(
            ["sum", "mean"]
        ).reset_index()
        sales_stats.columns = ["_cat", "comp_sales_total", "comp_sales_avg"]
        comp_agg = comp_agg.merge(sales_stats, on="_cat", how="left")

    # 热销品 Top3（按品类分组取销量最高3个）
    hot_by_cat = {}
    if sales_col:
        sales_valid = all_comp_categorized_copy.dropna(subset=["_sales_num"])
        sales_valid = sales_valid[sales_valid["_sales_num"] > 0]
        name_col = "name" if "name" in sales_valid.columns else None
        sku_col = "sku_id" if "sku_id" in sales_valid.columns else None
        if name_col and sku_col:
            top3 = sales_valid.sort_values("_sales_num", ascending=False).groupby("_cat").head(3)
            for _, r in top3.iterrows():
                cat = r["_cat"]
                hot_by_cat.setdefault(cat, []).append({
                    "name": _clean(r.get(name_col)),
                    "sales": int(r["_sales_num"]) if pd.notna(r["_sales_num"]) else 0,
                    "price": _clean(r.get(price_col)) if price_col else "",
                    "shop": _clean(r.get("_shop")),
                    "sku": _clean(r.get(sku_col)),
                })

    # ------------------------------------------------------------------
    # Step 2: 读取自家数据，按三级类目聚合
    # ------------------------------------------------------------------
    own_f = os.path.join(OUTPUT, "自家_归一化后.xlsx")
    own_agg = pd.DataFrame()
    total_own = 0
    if os.path.exists(own_f):
        try:
            own = pd.read_excel(own_f, dtype=str)
            # P4.3：渠道重复行按 SKUID 去重（品数/价格统计按唯一 SKU 计，防 ~5x 虚高）
            if "SKUID" in own.columns:
                own = own.drop_duplicates(subset=["SKUID"], keep="first")
            total_own = len(own)
            own_cat = "美团闪购类目三级" if "美团闪购类目三级" in own.columns else "类目三级_归一"
            if own_cat in own.columns:
                own["_cat"] = own[own_cat].fillna("").str.strip().replace("", "(未分类)")
                own_categorized = own[own["_cat"] != "(未分类)"]
                own_agg = own_categorized.groupby("_cat").agg(
                    own_product_count=("_cat", "size"),
                ).reset_index()
                # 自家价格统计
                own_price_col = "美团闪购渠道售价" if "美团闪购渠道售价" in own.columns else None
                if own_price_col:
                    own_categorized_copy = own_categorized.copy()
                    own_categorized_copy["_price_num"] = pd.to_numeric(
                        own_categorized_copy[own_price_col], errors="coerce")
                    own_price_stats = own_categorized_copy.groupby("_cat")["_price_num"].agg(
                        ["median"]
                    ).reset_index()
                    own_price_stats.columns = ["_cat", "median_price_own"]
                    own_agg = own_agg.merge(own_price_stats, on="_cat", how="left")
        except Exception:
            pass

    # 合并竞品+自家数据（提前创建 merged，供 Step 3 使用）
    merged = comp_agg.copy()
    if len(own_agg):
        merged = merged.merge(own_agg, on="_cat", how="left")
    else:
        merged["own_product_count"] = 0
        merged["median_price_own"] = None

    # ------------------------------------------------------------------
    # Step 3: 读取对齐数据，计算覆盖率（模糊匹配品类标签）
    # ------------------------------------------------------------------
    lib_f = os.path.join(OUTPUT, "对齐汇总库.xlsx")
    aligned_by_cat = {}
    if os.path.exists(lib_f):
        try:
            spu_lib = pd.read_excel(lib_f, sheet_name="SPU匹配层", dtype=str)
            if "竞品类目" in spu_lib.columns:
                # SPU 层品类 → 对齐数
                spu_aligned = spu_lib.groupby(
                    spu_lib["竞品类目"].fillna("").str.strip()
                ).size().to_dict()
                spu_aligned.pop("", None)

                # 对每张卡片的三级类目做模糊匹配（子串包含）
                # 例：SPU层"水暖管件" 匹配 竞品三级类目"水暖管件及配件"
                all_cats = set(merged["_cat"])
                for spu_cat, cnt in spu_aligned.items():
                    # 精确匹配优先
                    if spu_cat in all_cats:
                        aligned_by_cat[spu_cat] = aligned_by_cat.get(spu_cat, 0) + cnt
                        continue
                    # 子串匹配：SPU品类名包含在竞品三级类目中，或反过来
                    matched = False
                    for comp_cat in all_cats:
                        if comp_cat in spu_cat or spu_cat in comp_cat:
                            aligned_by_cat[comp_cat] = aligned_by_cat.get(comp_cat, 0) + cnt
                            matched = True
                            break
                    # 都没匹配上也不丢（保留到 SPU 品类本身）
                    if not matched:
                        aligned_by_cat[spu_cat] = aligned_by_cat.get(spu_cat, 0) + cnt
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Step 4: 读取竞争分析报告，补充缺口
    # ------------------------------------------------------------------
    gap_cats = set()
    ra_f = os.path.join(OUTPUT, "竞争分析报告.xlsx")
    if os.path.exists(ra_f):
        try:
            cat_gap = pd.read_excel(ra_f, sheet_name="类目缺口", dtype=str)
            if len(cat_gap) and "类目" in cat_gap.columns:
                gap_cats = set(cat_gap["类目"].fillna("").str.strip())
                gap_cats.discard("")
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Step 5: 构建卡片
    # ------------------------------------------------------------------

    # 填充缺失
    merged["comp_product_count"] = merged["comp_product_count"].fillna(0).astype(int)
    merged["comp_shop_count"] = merged["comp_shop_count"].fillna(0).astype(int)
    merged["own_product_count"] = merged["own_product_count"].fillna(0).astype(int)
    if "comp_sales_total" not in merged.columns:
        merged["comp_sales_total"] = 0
    else:
        merged["comp_sales_total"] = merged["comp_sales_total"].fillna(0)

    cards = []
    for _, row in merged.iterrows():
        cat = row["_cat"]
        if not cat or cat == "(未分类)":
            continue

        comp_count = int(row["comp_product_count"])
        own_count = int(row["own_product_count"])
        aligned_count = aligned_by_cat.get(cat, 0)
        shop_count = int(row["comp_shop_count"])
        sales_total = int(row.get("comp_sales_total", 0)) if pd.notna(row.get("comp_sales_total")) else 0

        # 价差率
        price_diff_pct = None
        median_comp = row.get("median_price_comp")
        median_own = row.get("median_price_own")
        if pd.notna(median_comp) and pd.notna(median_own):
            mc, mo = float(median_comp), float(median_own)
            if mc > 0:
                price_diff_pct = round((mo - mc) / mc * 100, 1)

        gap_flag = cat in gap_cats
        reliability = _reliability(comp_count, aligned_count, shop_count)
        coverage_rate = round(aligned_count / comp_count, 3) if comp_count > 0 else 0

        card = {
            "cat": cat,
            "comp_product_count": comp_count,
            "own_product_count": own_count,
            "aligned_count": aligned_count,
            "comp_shop_count": shop_count,
            "comp_sales_total": sales_total,
            "coverage_rate": coverage_rate,
            "median_price_comp": round(float(median_comp), 1) if pd.notna(median_comp) else None,
            "median_price_own": round(float(median_own), 1) if pd.notna(median_own) else None,
            "avg_price_comp": round(float(row.get("avg_price_comp", 0)), 1) if pd.notna(row.get("avg_price_comp")) else None,
            "price_diff_pct": price_diff_pct,
            "gap_flag": gap_flag,
            "data_reliability": reliability,
            "hot_items": hot_by_cat.get(cat, []),
            # 占位（AI 填充）
            "diagnosis": "",
            "actions": [],
        }

        # 健康分 & 颜色
        card["health_score"] = compute_health_score(card)
        card["color"] = _score_to_color(card["health_score"])

        # 关键指标摘要文字
        parts = []
        if aligned_count > 0:
            parts.append(f"对齐{aligned_count}品")
        else:
            parts.append("未对齐")
        parts.append(f"竞品{comp_count}品")
        if sales_total > 0:
            parts.append(f"月售{sales_total}")
        if price_diff_pct is not None:
            parts.append(f"价差{price_diff_pct:+.0f}%")
        card["key_metric"] = " | ".join(parts)

        cards.append(card)

    # 筛选：排除商品数太少的品类（<3）
    cards = [c for c in cards if c["comp_product_count"] >= 3]

    # 智能筛选：保留最多 ~40 张高价值卡片
    # 优先级：有对齐数据 > 有缺口标记 > 竞品品数多 > 月售高
    def _card_priority(c):
        """排序键：优先级高的排前面（值大=优先）。"""
        p = 0
        if c["aligned_count"] > 0:
            p += 10000
        if c["gap_flag"]:
            p += 5000
        p += c["comp_product_count"] * 10
        p += c["comp_sales_total"]
        return p

    cards.sort(key=_card_priority, reverse=True)
    cards = cards[:40]  # 最多保留 40 张

    # 最终排序：红 → 黄 → 绿，同色按分数升序（最差的排最前）
    color_order = {"red": 0, "yellow": 1, "green": 2}
    cards.sort(key=lambda c: (color_order.get(c["color"], 1), c["health_score"]))

    # 汇总
    green_n = sum(1 for c in cards if c["color"] == "green")
    yellow_n = sum(1 for c in cards if c["color"] == "yellow")
    red_n = sum(1 for c in cards if c["color"] == "red")

    return {
        "cards": cards,
        "summary": {
            "total_categories": len(cards),
            "green_count": green_n,
            "yellow_count": yellow_n,
            "red_count": red_n,
            "total_comp_products": int(merged["comp_product_count"].sum()),
            "total_own_products": total_own,
            "shops": shops,
        },
    }


def call_category_diagnosis(card_data, config):
    """调 DeepSeek 生成品类诊断。返回 diagnoses 列表。
    每个元素: {"cat": "品类名", "diagnosis": "诊断文本", "actions": [{"text": "建议", "link": {...}}]}
    """
    import json as _json

    # 构建精简版品类数据（去掉 hot_items 的冗余字段，控制 token 量）
    compact = []
    for c in card_data["cards"]:
        item = {
            "cat": c["cat"],
            "comp_count": c["comp_product_count"],
            "own_count": c["own_product_count"],
            "aligned": c["aligned_count"],
            "shops": c["comp_shop_count"],
            "sales": c["comp_sales_total"],
            "med_price_comp": c["median_price_comp"],
            "med_price_own": c["median_price_own"],
            "price_diff_pct": c["price_diff_pct"],
            "gap": c["gap_flag"],
            "reliability": c["data_reliability"],
            "hot_items": [{"name": h["name"], "sales": h["sales"], "price": h["price"]}
                          for h in c.get("hot_items", [])],
        }
        compact.append(item)

    cards_json = _json.dumps(compact, ensure_ascii=False, indent=1)

    user_text = (
        f"以下是本地竞品分析的品类数据摘要（共 {len(compact)} 个品类）：\n"
        f"{cards_json}\n\n"
        f"请为每个品类输出诊断和建议。"
    )

    from ai_client import call_llm
    raw = call_llm(
        "category_diag",
        [
            {"role": "system", "content": CATEGORY_DIAG_SYSTEM_PROMPT},
            {"role": "user", "content": user_text},
        ],
        config=config,
        timeout=180,
        temperature=0.5,
        max_tokens=4000,
    )

    # 解析 JSON 响应
    try:
        # 尝试直接解析
        parsed = _json.loads(raw)
    except (ValueError, _json.JSONDecodeError):
        # 尝试从 Markdown 代码块中提取
        import re
        m = re.search(r"```(?:json)?\s*\n?(.*?)\n?```", raw, re.DOTALL)
        if m:
            try:
                parsed = _json.loads(m.group(1))
            except (ValueError, _json.JSONDecodeError):
                parsed = {}
        else:
            # 尝试找到第一个 [ 和最后一个 ]
            start = raw.find("[")
            end = raw.rfind("]")
            if start >= 0 and end > start:
                try:
                    parsed = {"diagnoses": _json.loads(raw[start:end + 1])}
                except (ValueError, _json.JSONDecodeError):
                    parsed = {}
            else:
                parsed = {}

    diagnoses = parsed.get("diagnoses", parsed if isinstance(parsed, list) else [])

    # 规范化每条诊断
    result = []
    for d in diagnoses:
        if not isinstance(d, dict):
            continue
        actions_raw = d.get("actions", [])
        actions = []
        for a in actions_raw:
            if isinstance(a, str):
                actions.append({"text": a, "link": None})
            elif isinstance(a, dict):
                actions.append({
                    "text": a.get("text", str(a)),
                    "link": a.get("link"),
                })
        result.append({
            "cat": d.get("cat", ""),
            "diagnosis": d.get("diagnosis", ""),
            "actions": actions,
        })
    return result


if __name__ == "__main__":
    import argparse
    _p = argparse.ArgumentParser(description="AI 经营分析（方案A）")
    _p.add_argument("--shops", help="店铺列表，逗号分隔")
    _p.add_argument("--directions", default="hot,price,display,insight", help="方向: hot,price,display,insight")
    _p.add_argument("--mock", action="store_true", help="本地模拟（不调用真实接口）")
    _a = _p.parse_args()
    shops = [s.strip() for s in (_a.shops or "").split(",") if s.strip()]
    dirs = [d.strip() for d in _a.directions.split(",") if d.strip() in DIRECTIONS]
    text, st, path = generate_report(shops, dirs, mock=_a.mock)
    print(f"报告已生成: {path}（板块: {st['sections']}）")
    print(text[:500])

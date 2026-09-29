# -*- coding: utf-8 -*-
"""
闪电仓选品 - 人工确认 Web 面板
左右对照：左侧竞品品固定，右侧可搜索自家商品库，用户对照后确认或标记无对应。
启动：python scripts/web_panel.py  （自动打开浏览器 http://127.0.0.1:5000）
"""
import os
import json
import re
import shutil
import sys
import time
import threading
import uuid
import webbrowser
import hashlib
from datetime import datetime
from urllib.parse import quote, urlparse

import pandas as pd
from flask import (Flask, render_template, request, jsonify, abort, send_from_directory,
                   send_file, redirect, g, Response)
from werkzeug.utils import secure_filename

# 导入适配模块（与本文件同目录，import_adapter 作为模块可直接导入）
import import_adapter
from summary_library import (LIBRARY_FILE, load_library, save_library,
                             upsert_manual, upsert_rows)
import snapcache as _snap   # 派生结论/Excel 读取的磁盘快照层（只当缓存，带依赖指纹）
import sales_text as _ST    # 销量文本唯一解析口径（一列里混着「月售35」与「35人想买」）
import store_wh as _store_wh  # 店↔仓显式指认表（补货参考销量归属；禁止静默回退合计）

# ---------------------------------------------------------------------------
# 路径配置
# ---------------------------------------------------------------------------
# 项目根目录：相对脚本位置动态定位（绿色打包版/原目录均可运行）
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
SCRIPTS_DIR = os.path.join(BASE_DIR, "scripts")

CONFIRM_POOL_FILE = os.path.join(OUTPUT_DIR, "对齐_人工确认池.xlsx")
COMPETITOR_FILE = os.path.join(OUTPUT_DIR, "竞品_归一化后.xlsx")
OWN_FILE = os.path.join(OUTPUT_DIR, "自家_归一化后.xlsx")
ALIGN_SUMMARY_FILE = os.path.join(OUTPUT_DIR, "对齐_汇总.xlsx")
NO_MATCH_FILE = os.path.join(OUTPUT_DIR, "无对应品清单.xlsx")
CATEGORY_MAP_FILE = os.path.join(OUTPUT_DIR, "类目映射表.xlsx")
PROGRESS_FILE = os.path.join(SCRIPTS_DIR, "progress.json")

TEMPLATE_DIR = os.path.join(SCRIPTS_DIR, "templates")
STATIC_DIR = os.path.join(SCRIPTS_DIR, "static")
UPLOAD_DIR = os.path.join(BASE_DIR, "config", "import_uploads")
IMPORT_RECORDS_FILE = os.path.join(UPLOAD_DIR, ".import_records.json")
IMG_CACHE_DIR = os.path.join(OUTPUT_DIR, "img_cache")

PAGE_SIZE = 10
SEARCH_LIMIT = 60

# 无试跑目录时单店池的归属店（兼容旧单店链路）
DEFAULT_SHOP = "示例竞店B"

app = Flask(__name__, template_folder=TEMPLATE_DIR, static_folder=STATIC_DIR)

# ---------------------------------------------------------------------------
# 生产化：日志（方案A）——按天落盘 output/logs/panel.log，保留 14 天
# ---------------------------------------------------------------------------
import logging
from logging.handlers import TimedRotatingFileHandler

LOG_DIR = os.path.join(OUTPUT_DIR, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
_fh = TimedRotatingFileHandler(os.path.join(LOG_DIR, "panel.log"),
                               when="midnight", backupCount=14, encoding="utf-8")
_fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s [%(threadName)s] %(message)s"))
logging.basicConfig(level=logging.INFO, handlers=[_fh])
logging.getLogger("werkzeug").addHandler(_fh)
_logger = logging.getLogger("panel")


# ---------------------------------------------------------------------------
# 生产化：HTTP Basic 共享密码鉴权（方案A）
# 环境变量 PANEL_USER / PANEL_PASS 设置后启用；不设置=单机模式不拦截（现状行为）。
# ---------------------------------------------------------------------------
PANEL_USER = os.environ.get("PANEL_USER", "").strip()
PANEL_PASS = os.environ.get("PANEL_PASS", "").strip()
_AUTH_ON = bool(PANEL_USER and PANEL_PASS)
# 免鉴权路径：静态资源 / 健康检查 / 版本指纹（launch_panel 本机探测用）
_AUTH_EXEMPT = {"/healthz", "/api/version"}


@app.before_request
def _auth_gate():
    # P4.3：own 数据新鲜度检查（bind 类外部脚本重写后自动重载+失效缓存）
    _check_own_data_fresh()
    if not _AUTH_ON:
        return None
    p = (request.path or "").rstrip("/") or "/"
    if request.path.startswith("/static/") or request.path.startswith("/healthz") \
            or p in _AUTH_EXEMPT:
        return None
    a = request.authorization
    if a and a.username == PANEL_USER and a.password == PANEL_PASS:
        return None
    return ("", 401, {"WWW-Authenticate": 'Basic realm="panel"'})


@app.route("/healthz")
def healthz():
    """探活（免鉴权）：服务存活 + 是否启用鉴权。"""
    return jsonify({"ok": True, "auth": _AUTH_ON, "version": PANEL_VERSION})


# ---------------------------------------------------------------------------
# 生产化：写端点全局文件锁（方案A）——POST/PUT/DELETE 自动串行化，
# 防 Excel/progress.json 并发写坏。锁在 teardown 释放；超时返回 503。
# ---------------------------------------------------------------------------
from data_lock import data_lock as _data_lock_ctx

_LOCK_TIMEOUT = int(os.environ.get("PANEL_LOCK_TIMEOUT", "600"))
# 免锁白名单的划分标准是**会不会写核心数据文件**（progress / 汇总库 / 候选池 / 归一化表），
# 不是"耗时长短"：只等外部大模型返回的请求持锁，会把全面板的保存动作堵到 600 秒超时
# （员工机上表现就是"点了没反应"）。这类端点写的都是自己的产物文件，且写入是原子替换。
# 会改核心数据的（如 /api/ai/prescreen 要写 progress + 出报表）必须留在锁内。
_LOCK_EXEMPT = (
    "/api/top50/advice",       # 生成前50品建议（只读数据 + 写建议 xlsx）
    "/api/assistant/analyze",  # 抽屉场景分析（只读缓存产物）
    "/api/assistant/ask",      # 就某条洞察追问（只读凭据行）
    "/api/assistant/chat",     # 聊天助手（工具全只读，读当期缓存，0 写库）
    "/api/assistant/sessions", # 会话管理（只写自己的 ai_chat 会话文件，原子替换）
    "/api/insight/interpret",  # 台账 AI 解读（只原子更新 台账 json）
    "/api/pool/llm_verify",    # 大模型核实同品（结果缓存进候选池 json，原子写）
    "/api/ai/analysis",        # AI 经营分析报告（写自己的草稿文件）
    "/api/ai/category_diag",   # 品类诊断（可含 AI）
    "/api/ai/export",          # 报告导出
    "/api/advice/generate",    # 已对齐建议生成
)


@app.before_request
def _acquire_write_lock():
    if request.method in ("POST", "PUT", "DELETE", "PATCH"):
        if any(request.path.startswith(p) for p in _LOCK_EXEMPT):
            return None
        try:
            g._wlock = _data_lock_ctx(timeout=_LOCK_TIMEOUT)
            g._wlock.__enter__()
        except TimeoutError as e:
            return jsonify({"ok": False, "msg": str(e)}), 503
    return None


@app.teardown_appcontext
def _release_write_lock(exc=None):
    lk = getattr(g, "_wlock", None)
    if lk is not None:
        try:
            lk.__exit__(None, None, None)
        except Exception:
            pass
        g._wlock = None


@app.template_filter("img")
def img_filter(u):
    """图片地址过滤器：外站 http(s) 图 → 走本地代理缓存；本地/空值原样返回。"""
    if not u:
        return ""
    u = str(u)
    if u.startswith(("http://", "https://")):
        return "/proxy_img?url=" + quote(u, safe="")
    return u


# ---------------------------------------------------------------------------
# 全局数据（启动时加载一次）
# ---------------------------------------------------------------------------
pool_df = None          # 多店合并人工确认池（含 竞店名称/__rowid 列）
own_df = None           # 自家_归一化后
competitor_lookup = {}  # {店名: {Sku_id: row dict}}
own_lookup = {}         # SKUID -> row dict
own_spu_map = {}        # SKUID -> (自家品组ID, 品名)（从对齐_SPU细化 品对品 读）
aligned_map = {}        # 竞品Sku_id -> {自家SKUID,自家商品名,匹配层,匹配依据}（汇总库有效行）
spu_layer_keys = set()  # {(竞店名称, 竞品品ID)}（汇总库SPU层有效行，品已锁定）
spu_layer_info = {}     # {(竞店名称, 竞品品ID): 自家品名}（品锁定行的自家信息）
cat_map_cache = {}      # 竞品类目 -> 建议映射（类目映射表缓存）
own_categories = []
own_brands = []
progress = {"confirmed": {}, "no_match": {}}

# 结果级内存缓存（都在 _invalidate_stats 里统一作废：人工确认/导入/改名/删店/回填即失效）
_browse_rows_cache = {}   # {店: 竞品库浏览行列表}
_browse_meta_cache = {}   # {店: 匹配状态计数}
_slow_rank_cache = {}     # {(门店, id(在售集合)): 滞销榜结果}
_zone_rank_cache = {}     # {(区名, 勾选子集元组): build_zone_rank 原始结果}


def _safe(v):
    """把 NaN/None 转为空字符串，其余转 str。"""
    if v is None:
        return ""
    try:
        if pd.isna(v):
            return ""
    except (TypeError, ValueError):
        pass
    return str(v)


def _xread(path, **kw):
    """pd.read_excel 的快照版：同一份 Excel 只在它自己变化后重读一次。
    实测 自家_归一化后.xlsx（44456 行）14.3s → 命中快照 0.05s。
    依赖=该文件本身（mtime+size），所以任何一次重新导入/回填都会让快照自动作废。
    快照名由「绝对路径 + 读取参数」哈希得出：同一文件不同 sheet/参数互不串用，
    不同文件即使同名（如默认店池 vs 试跑店池）也各自独立。
    """
    key = "%s|%s" % (os.path.normcase(os.path.abspath(path)),
                     ",".join("%s=%s" % (k, v) for k, v in sorted(kw.items())))
    name = "xlsx_" + hashlib.sha1(key.encode("utf-8")).hexdigest()[:16]
    obj, _hit = _snap.remember(name, [path], lambda: pd.read_excel(path, **kw))
    return obj


# 派生结论的可信边界：这些上游任一变化 → 快照作废。
# 取「宁多勿漏」——多列一个文件只是多算一次，漏列一个就是拿旧结论冒充新结论。
# 但按 kind 收窄：纯竞品侧的榜单不该被自家销量回填带废（见 _derived_deps）。
_DERIVED_EXTRA = (ALIGN_SUMMARY_FILE, CATEGORY_MAP_FILE,
                  os.path.join(OUTPUT_DIR, "对齐修正记录.xlsx"),
                  os.path.join(OUTPUT_DIR, "ai_prescreen.json"))


def _derived_deps(shop, kind="product"):
    """按结论类型给依赖：只列「这份结论真的会读的东西」。

    top500 = 纯竞品侧品级聚合（build_top50 只读该店归一表）→ 自家销量回填不该让它重算；
    detail / product = 消费自家月销与人工对应 → 自家表、汇总库、progress 缺一不可。
    注：若日后 top50 开始读自家数据，务必把它加进它那一支（改了 top50.py 源码也会因
    算法指纹变化而作废，但「间接经其他模块读到自家数据」这种要靠人来补）。
    """
    trial = os.path.join(OUTPUT_DIR, "试跑_" + shop)
    comp_norm = os.path.join(trial, "竞品_归一化后.xlsx")
    if kind == "top500":
        return [comp_norm]
    return ([OWN_FILE, LIBRARY_FILE, PROGRESS_FILE, comp_norm,
             os.path.join(trial, "对齐_人工确认池.xlsx")] + list(_DERIVED_EXTRA))


def _code_fp(*mods):
    return "|".join(_snap.code_fingerprint(m) for m in mods)


_fp_derived_cache = {}


def _fp_derived():
    """榜单/对比/品卡三条计算链的算法指纹：任一生成器模块改过，旧快照即作废。"""
    if "v" not in _fp_derived_cache:
        import build_comparison
        import build_product_analysis
        import sales_text
        import spu_semantic
        import summary_library
        import top50
        # sales_text 也进指纹：上面这些模块的销量解析都委托给它，只改它就该全部作废
        _fp_derived_cache["v"] = _code_fp(top50, build_comparison, build_product_analysis,
                                          spu_semantic, summary_library, sales_text)
    return _fp_derived_cache["v"]


def _derived(kind, shop, builder, code=None):
    """派生结论三层取数：内存 → 磁盘快照（按依赖指纹）→ 现场重算并落盘。

    重启面板不再重算历史结论；人工确认/重新导入/销量回填一改，指纹即变，
    下一次自动重算——旧快照只可能被「作废」，不可能被「当成新的用」。
    """
    key = (shop, kind)
    if key in _compare_cache:
        return _compare_cache[key]
    payload, _hit = _snap.remember("derived_%s_%s" % (kind, shop),
                                   _derived_deps(shop, kind), builder, code=code)
    _compare_cache[key] = payload
    return payload


_fp_advice_cache = {}


def _advice_code_fp():
    """建议链算法指纹：改了提示词/建议组装/销量口径，旧建议就该判过期。"""
    if "v" not in _fp_advice_cache:
        import ai_top50
        import assistant_brain
        import summary_library
        import sales_text
        _fp_advice_cache["v"] = _code_fp(ai_top50, assistant_brain, summary_library, sales_text)
    return _fp_advice_cache["v"]


def _advice_deps(shop):
    """建议的可信边界 = build_summary 真正读到的东西（榜单源表 + 自家表 + 人工确认 + 汇总库）。"""
    import summary_library
    return (_derived_deps(shop, "top500")
            + [OWN_FILE, PROGRESS_FILE, summary_library.LIBRARY_FILE])


def _advice_slot(shop):
    return "advice_top50_" + shop


def advice_state(shop):
    """AI 建议三层取数：内存 → 磁盘快照（指纹一致才算现行结论）→ 空。

    和 _derived() 的唯一区别是**这里绝不现场调模型**：调模型要花钱，必须人点按钮。
    所以快照过期时只回 stale + 上次生成时间，不回旧行——页面宁可显示"已过期，请重新生成"，
    也不能把上一轮数据算出的建议当成现状。
    """
    deps = _advice_deps(shop)
    fp = _snap.key_of(deps, _advice_code_fp())
    slot = _advice_slot(shop)
    if shop in _top50_advice_cache:
        return {"rows": _top50_advice_cache[shop], "stale": False, "at": "", "src": "内存"}
    rows = _snap.load(slot, fp)
    if rows is not None:
        _top50_advice_cache[shop] = rows
        return {"rows": rows, "stale": False,
                "at": (_snap.peek(slot) or {}).get("built_at", ""), "src": "快照"}
    meta = _snap.peek(slot) or {}
    if meta.get("fingerprint") and meta.get("fingerprint") != fp:
        return {"rows": [], "stale": True, "at": meta.get("built_at", ""), "src": "过期"}
    return {"rows": [], "stale": False, "at": "", "src": "未生成"}


def remember_advice(shop, rows):
    """生成成功后落盘：重启、切店、无关写库都不必再花一次模型钱。"""
    _top50_advice_cache[shop] = rows
    deps = _advice_deps(shop)
    fp = _snap.key_of(deps, _advice_code_fp())
    _snap.store(_advice_slot(shop), fp, rows, deps=deps, note="AI建议(top50)")
    return fp


def _load_multi_shop_pool():
    """扫描 output/试跑_<店>/对齐_人工确认池.xlsx 自动合并（加 竞店名称/__rowid 列）；
    无试跑目录时回退单店文件（默认店），兼容旧单店链路。"""
    shops_pool = []
    trial_dirs = sorted([d for d in os.listdir(OUTPUT_DIR)
                         if d.startswith("试跑_") and os.path.isdir(os.path.join(OUTPUT_DIR, d))])
    for d in trial_dirs:
        f = os.path.join(OUTPUT_DIR, d, "对齐_人工确认池.xlsx")
        if not os.path.exists(f):
            continue
        shop = d[len("试跑_"):]
        df = _xread(f, dtype=str)
        if len(df):
            df["竞店名称"] = shop
            df["__rowid"] = range(len(df))
            shops_pool.append(df)
    # 单店池文件（旧单店链路）始终并入，归属默认店；
    # 避免多店时代历史确认记录因试跑目录存在而无行可显示
    if os.path.exists(CONFIRM_POOL_FILE):
        df = _xread(CONFIRM_POOL_FILE, dtype=str)
        if len(df):
            df["竞店名称"] = DEFAULT_SHOP
            df["__rowid"] = range(len(df))
            shops_pool.append(df)
    if not shops_pool:
        return pd.DataFrame()
    pool = pd.concat(shops_pool, ignore_index=True)
    for col in ["竞品品ID", "竞品Sku_id", "自家SPUID", "自家SKUID"]:
        if col in pool.columns:
            pool[col] = pool[col].apply(_safe)
    for col in ["确认结果", "人工备注"]:
        if col in pool.columns:
            pool[col] = pool[col].astype(object)
    return pool


def load_data():
    """加载多店人工池、自家表，构建按店竞品查找字典与匹配状态映射。"""
    global pool_df, own_df, competitor_lookup, own_lookup
    global own_spu_map, aligned_map
    global own_categories, own_brands

    pool_df = _load_multi_shop_pool()
    pool_df.reset_index(drop=True, inplace=True)
    # 列向量预构建（替代逐行 iloc）+ 缓存重置
    global _pool_shop_list, _pool_sku_list, _pool_value_cache, _pool_status_cache, _lib_cache
    _pool_shop_list = pool_df["竞店名称"].astype(str).tolist() if len(pool_df) else []
    _pool_sku_list = pool_df["竞品Sku_id"].astype(str).tolist() if len(pool_df) else []
    _pool_value_cache = None
    _pool_conf_cache = None
    _pool_status_cache = None
    _lib_cache = None

    own_df = _xread(OWN_FILE, dtype=str)
    for col in ["SPUID", "SKUID"]:
        if col in own_df.columns:
            own_df[col] = own_df[col].apply(_safe)

    # 竞品按店查找：{店名: {Sku_id: row}}
    competitor_lookup = {}
    trial_dirs = sorted([d for d in os.listdir(OUTPUT_DIR)
                         if d.startswith("试跑_") and os.path.isdir(os.path.join(OUTPUT_DIR, d))])
    for d in trial_dirs:
        f = os.path.join(OUTPUT_DIR, d, "竞品_归一化后.xlsx")
        if not os.path.exists(f):
            continue
        shop = d[len("试跑_"):]
        df = _xread(f, dtype=str)
        lk = {}
        # 列名兼容：新版归一化表为英文列 sku_id，旧表为 Sku_id
        sku_col = "sku_id" if "sku_id" in df.columns else (
            "Sku_id" if "Sku_id" in df.columns else None)
        if sku_col is None:
            continue
        for _, row in df.iterrows():
            sku = _safe(row.get(sku_col))
            if sku:
                lk[sku] = row
        competitor_lookup[shop] = lk
    if not competitor_lookup and os.path.exists(COMPETITOR_FILE):
        df = _xread(COMPETITOR_FILE, dtype=str)
        lk = {}
        for _, row in df.iterrows():
            sku = _safe(row.get("Sku_id"))
            if sku:
                lk[sku] = row
        competitor_lookup[DEFAULT_SHOP] = lk

    # 自家查找字典：SKUID -> 行数据
    own_lookup = {}
    for _, row in own_df.iterrows():
        sku = _safe(row.get("SKUID"))
        if sku:
            own_lookup[sku] = row

    # 自家品组价格带 {SPUID: (min售价, max售价)}（比价展示用；向量化聚合）
    global _own_price_cache
    _own_price_cache = {}
    try:
        if len(own_df) and "SPUID" in own_df.columns and "美团闪购渠道售价" in own_df.columns:
            _t = own_df[["SPUID", "美团闪购渠道售价"]].copy()
            _t["SPUID"] = _t["SPUID"].astype(str).str.strip()
            _t["_p"] = pd.to_numeric(_t["美团闪购渠道售价"].astype(str)
                                     .str.replace(r"[^\d.]", "", regex=True)
                                     .replace("", None), errors="coerce")
            _t = _t.dropna(subset=["_p"])
            _t = _t[(_t["SPUID"].str.len() > 0) & (_t["SPUID"] != "nan")]
            if len(_t):
                _g = _t.groupby("SPUID")["_p"].agg(["min", "max"])
                _own_price_cache = {k: (round(r["min"], 2), round(r["max"], 2))
                                    for k, r in _g.iterrows()}
    except Exception as _e:
        print("[load_data] 品组价格带构建失败:", _e)

    # 类目、品牌下拉选项
    cat_col = "类目三级_归一" if "类目三级_归一" in own_df.columns else "美团闪购类目三级"
    own_categories = sorted(
        [c for c in own_df[cat_col].dropna().unique().tolist() if _safe(c)]
    ) if cat_col in own_df.columns else []
    brand_col = "品牌（抽取）" if "品牌（抽取）" in own_df.columns else "品牌名称"
    own_brands = sorted(
        [b for b in own_df[brand_col].dropna().unique().tolist() if _safe(b)]
    ) if brand_col in own_df.columns else []

    # 自家SKU -> 自家品组ID（从任一店的 对齐_SPU细化 品对品 读，自家品组全局一致）
    own_spu_map = {}
    for d in sorted(os.listdir(OUTPUT_DIR)):
        if not d.startswith("试跑_"):
            continue
        f = os.path.join(OUTPUT_DIR, d, "对齐_SPU细化.xlsx")
        if not os.path.exists(f):
            continue
        try:
            pairs = _xread(f, sheet_name="品对品", dtype=str)
        except Exception:
            continue
        for _, r in pairs.iterrows():
            gid = _safe(r.get("自家SPU组ID"))
            for sku in str(r.get("自家SKU列表", "")).split(","):
                sku = sku.strip()
                if sku and sku.lower() not in ("nan", ""):
                    own_spu_map[sku] = gid
        break  # 任一店文件即可

    # 匹配状态映射：汇总库 SKU层（状态=有效）-> 竞品Sku_id 已对应
    aligned_map = {}
    global spu_layer_keys, spu_layer_info, cat_map_cache
    spu_layer_keys = set()
    spu_layer_info = {}
    cat_map_cache = {}
    try:
        spu_df_lib, sku_df_lib = load_library()
        for _, r in sku_df_lib.iterrows():
            sku = _safe(r.get("竞品Sku_id"))
            if sku and _safe(r.get("状态")) == "有效":
                aligned_map[sku] = {
                    "自家SKUID": _safe(r.get("自家SKUID")),
                    "自家商品名": _safe(r.get("自家商品名")),
                    "匹配层": _safe(r.get("匹配层")),
                    "匹配依据": _safe(r.get("匹配依据")),
                }
        for _, r in spu_df_lib.iterrows():
            if _safe(r.get("状态")) == "有效":
                key = (_safe(r.get("竞店名称")), _safe(r.get("竞品品ID")))
                spu_layer_keys.add(key)
                if _safe(r.get("自家品名")):
                    spu_layer_info.setdefault(key, _safe(r.get("自家品名")))
    except Exception:
        aligned_map = {}
    # 类目映射表缓存（收敛层级"类已锁定"判定用）
    try:
        if os.path.exists(CATEGORY_MAP_FILE):
            cm_df = pd.read_excel(CATEGORY_MAP_FILE, dtype=str).fillna("")
            if "三级类目" in cm_df.columns and "建议映射" in cm_df.columns:
                for _, r in cm_df.iterrows():
                    sug = _safe(r.get("建议映射"))
                    if sug and sug != "需人工映射":
                        cat_map_cache[_safe(r.get("三级类目"))] = sug
    except Exception:
        pass
    # 品类内价格分位分布（全店 SKU 到手价按类目；比价展示用）
    _build_cat_prices()


def _shop_of(idx):
    """合并池第 idx 行所属竞店（列向量索引，快）。"""
    return _pool_shop_list[idx] if idx < len(_pool_shop_list) else _safe(pool_df.iloc[idx].get("竞店名称")) or DEFAULT_SHOP


def _clean_sales(v):
    """月售清洗：统一为 '数字' 或 '数字(意向)' 两种格式（供列表分档排序与展示）。

    解析委托 sales_text——这里原本是站内第四套自己写的 `re.search(r"(\\d[\\d,.]*)")`，
    不认单位："1.5k" 被读成 1.5。
    """
    if not _ST.has_num(v):
        return ""
    n = _ST.num(v)
    txt = f"{n:,.0f}" if n == int(n) else f"{n:,.2f}"
    return txt + "(意向)" if _ST.is_intent(v) else txt




def _extract_num(v):
    """从单元格提取排序数字：取全部数字的中位数（'9~59'->34，'21(意向)'->21，'11'->11）。
    无任何数字返回 None（排序时永远排最后）。"""
    s = str(v).strip()
    if not s or s.lower() in ("nan", "none"):
        return None
    import re as _re
    nums = [float(m) for m in _re.findall(r"\d+(?:\.\d+)?", s)]
    if not nums:
        return None
    nums.sort()
    mid = len(nums) // 2
    if len(nums) % 2 == 0:
        return (nums[mid - 1] + nums[mid]) / 2   # 偶数个：中间两数平均（区间中心）
    return nums[mid]


def _discount_badge(price, orig):
    """折扣徽章：到手价/原价 → '8.5折'；价格缺失或折扣>=95% 或>10折 返回 ''。"""
    p, o = _extract_num(price), _extract_num(orig)
    if p is None or o is None or o <= 0:
        return ""
    r = p / o
    if r >= 0.95:
        return ""
    if r <= 0.10:
        return "1折"
    return f"{round(r * 10, 1):g}折"


def _price_diff(price, own_price):
    """价差 = 竞品到手价 - 自家售价（正=竞品贵/自家有价格优势，负=竞品便宜）。
    任一价格缺失返回 None。"""
    p, o = _extract_num(price), _extract_num(own_price)
    if p is None or o is None:
        return None
    return round(p - o, 1)


def _load_price_hist():
    """价格观测历史 → {(店, 竞品sku): (t, 上次到手价)}（取每 SKU 最新一次观测）。
    文件不存在 → 空；mtime 变化才重读（导入触发 append 后自动刷新）。"""
    global _price_hist_cache, _price_hist_mtime
    f = os.path.join(OUTPUT_DIR, "竞店标准库", "价格观测历史.jsonl")
    if not os.path.exists(f):
        if _price_hist_cache is None:
            _price_hist_cache = {}
        return _price_hist_cache
    mt = os.path.getmtime(f)
    if _price_hist_cache is not None and mt == _price_hist_mtime:
        return _price_hist_cache
    hist = {}
    try:
        with open(f, encoding="utf-8") as fh:
            for ln in fh:
                try:
                    rec = json.loads(ln)
                except (ValueError, TypeError):
                    continue
                key = (_safe(rec.get("店")), _safe(rec.get("id")))
                p = _extract_num(rec.get("到手价") or rec.get("折后价"))
                if not key[1] or p is None:
                    continue
                t = _safe(rec.get("t"))
                cur = hist.get(key)
                if cur is None or t >= cur[0]:
                    hist[key] = (t, p)
    except Exception as e:
        print("[价格观测] 读取失败:", e)
    _price_hist_mtime = mt
    _price_hist_cache = hist
    return hist


def _price_delta(shop, sku, cur_price, hist=None):
    """较上次观测的价格变动 = 当前到手价 - 上次到手价；无观测/无变化/价格缺失 → None。
    hist 可传入 _load_price_hist() 结果，避免行循环内重复 getmtime。"""
    p = _extract_num(cur_price)
    if p is None:
        return None
    if hist is None:
        hist = _load_price_hist()
    last = hist.get((shop, str(sku)))
    if last is None:
        return None
    d = round(p - last[1], 1)
    return d if d != 0 else None


# P1.2/P1.3：完整观测历史 {店: {sku: [(t, 到手价, 月售数, 库存数)]}}（时间升序），mtime 缓存
_hist_full_cache = {"mtime": None, "data": None}
_obs_meta_cache = {"mtime": None, "data": None}


def _load_hist_full():
    """完整价格观测历史（P1 记忆层数据源）：每次导入 = 一次全量快照，append-only。
    月售/库存做通用数字提取（'月售2'/'8人想买' 兼容），缺失为 None。"""
    global _hist_full_cache
    f = os.path.join(OUTPUT_DIR, "竞店标准库", "价格观测历史.jsonl")
    if not os.path.exists(f):
        return {}
    mt = os.path.getmtime(f)
    if _hist_full_cache["data"] is not None and mt == _hist_full_cache["mtime"]:
        return _hist_full_cache["data"]
    data = {}
    try:
        with open(f, encoding="utf-8") as fh:
            for ln in fh:
                try:
                    rec = json.loads(ln)
                except (ValueError, TypeError):
                    continue
                shop = _safe(rec.get("店"))
                sid = _safe(rec.get("id"))
                if not shop or not sid:
                    continue
                t = _safe(rec.get("t"))
                p = _extract_num(rec.get("到手价") or rec.get("折后价"))
                s = _extract_num(rec.get("月销量"))
                k = _extract_num(rec.get("库存"))
                data.setdefault(shop, {}).setdefault(sid, []).append((t, p, s, k))
    except Exception as e:
        print("[观测历史] 读取失败:", e)
        return _hist_full_cache["data"] or {}
    for shop in data:
        for sid in data[shop]:
            data[shop][sid].sort(key=lambda x: x[0])
    _hist_full_cache["mtime"] = mt
    _hist_full_cache["data"] = data
    return data


def _sku_history(shop, sku):
    """单个竞品 SKU 的观测序列（时间升序）；无观测 → []。"""
    return _load_hist_full().get(_safe(shop), {}).get(_safe(sku), [])


def _shop_obs_meta():
    """P1.4 各店观测概况 {店: {snaps: 快照次数, last_t: 最新观测时间}}（mtime 缓存）。"""
    global _obs_meta_cache
    f = os.path.join(OUTPUT_DIR, "竞店标准库", "价格观测历史.jsonl")
    if not os.path.exists(f):
        return {}
    mt = os.path.getmtime(f)
    if _obs_meta_cache["data"] is not None and _obs_meta_cache["mtime"] == mt:
        return _obs_meta_cache["data"]
    full = _load_hist_full()
    meta = {}
    for shop, skus in full.items():
        ts = {t for rows in skus.values() for (t, _p, _s, _k) in rows}
        meta[shop] = {"snaps": len(ts), "last_t": max(ts) if ts else ""}
    _obs_meta_cache["mtime"] = mt
    _obs_meta_cache["data"] = meta
    return meta


# P4.3 自家销量口径元数据（bind_own_sales 写入；UI 动态标注"销量是谁的销量"）
_own_sales_meta_cache = {"mtime": None, "data": None}


def _own_sales_meta():
    """读 output/own_sales_meta.json（bind_own_sales 每次回填时写入）。
    返回 {} 或 {snapshot, snapshot_date, stores[], store_codes[], rows, bind_time, note}。"""
    global _own_sales_meta_cache
    f = os.path.join(OUTPUT_DIR, "own_sales_meta.json")
    if not os.path.exists(f):
        return {}
    mt = os.path.getmtime(f)
    if _own_sales_meta_cache["data"] is not None and _own_sales_meta_cache["mtime"] == mt:
        return _own_sales_meta_cache["data"]
    try:
        with open(f, encoding="utf-8") as fh:
            data = json.load(fh)
    except Exception:
        data = {}
    _own_sales_meta_cache["mtime"] = mt
    _own_sales_meta_cache["data"] = data
    return data


def _own_sales_scope_text():
    """销量口径一句话（用于图例/chip title）：多仓合计：A+B / 单仓口径：A / 未回填提示。"""
    m = _own_sales_meta()
    if m and m.get("note"):
        d = m.get("snapshot_date", "")
        return f"补货参考快照 {m['note']}" + (f"（{d[:4]}-{d[4:6]}-{d[6:]}）" if len(d) == 8 else "")
    return "销量未回填（先运行 bind_own_sales 或等自动回填）"


# P4.3 店对店：分仓销量 {仓编码: {name, kw, sales:{自家SPUID: 销量}}}（仓的销量归仓；总库是关联）
_wh_sales_caches = {}  # {列名: {"mtime":, "data":}}——30天总销量/7天总销量 各一份


def _wh_sales(col="30天总销量"):
    """读 output/自家销量_分仓.xlsx → {仓编码: {name, kw, sales:{自家SPUID: 销量}}}。

    以「仓编码」为键——它才是销量口径的唯一身份；关键词 kw 只用于系统推定提示，
    绝不作为取值依据（同名仓、仓改名都会让关键词匹配悄悄拿错数）。
    """
    cache = _wh_sales_caches.setdefault(col, {"mtime": None, "data": None})
    f = os.path.join(OUTPUT_DIR, "自家销量_分仓.xlsx")
    if not os.path.exists(f):
        return {}
    mt = os.path.getmtime(f)
    if cache["data"] is not None and cache["mtime"] == mt:
        return cache["data"]
    out = {}
    try:
        df = _xread(f, dtype=str)
        if "门店/仓编码" in df.columns and "商品SKU" in df.columns:
            sku2spu = (dict(zip(own_df["SKUID"].astype(str).str.strip(),
                                own_df["SPUID"].astype(str).str.strip()))
                       if own_df is not None and len(own_df) else {})
            df["_spu"] = df["商品SKU"].astype(str).str.strip().map(sku2spu)
            df["_s"] = pd.to_numeric(df[col].apply(_extract_num), errors="coerce").fillna(0.0)
            sub = df[df["_spu"].notna() & (df["_spu"].astype(str).str.len() > 0)]
            for (code, name), g in sub.groupby(
                    ["门店/仓编码", "门店/仓名称"], dropna=False, sort=False):
                key = _safe(code) or _safe(name)
                if not key:
                    continue
                agg = g.groupby("_spu")["_s"].sum()
                d = out.setdefault(key, {"name": _safe(name),
                                         "kw": _store_wh.keyword_of(_safe(name)), "sales": {}})
                for spu, v in agg.items():
                    d["sales"][str(spu)] = d["sales"].get(str(spu), 0.0) + float(v)
    except Exception as e:
        print("[分仓销量] 读取失败:", e)
        return cache["data"] or {}
    cache["mtime"] = mt
    cache["data"] = out
    return out


def _wh_status(own_store):
    """该店的店对店销量口径状态：manual / pending / no_data / unmapped（含 usable 位）。"""
    return _store_wh.status(own_store, list(_wh_sales().keys()))


def _wh_map_for_store(own_store, col="30天总销量"):
    """自家店 → 其「已指认仓」的 {SPUID: 销量}；未指认或当期快照无该仓 → (None, "")。

    这里绝不回退多仓合计：合计会把别店（别的仓）的销量当成本店的，
    数字看着正常、结论是错的。要显示就得先在工具页把仓指认给这家店。
    """
    if not own_store:
        return None, ""
    st = _wh_status(own_store)
    if not st.get("usable"):
        return None, ""
    d = _wh_sales(col).get(str(st["code"])) or {}
    sales = d.get("sales") or {}
    return (sales if sales else None), st["kw"]


def _comp_cat(row):
    """竞品行类目（分位分组用）：三级归一 → 三级 → 二级 → 一级 回退取第一个非空。
    注意逐级 _safe：pandas 空单元格为 NaN（truthy），不能直接用 or 链短路。"""
    for k in ("类目三级_归一", "category3", "category2", "category1"):
        v = _safe(row.get(k))
        if v:
            return v
    return ""


def _build_cat_prices():
    """全店竞品 SKU 到手价按类目分组排序（品类内价格分位的分布）。"""
    global _cat_price_cache
    dist = {}
    for s, lk in competitor_lookup.items():
        for sku, row in lk.items():
            cat = _comp_cat(row)
            if not cat:
                continue
            p = _extract_num(row.get("price") or row.get("折后价"))
            if p is None:
                continue
            dist.setdefault(cat, []).append(p)
    _cat_price_cache = {k: sorted(v) for k, v in dist.items()}


def _price_percentile(cat, price):
    """价格在同类内的分位 0-100（越高越贵）；样本<8 或数据缺失 → None。"""
    if not _cat_price_cache or not cat:
        return None
    arr = _cat_price_cache.get(cat)
    if not arr or len(arr) < 8:
        return None
    p = _extract_num(price)
    if p is None:
        return None
    import bisect as _bs
    i = _bs.bisect_left(arr, p)
    return round(i / len(arr) * 100)


def _sku_of(idx):
    """合并池第 idx 行的竞品Sku_id（列向量索引，快）。"""
    return _pool_sku_list[idx] if idx < len(_pool_sku_list) else _safe(pool_df.iloc[idx].get("竞品Sku_id"))


def _is_confirmed(idx):
    shop, sku = _shop_of(idx), _sku_of(idx)
    return sku in progress.get("confirmed", {}).get(shop, {})


def _is_no_match(idx):
    shop, sku = _shop_of(idx), _sku_of(idx)
    return sku in progress.get("no_match", {}).get(shop, {})


def _mark_confirmed(idx, sku_id):
    shop = _shop_of(idx)
    progress.setdefault("confirmed", {}).setdefault(shop, {})[_sku_of(idx)] = {
        "sku_id": sku_id,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    progress.setdefault("no_match", {}).get(shop, {}).pop(_sku_of(idx), None)


def _mark_no_match(idx):
    shop = _shop_of(idx)
    progress.setdefault("no_match", {}).setdefault(shop, {})[_sku_of(idx)] = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    progress.setdefault("confirmed", {}).get(shop, {}).pop(_sku_of(idx), None)


def load_progress():
    """从 progress.json 读取进度（迁移旧行号结构 -> 按店+竞品Sku_id），并同步到 pool_df。"""
    global progress
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                progress = json.load(f)
        except Exception:
            progress = {"confirmed": {}, "no_match": {}}
    else:
        progress = {"confirmed": {}, "no_match": {}}

    # 迁移旧结构：confirmed/no_match 的键若是纯数字行号 -> (默认店, 竞品Sku_id)
    # 旧行号对应旧单店池文件（output/对齐_人工确认池.xlsx），不能用新合并池行号反查（行序不同）
    if pool_df is not None and len(pool_df):
        legacy_pool = None
        if os.path.exists(CONFIRM_POOL_FILE):
            try:
                legacy_pool = pd.read_excel(CONFIRM_POOL_FILE, dtype=str)
            except Exception:
                legacy_pool = None
        migrated = {"confirmed": {}, "no_match": {}}
        for kind in ("confirmed", "no_match"):
            for k, v in progress.get(kind, {}).items():
                if str(k).isdigit():
                    idx = int(k)
                    sku, shop = "", DEFAULT_SHOP
                    # 行号范围（旧池最多几百行）内 -> 从旧单店池反查 Sku_id
                    if legacy_pool is not None and 0 <= idx < len(legacy_pool):
                        sku = _safe(legacy_pool.iloc[idx].get("竞品Sku_id"))
                    elif 0 <= idx < len(pool_df) and idx < 10000:
                        sku = _sku_of(idx)
                    else:
                        # 超出行号范围的数字键：本身即竞品Sku_id（历史版本用Sku_id作键）
                        sku = str(k)
                    if sku:
                        migrated[kind].setdefault(shop, {})[sku] = v
                else:
                    migrated[kind][k] = v
        progress = migrated

        # 同步到内存 pool_df 的确认结果列
        if "确认结果" in pool_df.columns:
            for shop, m in progress.get("confirmed", {}).items():
                for sku in m:
                    pool_df.loc[(pool_df["竞店名称"] == shop) &
                                (pool_df["竞品Sku_id"] == sku), "确认结果"] = "确认"
            for shop, m in progress.get("no_match", {}).items():
                for sku in m:
                    pool_df.loc[(pool_df["竞店名称"] == shop) &
                                (pool_df["竞品Sku_id"] == sku), "确认结果"] = "无对应"


def save_progress():
    with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
        json.dump(progress, f, ensure_ascii=False, indent=2)
    _invalidate_stats()


_stats_cache = None  # {店: {total,pending,coarse,aligned,nomatch}} 缓存；进度变更时失效
_compare_cache = {}  # 同品对比/产品分析缓存 {(shop, view): (df, agg_df)}；数据变更时失效
_top50_advice_cache = {}  # {shop: [{id,品名,状态,竞品月售,自家品名,建议}]}；AI建议结果缓存
# dashboard 桥接映射缓存（跟 _compare_cache 同失效）：避免每次请求遍历 own_df(4.4万行)
_bridge_cache = {}  # {shop: (own_sku_map, spu_own, sku2own, concl_map)}
# 参考桥接缓存（P2.1 显示层前置）：候选池高置信品级候选 → 参考结论（未经人工确认，明确标注）
_ref_bridge_cache = {}  # {shop: {竞品品ID: (自家品组ID, 自家品名, 置信度)}}
SEARCH_LOG_FILE = os.path.join(OUTPUT_DIR, "搜索学习日志.jsonl")  # 搜索点击样本（词→品组人工标注）


def _log_search_click(entry, q, spu, name, results_count=0):
    """记录搜索点击样本：搜索词→点选品组（人工标注的切词样本，供关键词库反哺）。
    噪音控制：只记有点选（spu 非空）的强标注；泛词/低频由提炼脚本按 频次×指向品组数 过滤。"""
    try:
        from search_util import split_keywords
        words = split_keywords(q)
        if not words or not spu:
            return
        import json as _json
        os.makedirs(OUTPUT_DIR, exist_ok=True)
        with open(SEARCH_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(_json.dumps({
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "entry": entry, "q": q, "words": words,
                "spu": spu, "name": (name or "")[:80], "results": results_count,
            }, ensure_ascii=False) + "\n")
    except Exception:
        pass
_pool_shop_list = []  # pool_df 竞店名称 列向量（替代 iloc，快百倍）
_pool_sku_list = []   # pool_df 竞品Sku_id 列向量
_pool_value_cache = None  # 池行价值（月售,价格）惰性缓存；仅依赖竞品数据，load_data 重建
_pool_conf_cache = None   # 池行置信度惰性缓存（按置信度排序用）；load_data 重建
_pool_status_cache = None  # 池行状态惰性缓存；进度变更失效
_own_price_cache = None   # 自家品组价格带 {SPUID: (min售价, max售价)}；load_data 重建（比价展示用）
_price_hist_cache = None  # {(店, 竞品sku): (t, 上次到手价)} 最新价格观测；jsonl mtime 失效
_price_hist_mtime = -1.0
_cat_price_cache = None   # {类目: sorted[到手价]} 全店 SKU 价格分布；load_data 重建（品类内分位用）
_lib_cache = None      # 汇总库两层缓存；汇总库写入后失效
_level_cache = {}      # {(shop, sku): 收敛层级} 惰性缓存；进度变更失效
_pool_shop_skus = None  # {店: set(竞品Sku_id)} 人工池按店索引；pool_df 重建时失效
_warmup_lock = threading.Lock()   # 预热守卫锁
_warmup_state = {}     # {店: (state, ts)} state: running/ok/failed；failed 600s 退避防线程堆积

# ---- AI 预筛（优化项A）：结果徽章 + 后台批量任务状态 ----
_prescreen_cache = None       # {(店, 竞品Sku_id): rec} 惰性缓存；按文件 mtime 失效
_prescreen_mtime = 0.0
_PRESCREEN_STATE = {"running": False, "shop": "", "done": 0, "total": 0,
                    "ok": 0, "fail": 0, "last_err": "", "last_stats": None}
_PRESCREEN_LOCK = threading.Lock()

# ---- M7 阶段4：通用后台任务（备份/对齐重跑/销量回填）一次只跑一个 ----
_BG_JOB = {"running": False, "kind": "", "target": "", "started": "",
           "finished": "", "ok": None, "msg": "", "detail": ""}
_BG_LOCK = threading.Lock()


def _bg_start(kind, target):
    with _BG_LOCK:
        if _BG_JOB["running"]:
            return False, f"已有后台任务在跑：{_BG_JOB['kind']} {_BG_JOB['target']}"
        _BG_JOB.update(running=True, kind=kind, target=target,
                       started=time.strftime("%Y-%m-%d %H:%M:%S"),
                       finished="", ok=None, msg="启动中…", detail="")
        return True, ""


def _bg_finish(ok, msg, detail=""):
    _BG_JOB.update(running=False, ok=bool(ok), msg=str(msg)[:300],
                   detail=str(detail)[:500], finished=time.strftime("%Y-%m-%d %H:%M:%S"))


def _bg_snapshot():
    d = {k: _BG_JOB.get(k) for k in
         ("running", "kind", "target", "started", "finished", "ok", "msg", "detail")}
    d["prescreen"] = {k: _PRESCREEN_STATE.get(k) for k in
                      ("running", "shop", "done", "total", "ok", "fail")}
    return d


def _run_backup_bg():
    try:
        import backup_daily as _bd
        dest, n, mb, err = _bd.make_backup()
        if dest:
            _bg_finish(True, f"备份完成 {os.path.basename(dest)}（{n} 文件 {mb:.1f}MB）",
                       dest)
            try:
                load_data()
                _invalidate_stats()
            except Exception:
                pass
        else:
            _bg_finish(False, f"备份失败：{err or '未知'}", err)
    except Exception as e:
        _bg_finish(False, f"备份异常：{e}", str(e))


def _run_align_bg(shop):
    try:
        from build_comp_library import rerun_align as _ra
        ok = _ra(shop)
        if ok:
            try:
                load_data()
                _invalidate_stats()
            except Exception as e:
                _bg_finish(True, f"对齐重跑完成，但面板重载失败：{e}（请重启面板）", shop)
                return
            _bg_finish(True, f"{shop} 对齐重跑完成 ✓（归一化+对齐+汇总库已刷新）", shop)
        else:
            _bg_finish(False, f"{shop} 对齐重跑失败（见日志）", shop)
    except Exception as e:
        _bg_finish(False, f"对齐重跑异常：{e}", str(e))


def _run_backfill_bg(snap_file=""):
    """销量回填：子进程跑 bind_own_sales（独立进程写自家表，避免与面板锁死锁）。"""
    import subprocess
    try:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        cmd = [sys.executable, os.path.join(root, "scripts", "bind_own_sales.py")]
        if snap_file:
            cmd += ["--file", snap_file]
        r = subprocess.run(cmd, capture_output=True, text=True,
                           encoding="utf-8", errors="replace", cwd=root, timeout=600)
        tail = ((r.stdout or "") + "\n" + (r.stderr or ""))[-400:]
        if r.returncode == 0:
            try:
                load_data()
                _invalidate_stats()
            except Exception:
                pass
            _bg_finish(True, "销量回填完成 ✓（自家表已更新，请刷新页面看新口径）", tail)
        else:
            _bg_finish(False, f"销量回填失败 rc={r.returncode}", tail)
    except Exception as e:
        _bg_finish(False, f"销量回填异常：{e}", str(e))


def _prescreen_map():
    """AI 预筛结论 {(店, 竞品Sku_id): rec}；文件 mtime 变化才重读。"""
    global _prescreen_cache, _prescreen_mtime
    try:
        f = os.path.join(OUTPUT_DIR, "ai_prescreen.json")
        mt = os.path.getmtime(f) if os.path.exists(f) else 0.0
        if _prescreen_cache is None or mt != _prescreen_mtime:
            from ai_prescreen import load_results
            data = load_results()
            _prescreen_cache = {(shop, str(sku)): rec
                                for shop, m in (data or {}).items()
                                for sku, rec in (m or {}).items()}
            _prescreen_mtime = mt
        return _prescreen_cache
    except Exception:
        return {}


def _run_prescreen_bg(shop, limit):
    """后台线程：批量预筛（写 output/ai_prescreen.json，徽章随 mtime 自动刷新）。
    triage 模式（Phase2）：预筛完自动分流落地——AI预筛行写汇总库、高置信否标记无对应、
    导出全量明细 Excel；无对应标记走面板内存 progress（save_progress 落盘+失效缓存）。"""
    with _PRESCREEN_LOCK:
        if _PRESCREEN_STATE["running"]:
            return
        _PRESCREEN_STATE.update(running=True, shop=shop, done=0, total=0,
                                ok=0, fail=0, last_err="")
    try:
        from ai_prescreen import prescreen_shop, apply_triage, export_report

        def _cb(done, total, ok, fail):
            _PRESCREEN_STATE.update(done=done, total=total, ok=ok, fail=fail)

        st = prescreen_shop(shop, limit=limit, progress_cb=_cb)
        _PRESCREEN_STATE["last_stats"] = st

        # Phase2：triage 分流落地（mode=annotate 时 apply_triage 返回 skipped 不动作）
        def _nm_writer(sku, note):
            progress.setdefault("no_match", {}).setdefault(shop, {})[sku] = {
                "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "note": note, "source": "AI预筛",
            }
            progress.setdefault("confirmed", {}).get(shop, {}).pop(sku, None)

        tri = apply_triage(shop, no_match_writer=_nm_writer)
        _PRESCREEN_STATE["last_triage"] = tri
        if not tri.get("skipped"):
            save_progress()   # 落盘 progress + 失效统计/汇总库缓存
        n, path = export_report(shop)
        _PRESCREEN_STATE["last_report"] = path
    except Exception as e:
        _PRESCREEN_STATE["last_err"] = str(e)[:200]
        _logger.exception("AI预筛后台任务失败")
    finally:
        _PRESCREEN_STATE["running"] = False



def _advice_view(shop, rows, limit=50):
    """AI 建议行 + 跳转链接。

    跳转优先按代表 SKU 直达竞品详情页：以前是拿整串电商标题去搜索（`?q=标题&filter=opp`），
    实测长标题分词后基本搜不到、而且建议里的品大多不是"机会品"会被 opp 筛光 → 一点就空页。
    """
    out = []
    for a in (rows or [])[:limit]:
        aa = dict(a)
        sku = _safe(a.get("代表Sku_id"))
        if sku:
            aa["_url"] = "/comp_detail?shop=" + quote(shop or "") + "&sku=" + quote(sku)
            aa["_url_tip"] = "看该品详情与自家对照"
        else:
            aa["_url"] = ("/?shop=" + quote(shop or "") + "&q="
                          + quote(_safe(a.get("品名"))))
            aa["_url_tip"] = "在榜单中搜该品"
        out.append(aa)
    return out


def _invalidate_stats():
    global _stats_cache, _pool_status_cache, _lib_cache, _compare_cache, _top50_advice_cache, _level_cache, _pool_shop_skus, _bridge_cache, _ref_bridge_cache
    _stats_cache = None
    _pool_status_cache = None
    _lib_cache = None
    _compare_cache = {}
    _top50_advice_cache = {}
    _level_cache = {}
    _pool_shop_skus = None
    _bridge_cache = {}
    _ref_bridge_cache = {}
    # 竞品库行构建结果 / 店铺统计 / 滞销榜：与上面同源同生命周期，人工确认后状态必须变
    _browse_rows_cache.clear()
    _browse_meta_cache.clear()
    _slow_rank_cache.clear()
    _zone_rank_cache.clear()


def _ref_bridge(shop):
    """参考对应桥接（P2.1 显示层前置 → M6 扩展）：候选池三层候选并入参考映射，
    按竞品品ID → (自家品组ID, 自家品名, 置信度)。

    数据源与阈值（宁缺毋滥）：
      - 品级语义层（第2.5层）置信 ≥0.55（0.8/0.7 类目内强语义；0.55=类目缺失店唯一锚点）
      - 向量候选层（第3.6层）相似度 ≥0.62（FT3 校准门槛；未调參的灰区不进徽章）
      - 关键字候选层（第3.5层）置信 ≥0.62
    价值：**未经人工确认的对应也能在榜单/竞品库直接看到「参考·置信X」紫徽章**——
    小模型的产出不需要等人工确认才可见；确认链（progress/汇总库）永远优先。
    """
    if shop in _ref_bridge_cache:
        return _ref_bridge_cache[shop]
    out = {}
    sources = [
        (os.path.join(OUTPUT_DIR, f"试跑_{shop}", "对齐_第2.5层_品级语义_人工池.xlsx"), "置信度", 0.55),
        (os.path.join(OUTPUT_DIR, f"试跑_{shop}", "对齐_第3.6层_向量候选_人工池.xlsx"), "相似度", 0.62),
        (os.path.join(OUTPUT_DIR, f"试跑_{shop}", "对齐_第3.5层_关键字候选_人工池.xlsx"), "置信度", 0.62),
    ]
    for f, conf_col, min_conf in sources:
        if not os.path.exists(f):
            continue
        try:
            df = pd.read_excel(f, dtype=str)
            if not len(df) or not {"竞品品ID", "自家SPUID"} <= set(df.columns):
                continue
            for cid, grp in df.groupby("竞品品ID"):
                best = None
                for _, r in grp.iterrows():
                    gid = str(r.get("自家SPUID") or "").strip()
                    if not gid or gid.lower() == "nan":
                        continue
                    try:
                        conf = float(str(r.get(conf_col) or 0))
                    except (TypeError, ValueError):
                        conf = 0.0
                    if conf < min_conf:
                        continue
                    if best is None or conf > best[2]:
                        best = (gid, str(r.get("自家商品名") or "").strip(), conf)
                if best and (str(cid) not in out or out[str(cid)][2] < best[2]):
                    out[str(cid)] = best
        except Exception as e:
            print("[参考桥接] 读取失败:", shop, os.path.basename(f), e)
    _ref_bridge_cache[shop] = out
    return out


# P4.3 own 数据新鲜度：bind_own_sales/bind_own_images 是独立进程，重写 own 表后
# 运行中的面板必须自动感知（重载 own 数据 + 失效分析缓存），否则一直用启动时的旧数据
_own_mtime_seen = {"mtime": None}


def _reload_own_data():
    """重载 own 相关全局（own_df/own_lookup/品组价格带/类目品牌下拉）。"""
    global own_df, own_lookup, own_categories, own_brands, _own_price_cache
    own_df = _xread(OWN_FILE, dtype=str)
    for col in ["SPUID", "SKUID"]:
        if col in own_df.columns:
            own_df[col] = own_df[col].apply(_safe)
    own_lookup = {}
    for _, row in own_df.iterrows():
        sku = _safe(row.get("SKUID"))
        if sku:
            own_lookup[sku] = row
    _own_price_cache = {}
    try:
        if len(own_df) and "SPUID" in own_df.columns and "美团闪购渠道售价" in own_df.columns:
            _t = own_df[["SPUID", "美团闪购渠道售价"]].copy()
            _t["SPUID"] = _t["SPUID"].astype(str).str.strip()
            _t["_p"] = pd.to_numeric(_t["美团闪购渠道售价"].astype(str)
                                     .str.replace(r"[^\d.]", "", regex=True)
                                     .replace("", None), errors="coerce")
            _t = _t.dropna(subset=["_p"])
            _t = _t[(_t["SPUID"].str.len() > 0) & (_t["SPUID"] != "nan")]
            if len(_t):
                _g = _t.groupby("SPUID")["_p"].agg(["min", "max"])
                _own_price_cache = {k: (round(r["min"], 2), round(r["max"], 2))
                                    for k, r in _g.iterrows()}
    except Exception as _e:
        print("[reload_own] 品组价格带构建失败:", _e)
    cat_col = "类目三级_归一" if "类目三级_归一" in own_df.columns else "美团闪购类目三级"
    own_categories = sorted(
        [c for c in own_df[cat_col].dropna().unique().tolist() if _safe(c)]
    ) if cat_col in own_df.columns else []
    brand_col = "品牌（抽取）" if "品牌（抽取）" in own_df.columns else "品牌名称"
    own_brands = sorted(
        [b for b in own_df[brand_col].dropna().unique().tolist() if _safe(b)]
    ) if brand_col in own_df.columns else []


def _check_own_data_fresh():
    """own 表 mtime 变化 → 重载 own 数据 + 失效全部分析缓存（before_request 每请求 ~1ms）。"""
    try:
        mt = os.path.getmtime(OWN_FILE)
    except OSError:
        return
    seen = _own_mtime_seen.get("mtime")
    if seen is None:
        _own_mtime_seen["mtime"] = mt
        return
    if mt != seen:
        _own_mtime_seen["mtime"] = mt
        print("[own] 检测到 own 表被外部重写，重载并失效缓存")
        try:
            _reload_own_data()
        except Exception as e:
            print("[own] 重载失败:", e)
        _invalidate_stats()


def _build_pool_shop_skus():
    """人工池按店索引：{店: set(竞品Sku_id)}，一次构建，pool_df 重建时失效。"""
    global _pool_shop_skus
    if _pool_shop_skus is not None:
        return _pool_shop_skus
    m = {}
    if pool_df is not None and len(pool_df):
        for s, grp in pool_df.groupby("竞店名称")["竞品Sku_id"]:
            m[_safe(s)] = set(grp.astype(str))
    _pool_shop_skus = m
    return m


def _get_stats_all():
    """一次遍历构建全部店统计明细（缓存）。返回 {店: {5项}}。"""
    global _stats_cache
    if _stats_cache is None:
        per = {}
        for s, lk in competitor_lookup.items():
            conf_map = progress.get("confirmed", {}).get(s, {})
            nm_set = set(progress.get("no_match", {}).get(s, {}).keys())
            a = sum(1 for sku in lk if sku in aligned_map
                    or (sku in conf_map and not conf_map.get(sku, {}).get("spu_only")))
            c = 0
            for sku, row in lk.items():
                if sku in conf_map and conf_map.get(sku, {}).get("spu_only"):
                    c += 1
                elif _converge_level(s, sku, row) == "spu":
                    c += 1
            n = len(nm_set)
            per[s] = {"total": len(lk), "aligned": a, "coarse": c, "nomatch": n,
                      "pending": len(lk) - a - c - n}
        _stats_cache = per
    return _stats_cache


def _get_stats(shop=""):
    """当前范围（全部店或单店）5 项统计（缓存）。"""
    per = _get_stats_all()
    if shop:
        return per.get(shop, {"total": 0, "pending": 0, "coarse": 0, "aligned": 0, "nomatch": 0})
    st = {"total": 0, "pending": 0, "coarse": 0, "aligned": 0, "nomatch": 0}
    for v in per.values():
        for k in st:
            st[k] += v[k]
    return st


def get_competitor_info(sku_id, shop=None):
    """从指定竞店全表获取图片url/价格/月售等补充信息（兼容中英文列名）。"""
    info = {}
    shop = shop or DEFAULT_SHOP
    row = competitor_lookup.get(shop, {}).get(_safe(sku_id))
    if row is not None:
        # 新版归一化表英文列：image_url/price/original_price/sales/spec/category1..3/shop_name
        info["图片url"] = _safe(row.get("图片链接") or row.get("图片url") or row.get("image_url"))
        info["价格"] = _safe(row.get("到手价") or row.get("折后价") or row.get("price"))
        info["原价"] = _safe(row.get("原价") or row.get("original_price"))
        info["月售"] = _safe(row.get("月销量") or row.get("月售") or row.get("sales"))
        info["规格"] = _safe(row.get("规格") or row.get("spec"))
        info["一级类"] = _safe(row.get("一级类目") or row.get("一级类") or row.get("category1"))
        info["二级类"] = _safe(row.get("二级类目") or row.get("二级类") or row.get("category2"))
        info["三级类"] = _safe(row.get("三级类目") or row.get("美团后台三级类") or row.get("category3"))
        info["店名"] = _safe(row.get("shop_name"))
    return info


def get_own_info(sku_id):
    """从自家全表获取商品完整信息。"""
    info = {}
    row = own_lookup.get(_safe(sku_id))
    if row is not None:
        info["SPUID"] = _safe(row.get("SPUID"))
        info["SKUID"] = _safe(row.get("SKUID"))
        info["商品名称"] = _safe(row.get("商品名称"))
        info["规格名称"] = _safe(row.get("规格名称"))
        info["品牌名称"] = _safe(row.get("品牌名称"))
        info["品牌抽取"] = _safe(row.get("品牌（抽取）"))
        info["类目一级"] = _safe(row.get("美团闪购类目一级"))
        info["类目二级"] = _safe(row.get("美团闪购类目二级"))
        info["类目三级"] = _safe(row.get("美团闪购类目三级"))
        info["类目归一"] = _safe(row.get("类目三级_归一"))
        info["售价"] = _safe(row.get("美团闪购渠道售价"))
        info["采购价"] = _safe(row.get("采购价"))
        info["月销量"] = _safe(row.get("月销量"))
        info["7天总销量"] = _safe(row.get("7天总销量"))
        info["14天总销量"] = _safe(row.get("14天总销量"))
        info["预测日均销量"] = _safe(row.get("预测日均销量"))
        info["可用库存"] = _safe(row.get("可用库存"))
        info["是否有图片"] = _safe(row.get("是否有图片"))
        info["商品图片url"] = _safe(row.get("商品图片url"))
        info["型号"] = _safe(row.get("型号"))
        info["商品条码"] = _safe(row.get("商品条码"))
    return info


def get_pool_item(idx):
    """获取候选池第 idx 行的完整信息（竞品+候选自家+品级候选）。"""
    if idx < 0 or idx >= len(pool_df):
        return None
    row = pool_df.iloc[idx]
    shop = _shop_of(idx)
    item = {
        "idx": idx,
        "竞店名称": shop,
        "竞品品ID": _safe(row.get("竞品品ID")),
        "竞品Sku_id": _safe(row.get("竞品Sku_id")),
        "竞品商品名": _safe(row.get("竞品商品名")),
        "竞品条形码": _safe(row.get("竞品条形码")),
        "竞品品牌": _safe(row.get("竞品品牌")),
        "竞品型号": _safe(row.get("竞品型号")),
        "竞品类目": _safe(row.get("竞品类目")),
        "竞品规格键": _safe(row.get("竞品规格键")),
        "自家SPUID": _safe(row.get("自家SPUID")),
        "自家SKUID": _safe(row.get("自家SKUID")),
        "自家商品名": _safe(row.get("自家商品名")),
        "自家商品条码": _safe(row.get("自家商品条码")),
        "自家品牌": _safe(row.get("自家品牌")),
        "自家型号": _safe(row.get("自家型号")),
        "自家类目": _safe(row.get("自家类目")),
        "自家规格键": _safe(row.get("自家规格键")),
        "匹配层": _safe(row.get("匹配层")),
        "匹配依据": _safe(row.get("匹配依据")),
        "置信度": _safe(row.get("置信度")),
        "备注": _safe(row.get("备注")),
        "建议动作": _safe(row.get("建议动作")),
        "确认结果": _safe(row.get("确认结果")),
        "人工备注": _safe(row.get("人工备注")),
        "品级置信度": _safe(row.get("品级置信度")),
    }
    # 补充竞品图片/价格/月售（按店查询）
    comp_info = get_competitor_info(item["竞品Sku_id"], shop)
    item["竞品图片url"] = comp_info.get("图片url", "")
    item["竞品价格"] = comp_info.get("价格", "")
    item["竞品原价"] = comp_info.get("原价", "")
    item["竞品月售"] = comp_info.get("月售", "")
    item["竞品规格"] = comp_info.get("规格", "")
    # 候选自家的补充信息
    own_info = get_own_info(item["自家SKUID"])
    item["候选自家售价"] = own_info.get("售价", "")
    item["候选自有图"] = own_info.get("是否有图片", "")
    # 品级候选：解析 候选自家品 JSON -> 候选品列表（含组内SKU明细）
    item["候选品列表"] = parse_spu_candidates(_safe(row.get("候选自家品")), item["自家SKUID"])
    return item


def parse_spu_candidates(raw_json, fallback_own_sku=""):
    """解析人工池 候选自家品 JSON，返回候选品列表；
    每条含 品组ID/品名/品牌/型号/规格数/置信度/SKU列表（明细补售价/是否有图）。"""
    cands = []
    if raw_json and raw_json != "":
        try:
            data = json.loads(raw_json)
        except (ValueError, TypeError):
            data = []
        for c in (data or []):
            if not isinstance(c, dict):
                continue
            skus = []
            for sku in c.get("SKU列表") or []:
                oi = get_own_info(sku)
                skus.append({
                    "SKUID": _safe(sku),
                    "商品名称": oi.get("商品名称", ""),
                    "规格名称": oi.get("规格名称", ""),
                    "售价": oi.get("售价", ""),
                    "是否有图片": oi.get("是否有图片", ""),
                })
            cands.append({
                "品组ID": _safe(c.get("品组ID")),
                "品名": _safe(c.get("品名")),
                "品牌": _safe(c.get("品牌")),
                "型号": _safe(c.get("型号")),
                "规格数": _safe(c.get("规格数")),
                "置信度": _safe(c.get("置信度")),
                "SKU列表": skus,
            })
    # 无品级候选时，若已有候选自家SKU，作为单条候选（兼容旧链路）
    if not cands and fallback_own_sku:
        oi = get_own_info(fallback_own_sku)
        cands.append({
            "品组ID": "", "品名": oi.get("商品名称", ""), "品牌": "", "型号": "",
            "规格数": 1, "置信度": "",
            "SKU列表": [{"SKUID": fallback_own_sku, "商品名称": oi.get("商品名称", ""),
                          "规格名称": oi.get("规格名称", ""), "售价": oi.get("售价", ""),
                          "是否有图片": oi.get("是否有图片", "")}],
        })
    return cands


def append_to_excel(filepath, new_row_dict, columns=None):
    """向 Excel 追加一行（读取+追加+写回），合并列确保不丢列。"""
    if os.path.exists(filepath):
        existing = pd.read_excel(filepath)
        # 合并列：existing 的列 + new_row_dict 中多出的列
        all_cols = list(existing.columns) + [c for c in new_row_dict.keys() if c not in existing.columns]
        for c in all_cols:
            if c not in existing.columns:
                existing[c] = ""
        row_df = pd.DataFrame([{col: new_row_dict.get(col, "") for col in all_cols}])
        combined = pd.concat([existing[all_cols], row_df], ignore_index=True)
    else:
        cols = columns or list(new_row_dict.keys())
        combined = pd.DataFrame([{col: new_row_dict.get(col, "") for col in cols}])
    combined.to_excel(filepath, index=False)


def save_pool_excel():
    """按店把内存中 pool_df（含更新的确认结果）写回各店人工池文件。"""
    try:
        if pool_df is None or not len(pool_df):
            return
        for shop, g in pool_df.groupby("竞店名称", sort=False):
            trial_f = os.path.join(OUTPUT_DIR, f"试跑_{shop}", "对齐_人工确认池.xlsx")
            target = trial_f if os.path.exists(trial_f) else CONFIRM_POOL_FILE
            out = g.drop(columns=[c for c in ["竞店名称", "__rowid"] if c in g.columns])
            if "__rowid" in g.columns:
                out = out.assign(__rowid=g["__rowid"].values)
                out = out.sort_values("__rowid").drop(columns=["__rowid"])
            out.to_excel(target, index=False)
    except PermissionError:
        # 文件被 Excel 占用，忽略（进度已存 progress.json）
        pass


# ---------------------------------------------------------------------------
# 路由
# ---------------------------------------------------------------------------
def _build_bridge(shop):
    """构建首页桥接映射（向量化）：
    sku2own  对齐明细 竞品Sku_id→自家SKUID
    own_sku_map own_df SKUID→(SPUID,品名,月销量,售价)
    spu_own  汇总库SPU层 竞品品ID→(自家品组ID,品名)
    concl_map 产品分析 品组ID→(结论,自家月销)；产品分析缺→None(未就绪)
    依赖 (shop,"detail")/(shop,"product") 缓存——未命中对应映射降级为空（不阻塞渲染）。"""
    global _bridge_cache
    detail_d = None
    key_d = (shop, "detail")
    if key_d in _compare_cache:
        detail_d, _ = _compare_cache[key_d]
    # 向量化构建 sku2own（对齐明细 竞品Sku_id→自家SKUID）
    sku2own = {}
    if detail_d is not None and len(detail_d):
        _cs = detail_d["竞品Sku_id"].astype(str).str.strip() if "竞品Sku_id" in detail_d.columns else pd.Series(dtype=str)
        _os = detail_d["自家SKUID"].astype(str).str.strip() if "自家SKUID" in detail_d.columns else pd.Series(dtype=str)
        mask = (_cs.str.len() > 0) & (_os.str.len() > 0) & (_cs != "nan") & (_os != "nan")
        for cs, os_ in zip(_cs[mask], _os[mask]):
            sku2own[cs] = os_
    # 向量化构建 own_sku_map（own_df SKUID→(SPUID,品名,月销量,售价)）
    own_sku_map = {}
    if own_df is not None and len(own_df):
        _od = own_df
        _sk = _od["SKUID"].astype(str).str.strip() if "SKUID" in _od.columns else pd.Series(dtype=str, index=_od.index)
        _sp = _od["SPUID"].astype(str).str.strip() if "SPUID" in _od.columns else pd.Series(dtype=str, index=_od.index)
        _nm = _od["商品名称"] if "商品名称" in _od.columns else pd.Series(dtype=str, index=_od.index)
        _sl = _od["月销量"] if "月销量" in _od.columns else pd.Series(dtype=object, index=_od.index)
        _pr = _od["美团闪购渠道售价"] if "美团闪购渠道售价" in _od.columns else pd.Series(dtype=object, index=_od.index)
        mask = (_sk.str.len() > 0) & (_sk != "nan")
        for i in _sk[mask].index:
            own_sku_map[_sk[i]] = (_sp[i], _nm[i], _sl[i], _pr[i])
    # 向量化构建 spu_own（汇总库 SPU 层 竞品品ID→(自家品组ID,品名)）
    spu_own = {}
    try:
        spu_lib, _ = _load_library_cached()
        if len(spu_lib):
            m = spu_lib[spu_lib["状态"].astype(str) == "有效"]
            if len(m):
                m = m[m["竞店名称"] == shop]
                if len(m):
                    _pid = m["竞品品ID"].astype(str).str.strip() if "竞品品ID" in m.columns else pd.Series(dtype=str, index=m.index)
                    _gid = m["自家品组ID"].astype(str).str.strip() if "自家品组ID" in m.columns else pd.Series(dtype=str, index=m.index)
                    _gn = m["自家品名"] if "自家品名" in m.columns else pd.Series(dtype=str, index=m.index)
                    mask = (_pid.str.len() > 0) & (_pid != "nan")
                    for i in _pid[mask].index:
                        spu_own[_pid[i]] = (_gid[i], _gn[i])
    except Exception:
        pass
    # 向量化构建 concl_map（产品分析 品组ID→(结论,自家月销)）
    concl_map = None
    try:
        if (shop, "product") in _compare_cache:
            pdf, _ = _compare_cache[(shop, "product")]
            if pdf is not None and len(pdf):
                concl_map = {}
                _gid2 = pdf["品组ID"].astype(str).str.strip() if "品组ID" in pdf.columns else pd.Series(dtype=str, index=pdf.index)
                _con = pdf["结论"] if "结论" in pdf.columns else pd.Series(dtype=str, index=pdf.index)
                _osl = pdf["自家月销"] if "自家月销" in pdf.columns else pd.Series(dtype=object, index=pdf.index)
                mask = (_gid2.str.len() > 0) & (_gid2 != "nan")
                for i in _gid2[mask].index:
                    concl_map[_gid2[i]] = (_con[i], _osl[i])
    except Exception:
        concl_map = concl_map or {}
    if concl_map is None:
        concl_map = {}
    return own_sku_map, spu_own, sku2own, concl_map


def _warmup_all_shops_bg():
    """后台单线程顺序预热全部店（避免 GIL 争抢与内存尖峰）。
    数据导入/改名/删店后调用：_invalidate_stats 清的是全部店的缓存，
    只预热变更店会让其他店首次访问同步重算（top500 2-7s/店）。"""
    def _run():
        for s in list(competitor_lookup.keys()):
            _warmup_one_shop(s)
    threading.Thread(target=_run, daemon=True).start()


# ---------------------------------------------------------------------------
# P0.1 全站范围选择条（docs/总任务规划.md）：竞店侧=全部/区/单店；自家侧=总库/门店
# ---------------------------------------------------------------------------
def build_scope_bar(base, comp_active="", own_active="", keep=(),
                    zones=None, free_shops=None, own_stores=None, own_carry_qs="",
                    show_pick=True):
    """构建全站范围选择条数据。

    base        : 页面路径（如 "/"、"/browse"）
    comp_active : 竞店侧当前选中值（"__all__"/"__pick"/区名/店名）
    own_active  : 自家侧当前选中店名（""=总库）
    keep        : 切换范围时透传的 URL 参数名（filter/q/status/...）；page 强制回 1
    own_carry_qs: 自家选项需要携带的竞店范围查询串（如 "&zone=测试区&inc=x"）——
                  修复"切自家店竞店跳回/切竞店自家店跳回"的范围联动丢失问题。

    返回 {"comp": [{label,url,active,title}], "own": [...]}，供 _scope.html 渲染。
    竞店选项切换时自动携带当前自家店（own_store）；自家选项切换时携带 own_carry_qs。
    """
    from urllib.parse import urlencode

    base_args = {k: v for k, v in request.args.items() if k in keep and v}

    def _add_qs(url, qs):
        if not qs:
            return url
        return url + ("&" if "?" in url else "?") + qs.lstrip("&")

    def _url(**over):
        params = dict(base_args)
        params.pop("page", None)          # 切范围回第 1 页
        for k in ("zone", "shop", "own_store", "inc"):
            params.pop(k, None)           # 范围参数由 over 决定
        for k, v in over.items():
            if v not in ("", None):
                params[k] = v
        qs = urlencode(params)
        return base + ("?" + qs if qs else "")

    def _comp_url(**over):
        # 竞店侧切换：携带当前自家店（总库=空则不带）——修复"切竞店自家店跳回总库"
        return _url(own_store=own_active or "", **over)

    comp = [{"label": "🌐 全部竞店", "title": "全部竞店合并总榜",
             "url": _comp_url(zone="__all__"), "active": comp_active == "__all__"}]
    if show_pick:
        comp.append({"label": "⊕ 多选对比",
                     "title": "勾选任意几家竞店对比（无需先建区；自家店单选）",
                     "url": _comp_url(zone="__pick"), "active": comp_active == "__pick"})
    for z in (zones or []):
        comp.append({"label": "🗺️ " + str(z.get("name") or ""),
                     "title": "区内竞店：" + "、".join(z.get("comp_shops", [])),
                     "url": _comp_url(zone=str(z.get("name") or "")),
                     "active": comp_active == str(z.get("name") or "")})
    for s in (free_shops or []):
        comp.append({"label": str(s), "title": "单店视图",
                     "url": _comp_url(shop=str(s)), "active": comp_active == str(s)})
    own = [{"label": "🏠 总库", "title": "自家总库（全部商品）",
            "url": _add_qs(_url(own_store=""), own_carry_qs), "active": not own_active}]
    _scope_txt = _own_sales_scope_text()
    for st in (own_stores or []):
        own.append({"label": "🏪 " + str(st.get("name") or ""),
                    "title": "在售 {} 品 · 销量口径={}".format(
                        st.get("on_sale_spus", "?"), _scope_txt),
                    "url": _add_qs(_url(own_store=str(st.get("name") or "")), own_carry_qs),
                    "active": own_active == str(st.get("name") or "")})
    return {"comp": comp, "own": own}


# ---------------------------------------------------------------------------
# P0.3 首页"看建议"：上次导入变化摘要（价格观测历史 diff，jsonl mtime 缓存）
# ---------------------------------------------------------------------------
_changes_cache = {"mtime": None, "data": None}


def _import_changes_summary():
    """全店"上次导入 → 本次导入"变化摘要：降价/涨价/新品/下架/断货/被打价 计数。

    数据源 = price_diff.diff_shop（价格观测历史 jsonl 最近两次快照 diff）。
    按 jsonl mtime 缓存（无新导入不重算）；观测不足 2 次的店给 note。
    返回 {"shops": {店: {counts:{}, rows, prev_t, cur_t, note}}, "total": {变动: N}}
    """
    hist = os.path.join(OUTPUT_DIR, "竞店标准库", "价格观测历史.jsonl")
    try:
        mt = os.path.getmtime(hist)
    except OSError:
        return {"shops": {}, "total": {}}
    if _changes_cache["mtime"] == mt and _changes_cache["data"] is not None:
        return _changes_cache["data"]
    shops, total = {}, {}
    try:
        import price_diff
        for shop in list(competitor_lookup.keys()):
            try:
                rows, meta = price_diff.diff_shop(shop, with_undercut=True)
            except Exception:
                continue
            cnt = {}
            for r in rows:
                k = r.get("变动") or "?"
                cnt[k] = cnt.get(k, 0) + 1
            shops[shop] = {"counts": cnt, "rows": len(rows),
                           "prev_t": meta.get("prev_t", ""), "cur_t": meta.get("cur_t", ""),
                           "note": meta.get("note", "")}
            for k, v in cnt.items():
                total[k] = total.get(k, 0) + v
    except Exception as e:
        print("[变化摘要] 计算失败:", e)
        return {"shops": {}, "total": {}}
    data = {"shops": shops, "total": total}
    _changes_cache["mtime"] = mt
    _changes_cache["data"] = data
    return data


@app.route("/")
def dashboard():
    """首页·榜单（决策页）：竞品品级销量榜 + 机会/防守/领先/引进候选筛选 + 关键词搜索 + 行内展开对照。
    合并原 工作台/产品分析/top50 三个页面：用户只看到"品"，核对在行内展开卡完成。"""
    import time as _t
    _dbg = [("start", _t.time())]
    shops = list(competitor_lookup.keys())
    # 区配置提前加载（收纳语义：被收进区的店从单店 tab 移除，只在区内出现）
    try:
        import zone_lib as _zlib
        zones_all = _zlib.load_zones()
    except Exception:
        _zlib, zones_all = None, []
    zoned_shops = set()
    for _z in zones_all:
        zoned_shops.update(_z.get("comp_shops", []))
    free_shops = [s for s in shops if s not in zoned_shops]   # 未收纳（单店 tab 可见）
    per = _get_stats_all()
    _dbg.append(("stats_all", _t.time()))
    shop = (request.args.get("shop") or "").strip()
    if not shop:
        if per:
            pendings = {s: v["pending"] for s, v in per.items() if s in free_shops}
            shop = max(pendings, key=pendings.get) if pendings and max(pendings.values()) > 0 \
                else (free_shops[0] if free_shops else "")
        elif free_shops:
            shop = free_shops[0]
    shop_order = sorted(free_shops, key=lambda s: -per.get(s, {}).get("pending", 0)) if per else free_shops
    g = _get_stats("")
    filter_f = (request.args.get("filter") or "hot").strip()
    if filter_f not in ("hot", "opp", "def", "lead", "all", "onlist", "gap", "deflead"):
        filter_f = "hot"
    q = (request.args.get("q") or "").strip()
    # 自家门店维度（多自家门店比对）：选中后"自家无对应"细分为 上架候选/引进候选
    import own_store_lib
    own_stores = own_store_lib.list_stores()
    _dbg.append(("own_stores", _t.time()))
    own_store = (request.args.get("own_store") or "").strip()
    own_on_sale = None
    if own_store and any(s["name"] == own_store for s in own_stores):
        own_on_sale = own_store_lib.get_on_sale(own_store)
        if own_on_sale is None:
            own_store = ""
    else:
        own_store = ""
    # P4.3 店对店：自家店 → 已指认仓的 {SPUID: 30天销量}；未指认/无数据 → None 并出告警（不回退合计）
    wh_map, wh_label = _wh_map_for_store(own_store)
    wh_status = _wh_status(own_store) if own_store else {}
    # 自家表真实 SPUID 集合（区分语义占位ID：只有真实ID才能判"该店未上架"）
    own_real_spuids = set(own_df["SPUID"].astype(str).str.strip()) if own_df is not None and len(own_df) else set()
    # 分页：每页 50（用户原 60 截断的问题：205 项只看到前 60）
    try:
        page = max(1, int(request.args.get("page") or 1))
    except (TypeError, ValueError):
        page = 1
    PAGE_SIZE = 50

    # ---- 分区模式（区=店铺归类汇总层，"区的总库"）：功能与单店一致 ----
    # 机会/防守/领先/引进候选筛选、状态徽章、搜索、分页全量可用；
    # 合并品的对应状态/结论按区内各店 (店,品) 口径聚合：
    #   任一店已确认 → 已匹配；全部店无对应 → 无对应；其余 → 待匹配
    #   结论取自家品组的产品分析结论；无任何对应且总月售≥30 → 引进候选
    # （zones_all/_zlib 已在路由开头加载：收纳语义下被收进区的店不进单店 tab）
    zone_name = (request.args.get("zone") or "").strip()
    zone = next((z for z in zones_all if z.get("name") == zone_name), None) if zone_name else None
    if zone_name == "__all__":
        # 虚拟区：全部竞店合并总榜（P0.1 范围选择条——"区的总库"扩展到全部店）
        zone = {"name": "__all__", "comp_shops": list(shops), "own_stores": [], "pseudo": True}
    elif zone_name == "__pick":
        # P0.3 多选对比（用户方案：竞店勾选多选+自家店单选，无需预建区；分区保留为收纳/比对项目）
        zone = {"name": "__pick", "comp_shops": list(shops), "own_stores": [], "pseudo": True,
                "pick": True}
    # 所有竞店都被收纳且未指定店/区 → 自动进第一个区（避免单店空榜）
    if (zone is None and not shop and not free_shops
            and zones_all and zones_all[0].get("comp_shops")):
        return redirect("/?zone=" + quote(zones_all[0]["name"]))
    if zone is not None and _zlib is not None:
        # 临时勾选子集（URL inc 参数；不改区配置）
        # 多选对比（__pick）：不勾选时不构建榜单，只渲染勾选面板
        zone_inc = [s for s in request.args.getlist("inc") if s in zone.get("comp_shops", [])]
        if zone.get("pick") and not zone_inc:
            zrank = {"items": [], "shops": []}
            zitems = []
            prod_ready = True
        else:
            zrank = _zone_rank_cached(zone, zone_inc)
            zitems = zrank.get("items", [])
        # 区内每店桥接映射（懒加载：无则现算 + 触发预热；产品分析缺则结论暂空）
        zbridges = {}
        prod_ready = True
        for s in (zone_inc or zone.get("comp_shops", [])):
            if s not in competitor_lookup:
                continue
            if s in _bridge_cache:
                zbridges[s] = _bridge_cache[s]
                continue
            if (s, "detail") not in _compare_cache or (s, "product") not in _compare_cache:
                threading.Thread(target=_warmup_one_shop, args=(s,), daemon=True).start()
                if (s, "product") not in _compare_cache:
                    prod_ready = False
            zbridges[s] = _build_bridge(s)
            if zbridges[s][3]:
                # 只缓存非空结论桥接（同单店分支：防首访竞态缓存空结论）
                _bridge_cache[s] = zbridges[s]
        # 合并品聚合：状态/结论/自家品组ID/自家月销/自家价格带
        for it in zitems:
            gid, gname, matched, nomatched = "", "", False, False
            for m in it.get("店铺明细", []):
                br = zbridges.get(m.get("shop"))
                if br is None:
                    continue
                _own_sku_map, spu_own, _sku2own, _cm = br
                conf_map = progress.get("confirmed", {}).get(m.get("shop"), {})
                nm_map = progress.get("no_match", {}).get(m.get("shop"), {})
                if conf_map.get(m.get("sku_id") or ""):
                    matched = True
                if nm_map.get(m.get("sku_id") or ""):
                    nomatched = True
                g2, n2 = spu_own.get(m.get("comp_id") or "", ("", ""))
                if g2 and not gid:
                    gid, gname = g2, n2
            # 参考桥接（P2.1 显示层前置）：确认链全空时用品级语义高置信候选给参考结论
            it["对应参考"] = ""
            if not gid:
                for m in it.get("店铺明细", []):
                    ref = _ref_bridge(m.get("shop") or "").get(str(m.get("comp_id") or ""))
                    if ref:
                        gid, gname = ref[0], ref[1]
                        it["对应参考"] = f"参考·置信{ref[2]:.2f}"
                        break
            # 结论/自家月销：品组ID 在任一店的产品分析里查（同组结论同源）
            concl, own_sales = "", ""
            if gid:
                for br in zbridges.values():
                    cm = br[3]
                    if gid in cm:
                        concl, own_sales = cm[gid]
                        break
            it["自家品组ID"] = gid
            it["自家品名"] = gname
            it["自家月销"] = own_sales if own_sales not in (None, "") else ""
            # 店对店口径：选自家店且匹配到对应仓 → 对比条/展示用 仓级销量（结论判定仍=总库合计）
            if wh_map is not None and gid:
                _wv = wh_map.get(gid)
                if _wv is not None:
                    it["自家月销"] = round(_wv, 0)
            # 区榜的 有效月售 = 成交 + λ×收藏（见 zone_lib），与单店榜同一口径
            _zev = it.get("有效月售")
            zms = _ST.num(_zev) if _zev not in (None, "") else _ST.num(it.get("总月售"))
            if zms >= 30:
                if not gid:
                    concl = "引进候选"
                elif (own_on_sale is not None and gid in own_real_spuids
                      and gid not in own_on_sale["spu_set"]):
                    concl = "上架候选"
            it["结论"] = concl
            it["状态"] = "已匹配" if matched else ("无对应" if nomatched else "待匹配")
            if gid and not matched and not nomatched and not it.get("对应参考"):
                it["状态"] = "已匹配"   # SPU 层有效映射视同已对应（同单店口径）；参考级不算
            band = _own_price_cache.get(gid) if gid else None
            it["自家价格带"] = band or ("", "")
            # ---- P0.2 区缺口判定：≥2 店在售 + 自家没有（总库缺=引进；门店缺=上架候选） ----
            it["在售店数"] = len(it.get("店铺明细") or [])
            gap = ""
            if it["在售店数"] >= 2:
                if not gid:
                    gap = "总库缺"
                elif (own_on_sale is not None and gid in own_real_spuids
                      and gid not in own_on_sale["spu_set"]):
                    gap = "门店缺"
            it["缺口"] = gap
        # 主店代表制补齐：合并品取月售最高店为代表（规格/价格徽章/比价/对应确认口径），
        # 行结构与单店 rk-row 完全一致 → 模板复用同一套榜单行+展开卡
        def _zf(v):
            try:
                return float(str(v).replace(",", "").strip() or 0)
            except (TypeError, ValueError):
                return 0.0
        _hist = _load_price_hist()
        _hfull = _load_hist_full()  # P1.3 观测期数（区合并行同样标注）
        _zprods_cache = {}
        for it in zitems:
            ms = max(it.get("店铺明细") or [], key=lambda m: _zf(m.get("月售")), default=None)
            if ms is None:
                it["主店"], it["月售"] = "", (it.get("总月售") or 0)
                continue
            it["主店"] = ms.get("shop", "")
            it["comp_id"] = str(ms.get("comp_id") or "")
            # 主店产品详情（SKU列表/价格/原价/图片）——top500 缓存按店取
            if it["主店"] not in _zprods_cache:
                _zprods_cache[it["主店"]] = _zone_rank_getter(it["主店"]) or []
            pm = next((x for x in _zprods_cache[it["主店"]]
                       if str(x.get("comp_id")) == it["comp_id"]), None)
            if pm is not None:
                for k in ("SKU列表", "价格", "原价", "图片", "规格数", "类目", "品牌", "品名"):
                    if pm.get(k) not in (None, "", []):
                        it[k] = pm[k]
            if not it.get("SKU列表"):
                # 保底：主店 top500 缓存未命中时用店铺明细合成（模板依赖非空 SKU列表）
                it["SKU列表"] = [{"sku_id": str(ms.get("sku_id") or ""), "规格": "（默认）",
                                  "月售": ms.get("月售", ""), "价格": ms.get("价格", ""),
                                  "原价": ""}]
            # 价格徽章/比价（同单店口径：折扣/较上次/同类分位/自家价差/参考价带）
            it["折扣"] = _discount_badge(it.get("价格"), it.get("原价"))
            try:
                _msku = (it.get("SKU列表") or [{}])[0].get("sku_id", "")
            except (IndexError, AttributeError):
                _msku = ""
            it["较上次"] = _price_delta(it["主店"], _msku, it.get("价格"), _hist) if _msku else None
            it["观测期数"] = len(_hfull.get(it["主店"], {}).get(_msku, [])) if _msku else 0
            it["分位"] = _price_percentile(it.get("类目"), it.get("价格"))
            band = it.get("自家价格带") or ("", "")
            if band and band[0] != "":
                mid = round((band[0] + band[1]) / 2, 2)
                it["价差"] = _price_diff(it.get("价格"), mid)
            else:
                it["价差"] = None
            if not it.get("自家品组ID"):
                _prs = [_extract_num(s.get("价格")) for s in (it.get("SKU列表") or [])]
                _prs = [v for v in _prs if v is not None]
                it["参考价带"] = (round(min(_prs), 2), round(max(_prs), 2)) if _prs else ("", "")
            else:
                it["参考价带"] = ("", "")
            it["月售"] = it.get("总月售") or 0   # 竞售条/排名用合并总月售
        # 统计（筛选前）
        zstats = {"total": len(zitems), "matched": sum(1 for x in zitems if x["状态"] == "已匹配"),
                  "opp": sum(1 for x in zitems if x["结论"] == "机会品"),
                  "def": sum(1 for x in zitems if x["结论"] == "防守品"),
                  "lead": sum(1 for x in zitems if x["结论"] == "领先品"),
                  "intro": sum(1 for x in zitems if x["结论"] == "引进候选"),
                  "onlist": sum(1 for x in zitems if x["结论"] == "上架候选"),
                  "gap": sum(1 for x in zitems if x.get("缺口") and x["在售店数"] >= 2)}
        zstats["deflead"] = zstats["def"] + zstats["lead"]
        # 筛选（同单店口径）
        if filter_f == "opp":
            zitems = [x for x in zitems if x["结论"] in ("机会品", "引进候选", "上架候选")]
        elif filter_f == "def":
            zitems = [x for x in zitems if x["结论"] == "防守品"]
        elif filter_f == "lead":
            zitems = [x for x in zitems if x["结论"] == "领先品"]
        elif filter_f == "gap":
            # P0.2 区缺口榜：≥2 店在售且自家没有，按 店数→总月售 排序（多店共卖=需求信号更强）
            zitems = [x for x in zitems if x.get("缺口") and x["在售店数"] >= 2]
            zitems.sort(key=lambda x: (-x["在售店数"], -(x.get("总月售") or 0)))
        elif filter_f == "deflead":
            # P0.3 Tab 收敛：防守+领先合并一档（自家强 = 要守或已领先）
            zitems = [x for x in zitems if x["结论"] in ("防守品", "领先品")]
        zstats["filtered"] = len(zitems)
        # 搜索（同单店口径：品名/品牌/类目/规格词/自家品名）
        if q:
            try:
                from search_util import kw_hit
                def _zrow_keys(it):
                    sku_specs = " ".join(_safe(s.get("规格")) for s in (it.get("SKU列表") or []))
                    return [_safe(it.get("品名")), _safe(it.get("品牌")),
                            _safe(it.get("类目")), sku_specs,
                            _safe(it.get("自家品名") or "")]
                zitems = [it for it in zitems
                          if kw_hit({"_": "\n".join(_zrow_keys(it))}, q, ["_"])]
            except Exception as e:
                print("[区榜] 搜索异常:", e)
            zstats["filtered"] = len(zitems)
        ztotal_pages = max(1, (len(zitems) + PAGE_SIZE - 1) // PAGE_SIZE)
        if page > ztotal_pages:
            page = ztotal_pages
        zrows = zitems[(page - 1) * PAGE_SIZE: page * PAGE_SIZE]
        # 候选池状态（区模式行内一键入池：合并品跨店全量入池）
        try:
            _pool_keys = {(e.get("shop"), str(e.get("comp_id")))
                          for e in _cp.load_pool()["entries"]}
        except Exception:
            _pool_keys = set()
        for it in zrows:
            _ms = it.get("店铺明细") or []
            it["全在池"] = all((m.get("shop"), str(m.get("comp_id") or "")) in _pool_keys
                              for m in _ms) if _ms else False
        snaps = _shop_snapshot_times()
        for sm in zrank.get("shops", []):
            sm["快照时间"] = snaps.get(sm["shop"], "")
        # 区模式公共查询串（筛选/分页共用：zone + inc子集 + own_store + q）
        zqs = f"/?zone={quote(zone['name'])}"
        for s in zone_inc:
            zqs += f"&inc={quote(s)}"
        if own_store:
            zqs += f"&own_store={quote(own_store)}"
        if q:
            zqs += f"&q={quote(q)}"
        # P0.3 范围联动：自家选项携带当前竞店范围（区/多选勾选/全部）
        own_carry = "&zone=" + quote(zone['name']) + "".join("&inc=" + quote(s) for s in zone_inc)
        # 区/多选视图的 AI 建议：合并所选店的建议（最多 5 家，50 条），跳转各回自己店
        advice_z, _adv = [], {"rows": [], "stale": False, "at": ""}
        for s in (zone_inc or zone.get("comp_shops", []))[:5]:
            _st = advice_state(s)
            _adv["rows"] += _st["rows"]
            if _st["stale"] and not _adv["stale"]:
                _adv["stale"], _adv["at"] = True, _st["at"]
            advice_z += _advice_view(s, _st["rows"])
        advice_z = advice_z[:50]
        scope_bar = build_scope_bar("/", comp_active=zone_name, own_active=own_store,
                                    keep=("filter", "q"), zones=zones_all,
                                    free_shops=free_shops, own_stores=own_stores,
                                    own_carry_qs=own_carry)
        return render_template(
            "dashboard.html",
            g=g, shop_count=len(shops), shops=shop_order, current_shop=shop,
            filter_f=filter_f, q=q,
            own_store=own_store, own_stores=own_stores,
            rank_rows=[], rank_stats=zstats,
            rank_loading=False, prod_ready=prod_ready,
            page=page, total_pages=ztotal_pages, page_size=PAGE_SIZE,
            advice=advice_z, advice_stale=_adv["stale"], advice_at=_adv["at"],
            pool_keys=_pool_keys,
            zone=zone, zones=zones_all, zone_rows=zrows,
            zone_shops=zrank.get("shops", []),
            zone_inc=zone_inc, zqs=zqs, scope_bar=scope_bar,
            changes=_import_changes_summary(),
            sales_scope=_own_sales_scope_text(),
            wh_label=wh_label, wh_status=wh_status)

    # ---- 榜单数据：竞品品级聚合（月售降序，n=500 缓存）----
    rank_rows, rank_stats, rank_loading = [], {}, False
    rank_stats = {"total": 0, "matched": 0, "opp": 0, "def": 0, "lead": 0, "intro": 0,
                  "onlist": 0, "deflead": 0}
    prod_ready = False
    prods_all = []  # 筛选+搜索前全量
    try:
        from top50 import build_top50
        key5 = (shop, "top500")
        if key5 not in _compare_cache:
            _compare_cache[key5] = _derived("top500", shop,
                                            lambda: build_top50(shop, 500), code=_fp_derived())
        prods = list(_compare_cache[key5] or [])
        _dbg.append(("top50_build", _t.time()))
        # 链：榜单品SKU → 对齐明细(sku→自家SKUID) → own_df(SKUID→SPUID真实品组ID)
        # → 品卡 concl_map(品组ID→结论/自家月销)。
        # 旧链路绕汇总库 SPU 层，其"自家品组ID"多为语义占位ID(SPU_O*)，
        # 与品卡真实数字 ID 永不相交 → 语义对齐的店（无人工确认）结论桥接全断。
        # 四个映射缓存（跟 _compare_cache 同失效），避免每次请求遍历 own_df(4.4万行)
        bc_key = shop
        if bc_key in _bridge_cache:
            own_sku_map, spu_own, sku2own, concl_map = _bridge_cache[bc_key]
            prod_ready = bool(concl_map)
            if (shop, "detail") not in _compare_cache:
                threading.Thread(target=_warmup_one_shop, args=(shop,), daemon=True).start()
            if not prod_ready and (shop, "product") not in _compare_cache:
                threading.Thread(target=_warmup_one_shop, args=(shop,), daemon=True).start()
        else:
            if (shop, "detail") not in _compare_cache or (shop, "product") not in _compare_cache:
                threading.Thread(target=_warmup_one_shop, args=(shop,), daemon=True).start()
            own_sku_map, spu_own, sku2own, concl_map = _build_bridge(shop)
            prod_ready = bool(concl_map) and (shop, "product") in _compare_cache
            _dbg.append(("bridge_build", _t.time()))
            if concl_map:
                # 只缓存非空结论桥接（防首访竞态：品卡未预热完时把空结论缓存住，
                # 此后防守/领先永远为 0——用户实报"防守领先没有内容"的根因）
                _bridge_cache[bc_key] = (own_sku_map, spu_own, sku2own, concl_map)
        conf_map = progress.get("confirmed", {}).get(shop, {})
        nm_map = progress.get("no_match", {}).get(shop, {})
        _hist = _load_price_hist()  # 路由级取一次：Windows 单次 stat ~7ms，循环内每品查会拖垮请求
        _hfull = _load_hist_full()  # P1.3 观测期数
        _dbg.append(("pre_loop", _t.time()))
        for p in prods:
            pid = p["comp_id"]
            skus = [s["sku_id"] for s in p.get("SKU列表", [])]
            conf = [s for s in skus if conf_map.get(s)]
            nm = [s for s in skus if nm_map.get(s)]
            # 品级对应解析：优先 SKU 级真实 ID 链路（对齐明细），回退 SPU 层映射
            gid, gname = spu_own.get(pid, ("", ""))
            own_sales_agg = 0.0
            own_prices_sku = []  # SKU 链路收集的自家售价（比价用）
            if not gid or gid not in own_real_spuids:
                for s in skus:
                    own_sk = sku2own.get(s)
                    if own_sk and own_sk in own_sku_map:
                        g2, n2, sales2, price2 = own_sku_map[own_sk]
                        if g2 and g2 in concl_map:
                            gid = g2
                            gname = gname or n2
                            try:
                                own_sales_agg += float(sales2 or 0)
                            except (TypeError, ValueError):
                                pass
                            pv = _extract_num(price2)
                            if pv is not None:
                                own_prices_sku.append(pv)
            # 参考桥接（P2.1 显示层前置）：确认链全空时用品级语义高置信候选给参考结论
            p["对应参考"] = ""
            if not gid or gid not in own_real_spuids:
                ref = _ref_bridge(shop).get(str(pid))
                if ref:
                    gid, gname = ref[0], (gname or ref[1])
                    p["对应参考"] = f"参考·置信{ref[2]:.2f}"
            if conf:
                p["状态"] = "已匹配"
            elif nm:
                p["状态"] = "无对应"
            else:
                p["状态"] = "待匹配"
            p["自家品组ID"] = gid
            p["自家品名"] = gname
            concl, own_sales = concl_map.get(gid, ("", ""))
            if own_sales in ("", None) and own_sales_agg > 0:
                own_sales = int(own_sales_agg)
            p["自家月销"] = own_sales if own_sales is not None else ""
            # 店对店口径：选自家店且匹配到对应仓 → 对比条/展示用 仓级销量（结论判定仍=总库合计）
            if wh_map is not None and gid:
                _wv = wh_map.get(gid)
                if _wv is not None:
                    p["自家月销"] = round(_wv, 0)
            # 无品级对应且竞品月售达门槛 -> 引进候选（自家无对应，参考竞品价格带）
            # 选定自家门店后细分：总库有(gid)但该店未上架 -> 上架候选（动作成本更低）
            # 注意：gid 可能是语义占位ID（SPU_O*，非真实总库ID）——只有真实存在于
            # 自家表(own_df)的 gid 才判上架候选，占位ID保持原结论（防误判）。
            # 门槛按有效销量：收藏（「N人想买」）按 λ 折算后参与，算得轻但不是不算。
            _ev = p.get("有效销量")
            ms = (_ST.num(_ev) if _ev not in (None, "")
                  else _ST.eff_of(p.get("月售"), p.get("意向")))
            if ms >= 30:
                if not gid:
                    concl = "引进候选"
                elif (own_on_sale is not None and gid in own_real_spuids
                      and gid not in own_on_sale["spu_set"]):
                    concl = "上架候选"
            p["结论"] = concl
            # ---- 比价展示：竞品折扣 + 自家价格带 + 价差（预览框直出，无需展开）----
            p["折扣"] = _discount_badge(p.get("价格"), p.get("原价"))
            # 价格变动：主 SKU（销量最高规格）较上次观测的变动
            try:
                _msku = (p.get("SKU列表") or [{}])[0].get("sku_id", "")
            except (IndexError, AttributeError):
                _msku = ""
            p["较上次"] = _price_delta(shop, _msku, p.get("价格"), _hist) if _msku else None
            p["观测期数"] = len(_hfull.get(shop, {}).get(_msku, [])) if _msku else 0
            p["分位"] = _price_percentile(p.get("类目"), p.get("价格"))
            if gid:
                band = _own_price_cache.get(gid)
                if band is None and own_prices_sku:
                    band = (round(min(own_prices_sku), 2), round(max(own_prices_sku), 2))
                p["自家价格带"] = band or ("", "")
                if band:
                    mid = round((band[0] + band[1]) / 2, 2)
                    p["价差"] = _price_diff(p.get("价格"), mid)
                else:
                    p["价差"] = None
            else:
                p["自家价格带"] = ("", "")
                p["价差"] = None
            # 参考价格带（品内 SKU 价格 min-max，供引进候选参考）
            if not gid:
                prices = []
                for s in p.get("SKU列表", []):
                    v = _extract_num(s.get("价格"))
                    if v is not None:
                        prices.append(v)
                p["参考价带"] = (round(min(prices), 2), round(max(prices), 2)) if prices else ("", "")
        prods_all = list(prods)
        _dbg.append(("post_loop", _t.time()))
        # 全量统计（筛选前）
        rank_stats["total"] = len(prods)
        rank_stats["matched"] = sum(1 for p in prods if p["状态"] == "已匹配")
        rank_stats["opp"] = sum(1 for p in prods if p["结论"] == "机会品")
        rank_stats["def"] = sum(1 for p in prods if p["结论"] == "防守品")
        rank_stats["lead"] = sum(1 for p in prods if p["结论"] == "领先品")
        rank_stats["intro"] = sum(1 for p in prods if p["结论"] == "引进候选")
        rank_stats["onlist"] = sum(1 for p in prods if p["结论"] == "上架候选")
        rank_stats["deflead"] = rank_stats["def"] + rank_stats["lead"]
        # 筛选
        if filter_f == "opp":
            prods = [p for p in prods if p["结论"] in ("机会品", "引进候选", "上架候选")]
        elif filter_f == "onlist":
            prods = [p for p in prods if p["结论"] == "上架候选"]
        elif filter_f == "def":
            prods = [p for p in prods if p["结论"] == "防守品"]
        elif filter_f == "lead":
            prods = [p for p in prods if p["结论"] == "领先品"]
        elif filter_f == "deflead":
            # P0.3 Tab 收敛：防守+领先合并一档（自家强 = 要守或已领先）
            prods = [p for p in prods if p["结论"] in ("防守品", "领先品")]
        # 关键词搜索（多词 AND；与竞品总览/详情页同 search_util 逻辑一致）
        if q:
            try:
                from search_util import kw_hit
                def _row_keys(p):
                    sku_specs = " ".join(_safe(s.get("规格")) for s in p.get("SKU列表", []))
                    return [_safe(p.get("品名")), _safe(p.get("品牌")),
                            _safe(p.get("类目")), sku_specs,
                            _safe(p.get("自家品名") or "")]
                prods = [p for p in prods if kw_hit({"_": "\n".join(_row_keys(p))}, q, ["_"])]
            except Exception as e:
                print("[榜单] 搜索异常:", e)
        rank_stats["filtered"] = len(prods)
        # 分页
        total_pages = max(1, (len(prods) + PAGE_SIZE - 1) // PAGE_SIZE)
        if page > total_pages:
            page = total_pages
        rank_rows = prods[(page - 1) * PAGE_SIZE: page * PAGE_SIZE]
    except Exception as e:
        print("[榜单] 加载失败:", e)
        rank_loading = True
        total_pages = 1
        page = 1
    # AI 建议：跳转按代表 SKU 直达竞品详情（见 _advice_view）；取数走 advice_state（内存→快照）
    _adv = advice_state(shop)
    advice_view = _advice_view(shop, _adv["rows"])
    _dbg.append(("pre_render", _t.time()))
    _dbg_total = _dbg[-1][1] - _dbg[0][1]
    if _dbg_total > 0.5:
        parts = [f"{n}={_dbg[i+1][1]-t:.3f}" for i,(n,t) in enumerate(_dbg[:-1])]
        print(f"[dashboard] 总{_dbg_total:.2f}s | " + " ".join(parts))
    # 候选池成员（引进候选按钮：已在池则显示"已在池"）
    try:
        _pool_keys = {(e.get("shop"), str(e.get("comp_id")))
                      for e in _cp.load_pool()["entries"]}
    except Exception:
        _pool_keys = set()
    scope_bar = build_scope_bar("/", comp_active=shop, own_active=own_store,
                                keep=("filter", "q"), zones=zones_all,
                                free_shops=free_shops, own_stores=own_stores,
                                own_carry_qs=("&shop=" + quote(shop)) if shop else "")
    return render_template("dashboard.html",
                           g=g, shop_count=len(shops), shops=shop_order, current_shop=shop,
                           filter_f=filter_f, q=q,
                           own_store=own_store, own_stores=own_stores,
                           rank_rows=rank_rows, rank_stats=rank_stats,
                           rank_loading=rank_loading, prod_ready=prod_ready,
                           page=page, total_pages=total_pages, page_size=PAGE_SIZE,
                           advice=advice_view, advice_stale=_adv["stale"],
                           advice_at=_adv["at"], pool_keys=_pool_keys,
                           zone=None, zones=zones_all, zone_rows=[], zone_shops=[],
                           zone_inc=[], zqs="", scope_bar=scope_bar,
                           changes=_import_changes_summary(),
                           obs=_shop_obs_meta().get(shop, {}),
                           sales_scope=_own_sales_scope_text(),
                           wh_label=wh_label, wh_status=wh_status)


# 竞品库浏览行按店缓存在内存（实测每请求重建有 50 万次 Series 字段访问 ≈1.3s；
# 缓存字典定义在全局数据区，随 _invalidate_stats 一起作废）


def _browse_pool_idx_map():
    if pool_df is None or not len(pool_df):
        return {}
    return {_safe(sku): pool_df.index[i]
            for i, sku in enumerate(pool_df["竞品Sku_id"])}


def _browse_shop_rows(s):
    """单店竞品浏览行（含匹配状态/比价字段/AI 预筛徽章）。按店缓存后复用。"""
    if s in _browse_rows_cache:
        return _browse_rows_cache[s]
    from ai_prescreen import verdict_badge
    lk = competitor_lookup.get(s) or {}
    pool_idx_map = _browse_pool_idx_map()
    prescreen = _prescreen_map()
    hist = _load_price_hist()
    hfull = _load_hist_full()
    conf_map = progress.get("confirmed", {}).get(s, {})
    nm_set = set(progress.get("no_match", {}).get(s, {}).keys())
    out = []
    for sku, row in lk.items():
        lvl = _converge_level(s, sku, row)
        rec = conf_map.get(sku)
        if lvl == "sku":
            st = "aligned"
            own_info = aligned_map.get(sku, {})
        elif lvl == "spu":
            st = "coarse"
            pid = _safe(row.get("comp_id") or row.get("商品ID"))
            own_name = spu_layer_info.get((s, pid), "") if pid else ""
            own_info = {"自家SKUID": "", "自家商品名": own_name,
                        "匹配依据": "品已锁定（规格待补）"}
        elif sku in nm_set:
            st = "nomatch"
            own_info = {}
        else:
            st = "pending"
            own_info = {}
        # AI 预筛徽章：待处理行必挂；AI 自动判无对应的行也挂（可撤销）
        ps_rec = prescreen.get((s, str(sku))) or {}
        show_ai = st == "pending" or ps_rec.get("triage") == "自动无对应"
        # 比价字段：原价/折扣；已对应行补自家售价与价差（O(1) 查找）
        _price = _safe(row.get("price") or row.get("到手价") or row.get("折后价"))
        _orig = _safe(row.get("original_price") or row.get("原价"))
        _own_price, _diff = "", None
        if st == "aligned" and own_info.get("自家SKUID"):
            _own_price = _safe(get_own_info(own_info["自家SKUID"]).get("售价"))
            _diff = _price_diff(_price, _own_price)
        _cat = _comp_cat(row)
        out.append({
            "店铺": s,
            "Sku_id": sku,
            "商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
            "品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
            "型号": _safe(row.get("型号")),
            "类目": _cat,
            "价格": _price,
            "原价": _orig,
            "折扣": _discount_badge(_price, _orig),
            "自家售价": _own_price,
            "价差": _diff,
            "较上次": _price_delta(s, sku, _price, hist),
            "观测期数": len(hfull.get(s, {}).get(sku, [])),
            "分位": _price_percentile(_cat, _price),
            "月售": _clean_sales(row.get("sales") or row.get("月销量") or row.get("月售")),
            "图片": _safe(row.get("image_url") or row.get("图片url") or row.get("图片链接")),
            "状态": st,
            "收敛层级": lvl,
            "自家SKUID": own_info.get("自家SKUID", ""),
            "自家商品名": own_info.get("自家商品名", ""),
            "匹配依据": own_info.get("匹配依据", ""),
            "pool_idx": pool_idx_map.get(sku, None),
            "AI预筛": verdict_badge(ps_rec)[0] if show_ai and ps_rec else "",
            "AI预筛理由": _safe(ps_rec.get("reason")) if show_ai and ps_rec else "",
        })
    _browse_rows_cache[s] = out
    return out


def _browse_shops_meta():
    """各店匹配状态计数（按店缓存；口径：已对应=SKU层有效，已粗对齐=品锁定）。"""
    if _browse_meta_cache:
        return _browse_meta_cache
    pool_shop_skus = _build_pool_shop_skus()
    meta = {}
    for s, lk in competitor_lookup.items():
        conf_map = progress.get("confirmed", {}).get(s, {})
        nm_set = set(progress.get("no_match", {}).get(s, {}).keys())
        n_align = sum(1 for sku in lk if sku in aligned_map
                      or (sku in conf_map and not conf_map.get(sku, {}).get("spu_only")))
        n_spu = sum(1 for sku, row in lk.items()
                    if sku in conf_map and conf_map.get(sku, {}).get("spu_only")
                    or _converge_level(s, sku, row) == "spu")
        n_nm = sum(1 for sku in lk if sku in nm_set)
        meta[s] = {"name": s, "total": len(lk), "aligned": n_align,
                   "coarse": n_spu, "pending": len(lk) - n_align - n_spu - n_nm,
                   "nomatch": n_nm}
    _browse_meta_cache.update(meta)
    return meta


@app.route("/browse")
def browse():
    """竞品总览（可视化优先）：全部竞品分页浏览 + 搜索/品牌/类目/状态筛选 + 价格/月售排序。
    shop 为空 = 全部店铺跨店视图。匹配仅是辅助标注。"""
    page = request.args.get("page", 1, type=int)
    shop = (request.args.get("shop") or "").strip()
    q = (request.args.get("q") or "").strip()
    brand = (request.args.get("brand") or "").strip()
    cat = (request.args.get("cat") or "").strip()
    status_f = (request.args.get("status") or "all").strip()
    if status_f not in ("all", "aligned", "coarse", "pending", "nomatch", "none"):
        status_f = "all"
    level_f = (request.args.get("level") or "").strip()
    if level_f not in ("sku", "spu", "cat", "none"):
        level_f = ""
    untreated = request.args.get("untreated") == "1"
    only_cat = request.args.get("only_cat") == "1"
    # M7 阶段2：AI 关注品筛选（harness → 列表页）
    focus_skus_raw = (request.args.get("focus_skus") or "").strip()
    focus_skus = set(s.strip() for s in focus_skus_raw.split(",") if s.strip()) \
        if focus_skus_raw else set()
    sort_by = (request.args.get("sort") or "").strip()
    if sort_by not in ("price", "sales"):
        sort_by = ""
    order = (request.args.get("order") or "asc").strip()
    if order not in ("asc", "desc"):
        order = "asc"

    # ---- P0.1 范围选择：zone 参数 → 区内多店视图（其余店不出现在列表） ----
    try:
        import zone_lib as _zlib_b
        zones_b = _zlib_b.load_zones()
    except Exception:
        zones_b = []
    zone_name_b = (request.args.get("zone") or "").strip()
    zone_b = next((z for z in zones_b if z.get("name") == zone_name_b), None) if zone_name_b else None

    # 店铺统计（口径：已对应=SKU层有效；已粗对齐=品锁定；无对应；待处理=其余）按店缓存
    _meta_all = _browse_shops_meta()
    shops_meta = [_meta_all[s] for s in competitor_lookup if s in _meta_all]

    if shop and shop not in competitor_lookup:
        return render_template("browse.html", items=[], page=1, total_pages=1,
                               total=0, stats=None, shops=shops_meta, current_shop=shop,
                               q=q, brand=brand, cat=cat, status=status_f,
                               brands=[], cats=[], total_all=0, sort=sort_by, order=order)

    # 构建浏览行（含匹配状态；shop 为空=全部店铺）——按店取内存缓存，不再每请求重建
    rows = []
    if zone_b:
        # 区视图：区内各竞店全部 SKU（存在店才进迭代，防区配置过期）
        shop_iter = [s for s in zone_b.get("comp_shops", []) if s in competitor_lookup]
    elif shop:
        shop_iter = [shop]
    else:
        shop_iter = list(competitor_lookup.keys())
    for s in shop_iter:
        rows.extend(_browse_shop_rows(s))

    # 全量选项（未筛选口径，避免选了类目后其他类目消失）
    all_brands = sorted({r["品牌"] for r in rows if r["品牌"]})
    all_cats = sorted({r["类目"] for r in rows if r["类目"]})

    # 筛选
    if q:
        from search_util import kw_hit
        rows = [r for r in rows if kw_hit(r, q, ["商品名", "Sku_id", "品牌"])]
    if brand:
        rows = [r for r in rows if r["品牌"] == brand]
    if cat:
        rows = [r for r in rows if r["类目"] == cat]
    if status_f != "all":
        rows = [r for r in rows if r["状态"] == status_f]
    if level_f:
        rows = [r for r in rows if r["收敛层级"] == level_f]
    if untreated:
        rows = [r for r in rows if r["状态"] in ("pending", "coarse")]
    if only_cat:
        rows = [r for r in rows if r["类目"] and r["类目"] in cat_map_cache]
    if focus_skus:
        rows = [r for r in rows if str(r.get("Sku_id") or "") in focus_skus]

    # 排序（价格/月售）
    def _num(v):
        try:
            return float(str(v).replace(",", "").replace("¥", "").strip() or "nan")
        except ValueError:
            return float("nan")

    if sort_by == "sales":
        # 月售分档排序：真实成交 > 意向 > 无数据；同档内按数字排
        desc = (order == "desc")
        real, intent, empty = [], [], []
        for r in rows:
            s = _clean_sales(r["月售"])
            if not s:
                empty.append(r)
            elif "(意向)" in s:
                intent.append(r)
            else:
                real.append(r)
        def _num_key(r):
            try:
                return float(_clean_sales(r["月售"]).replace(",", "").replace("(意向)", ""))
            except ValueError:
                return 0
        real.sort(key=_num_key, reverse=desc)
        intent.sort(key=_num_key, reverse=desc)
        rows = real + intent + empty
    elif sort_by == "price":
        # 价格排序：提取数字中位数，无数字排最后
        def _k(r):
            return _extract_num(r["价格"])
        has_rows = [r for r in rows if _k(r) is not None]
        none_rows = [r for r in rows if _k(r) is None]
        has_rows.sort(key=_k, reverse=(order == "desc"))
        rows = has_rows + none_rows

    # 统计（当前范围未筛选口径，5 项：总数/待处理/已粗对齐/已对应/无对应）
    stats = {"total": 0, "pending": 0, "coarse": 0, "aligned": 0, "nomatch": 0}
    _allowed = set(shop_iter)
    for s in shops_meta:
        if s["name"] in _allowed:
            stats["total"] += s["total"]
            stats["aligned"] += s["aligned"]
            stats["coarse"] += s["coarse"]
            stats["pending"] += s["pending"]
            stats["nomatch"] += s["nomatch"]

    total = len(rows)
    total_pages = max(1, (total + 50 - 1) // 50)
    page = max(1, min(page, total_pages))
    start = (page - 1) * 50
    items = rows[start:start + 50]

    # P0.1 范围选择条 + 模板链接透传串（zone/shop 由 scope_qs 统一携带）
    _zoned_set = {z for zz in zones_b for z in zz.get("comp_shops", [])}
    free_shops_b = [s for s in competitor_lookup.keys() if s not in _zoned_set]
    # P1.4 数据时间显性化：当前范围最新观测时间
    om = _shop_obs_meta()
    _lasts = [om[s]["last_t"] for s in shop_iter if om.get(s, {}).get("last_t")]
    obs_last = max(_lasts) if _lasts else ""
    scope_bar = build_scope_bar(
        "/browse", comp_active=(zone_name_b if zone_b else shop),
        keep=("q", "brand", "cat", "status", "level", "untreated", "only_cat", "sort", "order"),
        zones=zones_b, free_shops=free_shops_b, show_pick=False)
    scope_qs = ("&zone=" + quote(zone_name_b)) if zone_b else (("&shop=" + quote(shop)) if shop else "")

    return render_template("browse.html", items=items, page=page, total_pages=total_pages,
                           total=total, stats=stats, shops=shops_meta, current_shop=shop,
                           q=q, brand=brand, cat=cat, status=status_f,
                           level=level_f, untreated=untreated, only_cat=only_cat,
                           brands=all_brands, cats=all_cats, total_all=stats["total"],
                           sort=sort_by, order=order,
                           zone_name=(zone_name_b if zone_b else ""),
                           scope_bar=scope_bar, scope_qs=scope_qs,
                           focus_n=len(focus_skus),
                           obs_last=obs_last)


@app.route("/comp_detail")
def comp_detail():
    """竞品详情（全部竞品可看）：竞品完整信息 + 匹配状态 + 候选/对应展示。"""
    shop = (request.args.get("shop") or "").strip()
    sku = (request.args.get("sku") or "").strip()
    if shop not in competitor_lookup or not sku:
        abort(404)
    row = competitor_lookup.get(shop, {}).get(sku)
    if row is None:
        abort(404)

    # 匹配状态（统一用收敛层级判定：spu_only 记录 = 已粗对齐而非已对应）
    nm_set = set(progress.get("no_match", {}).get(shop, {}).keys())
    lvl = _converge_level(shop, sku, row)
    if lvl == "sku":
        st = "aligned"
        own_info = aligned_map.get(sku, {})
    elif lvl == "spu":
        st = "coarse"
        pid = _safe(row.get("comp_id") or row.get("商品ID"))
        own_name = spu_layer_info.get((shop, pid), "") if pid else ""
        own_info = {"自家SKUID": "", "自家商品名": own_name,
                    "匹配依据": "品已锁定（规格待补）"}
    elif sku in nm_set:
        st = "nomatch"
        own_info = {}
    else:
        pool_idx = None
        own_info = {}
        if pool_df is not None and len(pool_df):
            m = pool_df[(pool_df["竞店名称"] == shop) & (pool_df["竞品Sku_id"] == sku)]
            st = "pending" if len(m) else "none"
            pool_idx = int(m.index[0]) if len(m) else None
        else:
            st = "none"

    item = {
        "店铺": shop,
        "Sku_id": sku,
        "商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
        "品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
        "型号": _safe(row.get("型号")),
        "类目": _safe(row.get("类目三级_归一") or row.get("category3") or row.get("美团三级分类")),
        "价格": _safe(row.get("price") or row.get("到手价") or row.get("折后价")),
        "原价": _safe(row.get("original_price") or row.get("原价")),
        "月售": _clean_sales(row.get("sales") or row.get("月销量") or row.get("月售")),
        "条码": _safe(row.get("barcode") or row.get("条形码")),
        "规格": _safe(row.get("spec") or row.get("规格")),
        "图片": _safe(row.get("image_url") or row.get("图片url") or row.get("图片链接")),
        "状态": st,
        "自家SKUID": own_info.get("自家SKUID", ""),
        "自家商品名": own_info.get("自家商品名", ""),
        "匹配依据": own_info.get("匹配依据", ""),
        "pool_idx": pool_idx if st == "pending" else None,
    }
    # 粗对齐候选品组：凡带「候选自家品」JSON 的池行都直接在详情页展示（M6 反馈修复）。
    # 此前只认「品级候选(人工)」层——品级语义/关键字/向量层的候选被跳过，
    # 详情页只剩"去人工确认"跳转，候选要跑到池里才看得见（用户指出的麻烦）。
    spu_cands = []
    pool_layer_txt = ""
    CAND_LAYERS = {"品级候选(人工)", "品级语义(人工)", "关键字候选(人工)", "向量候选(人工)"}
    if pool_df is not None and len(pool_df):
        m = pool_df[(pool_df["竞店名称"] == shop) & (pool_df["竞品Sku_id"] == sku)]
        if len(m):
            pool_layer_txt = _safe(m.iloc[0].get("匹配层"))
        for pi, pr in m.iterrows():
            if str(pr.get("匹配层")) not in CAND_LAYERS:
                continue
            raw = _safe(pr.get("候选自家品"))
            if not raw:
                continue
            try:
                cands = json.loads(raw)[:5]
            except (ValueError, TypeError):
                continue
            layer_src = str(pr.get("匹配层") or "").replace("(人工)", "")
            sim_row = _safe(pr.get("相似度")) or _safe(pr.get("置信度"))
            for c in cands:
                if not isinstance(c, dict):
                    continue
                sim = _safe(c.get("相似度")) or _safe(pr.get("相似度")) or ""
                spu_cands.append({
                    "品组ID": _safe(c.get("品组ID")),
                    "品名": _safe(c.get("品名")),
                    "品牌": _safe(c.get("品牌")),
                    "型号": _safe(c.get("型号")),
                    "置信度": _safe(c.get("置信度")) or "",
                    "相似度": _safe(c.get("相似度")) or "",
                    "来源层": layer_src,
                    "行依据": _safe(pr.get("匹配依据")),
                    "规格数": _safe(c.get("规格数")),
                    "pool_idx": int(pi),
                })
        spu_cands = spu_cands[:8]
    item["匹配层文本"] = pool_layer_txt
    # 类目映射建议（品确认时顺带确认类目）
    cat_suggestion = _category_map_suggestion(_safe(row.get("类目三级_归一") or row.get("category3")))

    # 机器收敛行数据：类 / 品 / 规格 三层锁定状态与置信
    comp_cat = _safe(row.get("类目三级_归一") or row.get("category3"))
    conf_data = {
        "类": {"锁定": bool(cat_suggestion), "名称": comp_cat, "置信": "0.9+" if cat_suggestion else ""},
        "品": {"锁定": bool(spu_cands), "名称": spu_cands[0]["品名"] if spu_cands else "",
               "置信": spu_cands[0]["置信度"] if spu_cands and spu_cands[0]["置信度"] else ""},
        "规格": {"锁定": False, "名称": "", "置信": ""},
    }
    if pool_df is not None and len(pool_df):
        m = pool_df[(pool_df["竞店名称"] == shop) & (pool_df["竞品Sku_id"] == sku)
                    & (pool_df["匹配层"].astype(str).isin(["3", "4"]))]
        if len(m):
            conf_data["规格"] = {"锁定": True, "名称": _safe(m.iloc[0].get("自家商品名")),
                                 "置信": _safe(m.iloc[0].get("置信度"))}
    elif st == "aligned":
        conf_data["规格"] = {"锁定": True, "名称": item.get("自家商品名", ""), "置信": "已对应"}
    # 返回来源：from=coarse -> 粗对齐页；否则竞品总览
    from_page = (request.args.get("from") or "").strip()
    if from_page == "coarse":
        back_url = f"/coarse?shop={shop}"
    else:
        back_url = f"/browse?shop={shop}"
    # 已对应时补充自家品完整信息（价格/月售/条码/规格等，供一键查看与复制）
    if st == "aligned" and item["自家SKUID"]:
        oi = get_own_info(item["自家SKUID"])
        item["自家售价"] = oi.get("售价", "")
        item["自家采购价"] = oi.get("采购价", "")
        item["自家月销"] = oi.get("月销量", "")
        item["自家条码"] = oi.get("商品条码", "")
        item["自家规格"] = oi.get("规格名称", "")
        item["自家有图"] = oi.get("是否有图片", "")
        item["自家品牌"] = oi.get("品牌名称", "")
        item["自家SPUID"] = oi.get("SPUID", "")
        item["自家图片url"] = oi.get("商品图片url", "")
        # P4.3 补货参考回填列（口径见 own_sales_meta.json）
        item["自家7天销量"] = oi.get("7天总销量", "")
        item["自家可用库存"] = oi.get("可用库存", "")
        _s7 = _extract_num(item["自家7天销量"])
        _s30 = _extract_num(item["自家月销"])
        item["自家趋势"] = ""
        if _s7 is not None and _s30 is not None and _s30 > 0 and _s7 > 0:
            _r = (_s7 / 7.0) / (_s30 / 30.0)
            item["自家趋势"] = "↑热" if _r >= 1.2 else ("↓冷" if _r <= 0.8 else "→平")
        elif _s7 is not None and _s7 > 0:
            item["自家趋势"] = "↑热"
    # 下一个待处理（同店：当前 sku 之后的第一个 待处理/已粗对齐 竞品；无则循环取第一个）
    next_sku = ""
    next_name = ""
    try:
        lk = competitor_lookup.get(shop, {})
        sku_list = list(lk.keys())
        if sku in sku_list:
            i0 = sku_list.index(sku)
            cand = sku_list[i0 + 1:] + sku_list[:i0]
            for sk in cand:
                lv = _converge_level(shop, sk, lk.get(sk))
                if lv in ("cat", "none", "spu"):
                    next_sku = sk
                    next_name = _safe((lk.get(sk) or {}).get("name") or (lk.get(sk) or {}).get("商品名字"))
                    break
    except Exception:
        pass
    # P1.2：该 SKU 的历史观测轨迹（历次快照 价格/月售/库存，时间升序；≥2 期才有对比意义）
    hist_rows = []
    _h = _sku_history(shop, sku)
    for _i, (_t, _p, _s, _k) in enumerate(_h):
        _prev = _h[_i - 1] if _i > 0 else None
        _d = {}
        if _prev is not None:
            if _p is not None and _prev[1] is not None and _prev[1] != _p:
                _d["price"] = round(_p - _prev[1], 1)
            if _s is not None and _prev[2] is not None and _prev[2] != _s:
                _d["sales"] = _s - _prev[2]
            if _k is not None and _prev[3] is not None and _prev[3] != _k:
                _d["stock"] = _k - _prev[3]
        hist_rows.append({"t": _t[:16], "price": _p, "sales": _s, "stock": _k, "delta": _d})
    item["hist_rows"] = hist_rows
    item["hist_n"] = len(hist_rows)

    return render_template("comp_detail.html", item=item, spu_cands=spu_cands,
                           back_url=back_url, from_page=from_page,
                           cat_suggestion=cat_suggestion, conf_data=conf_data,
                           next_sku=next_sku, next_name=next_name)


@app.route("/matrix")
def matrix():
    """多店横比：以竞品品为单位，横向列出各店同款竞品（价格/原价/月售/图片/状态），
    支持自家对齐联动（同自家SKU/同自家品组优先），退化到 同竞品品ID / 同商品名。
    入口：竞品详情页「多店横比」按钮。"""
    shop = (request.args.get("shop") or "").strip()
    sku = (request.args.get("sku") or "").strip()
    mode = (request.args.get("mode") or "auto").strip()
    if mode not in ("auto", "pid", "name"):
        mode = "auto"
    src_row, src_pid, src_name = None, "", ""
    own_sku = own_spu = ""
    if shop in competitor_lookup and sku in competitor_lookup[shop]:
        src_row = competitor_lookup[shop][sku]
        src_pid = _safe(src_row.get("comp_id") or src_row.get("商品ID"))
        src_name = _safe(src_row.get("name") or src_row.get("商品名字") or src_row.get("清洗后商品名"))
        am = aligned_map.get(sku, {})
        own_sku = _safe(am.get("自家SKUID"))
        oi = get_own_info(own_sku) if own_sku else {}
        own_spu = _safe(oi.get("SPUID")) or own_spu_map.get(own_sku, "")
    # 汇聚各店同款竞品（全店全 SKU 一次遍历）
    matches = []
    if src_row is not None:
        for s, lk in competitor_lookup.items():
            for s2, row in lk.items():
                pid = _safe(row.get("comp_id") or row.get("商品ID"))
                nm = _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名"))
                am2 = aligned_map.get(s2, {})
                o2 = _safe(am2.get("自家SKUID"))
                oi2 = get_own_info(o2) if o2 else {}
                o2spu = _safe(oi2.get("SPUID")) or own_spu_map.get(o2, "")
                hit = False
                if mode == "pid":
                    hit = bool(src_pid and pid and pid == src_pid)
                elif mode == "name":
                    hit = bool(src_name and nm and nm == src_name)
                else:
                    hit = bool(
                        (own_spu and o2spu and o2spu == own_spu)
                        or (own_sku and o2 and o2 == own_sku)
                        or (src_pid and pid and pid == src_pid)
                        or (src_name and nm and nm == src_name)
                    )
                if not hit:
                    continue
                lvl = _converge_level(s, s2, row)
                matches.append({
                    "店铺": s, "Sku_id": s2,
                    "商品名": nm,
                    "品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
                    "型号": _safe(row.get("型号")),
                    "类目": _safe(row.get("类目三级_归一") or row.get("category3")),
                    "价格": _safe(row.get("price") or row.get("到手价") or row.get("折后价")),
                    "原价": _safe(row.get("original_price") or row.get("原价")),
                    "月售": _clean_sales(row.get("sales") or row.get("月销量") or row.get("月售")),
                    "条码": _safe(row.get("barcode") or row.get("条形码")),
                    "图片": _safe(row.get("image_url") or row.get("图片url") or row.get("图片链接")),
                    "状态": lvl,
                    "own_name": _safe(oi2.get("商品名称")),
                    "own_price": _safe(oi2.get("售价")),
                })
    # 月售高的在前（分档：真实成交 > 意向 > 无数据）
    _real, _intent, _empty = [], [], []
    for m in matches:
        s = _clean_sales(m["月售"])
        if not s:
            _empty.append(m)
        elif "(意向)" in s:
            _intent.append(m)
        else:
            _real.append(m)
    def _sn(m):
        try:
            return float(_clean_sales(m["月售"]).replace(",", "").replace("(意向)", ""))
        except ValueError:
            return 0
    _real.sort(key=_sn, reverse=True)
    _intent.sort(key=_sn, reverse=True)
    matches = _real + _intent + _empty
    # 聚合统计：覆盖店数 / 均价 / 最低价 / 最高价 / 月售合计
    agg = {"shops": len({m["店铺"] for m in matches}), "avg": "", "min": "", "max": "", "sum_sales": ""}
    prices = [_extract_num(m["价格"]) for m in matches if _extract_num(m["价格"]) is not None]
    sales = [_extract_num(m["月售"]) for m in matches
             if _extract_num(m["月售"]) is not None and "(意向)" not in m["月售"]]
    if prices:
        agg["avg"] = round(sum(prices) / len(prices), 2)
        agg["min"] = min(prices)
        agg["max"] = max(prices)
    if sales:
        agg["sum_sales"] = sum(int(x) for x in sales)
    src_cat = _safe(src_row.get("类目三级_归一") or src_row.get("category3")) if src_row is not None else ""
    return render_template("matrix.html", shop=shop, sku=sku, mode=mode,
                           src_row=src_row, src_pid=src_pid, src_name=src_name,
                           own_sku=own_sku, own_spu=own_spu, src_cat=src_cat,
                           matches=matches, n=len(matches), agg=agg,
                           shop_count=len(competitor_lookup))


@app.route("/proxy_img")
def proxy_img():
    """竞品图片本地代理缓存：外站直链 → 下载到 output/img_cache 后返回，加速重复加载、规避防盗链失败。"""
    url = (request.args.get("url") or "").strip()
    if not url.startswith(("http://", "https://")):
        return "", 204
    os.makedirs(IMG_CACHE_DIR, exist_ok=True)
    try:
        path = urlparse(url).path
    except Exception:
        path = ""
    ext = os.path.splitext(path)[1][:8].lower()
    if ext not in (".jpg", ".jpeg", ".png", ".webp", ".gif", ".bmp"):
        ext = ".jpg"
    fname = hashlib.md5(url.encode("utf-8")).hexdigest() + ext
    fpath = os.path.join(IMG_CACHE_DIR, fname)
    if os.path.exists(fpath) and os.path.getsize(fpath) > 0:
        return send_from_directory(IMG_CACHE_DIR, fname, max_age=86400)
    try:
        import requests as _req
        resp = _req.get(url, timeout=6, verify=False, stream=True,
                        headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/120.0"})
        if resp.status_code != 200:
            return "", 204
        ctype = resp.headers.get("Content-Type", "")
        if ctype and not ctype.startswith("image/"):
            return "", 204
        with open(fpath, "wb") as f:
            for chunk in resp.iter_content(65536):
                if chunk:
                    f.write(chunk)
        return send_from_directory(IMG_CACHE_DIR, fname, max_age=86400)
    except Exception:
        if os.path.exists(fpath):
            try:
                os.remove(fpath)
            except Exception:
                pass
        return "", 204


def _confirm_category_map(comp_cat, own_cat):
    """品确认时顺带确认类目映射：竞品类目 -> 自家类目 写回类目映射表（重跑归一化生效）。
    同时立即更新内存缓存 cat_map_cache，总览/收敛层级即时生效（无需重启面板）。
    等值映射（竞品类目==自家类目）也写入缓存：收敛层级"类已锁定"判定依赖 cat_map_cache。"""
    if not comp_cat or not own_cat:
        return
    cat_map_cache[comp_cat] = own_cat
    if comp_cat == own_cat:
        return
    try:
        if not os.path.exists(CATEGORY_MAP_FILE):
            return
        df = pd.read_excel(CATEGORY_MAP_FILE, dtype=str).fillna("")
        if "三级类目" not in df.columns or "建议映射" not in df.columns:
            return
        mask = df["三级类目"].astype(str) == comp_cat
        if mask.any():
            df.loc[mask, "建议映射"] = own_cat
        else:
            df = pd.concat([df, pd.DataFrame([{"三级类目": comp_cat, "竞品计数": "", "自家计数": "",
                                               "是否一致": "仅竞品", "建议映射": own_cat}])],
                           ignore_index=True)
        df.to_excel(CATEGORY_MAP_FILE, index=False)
    except Exception:
        pass


def _category_map_suggestion(comp_cat):
    """查类目映射建议：竞品类目 -> 映射表建议自家类目（无则空）。用缓存。"""
    if not comp_cat:
        return ""
    return cat_map_cache.get(comp_cat, "")


def _converge_level(shop, sku, row=None):
    """收敛层级：sku=规格已锁定(已对应) / spu=品已锁定(已粗对齐) / cat=类已锁定 / none=未收敛。
    row 为竞品行 dict（用于取竞品品ID/类目）；缺省时尽量推断。结果带惰性缓存（_invalidate_stats 失效）。"""
    rec = progress.get("confirmed", {}).get(shop, {}).get(sku)
    if sku in aligned_map or (rec and not rec.get("spu_only")):
        return "sku"
    if rec and rec.get("spu_only"):
        return "spu"
    pid = ""
    cat = ""
    if row is not None:
        pid = _safe(row.get("comp_id") or row.get("商品ID"))
        cat = _safe(row.get("类目三级_归一") or row.get("category3"))
    ck = (shop, sku, pid, cat)
    v = _level_cache.get(ck)
    if v is not None:
        return v
    if pid and (shop, pid) in spu_layer_keys:
        v = "spu"
    elif cat and cat in cat_map_cache:
        v = "cat"
    else:
        v = "none"
    _level_cache[ck] = v
    return v


def _aggregate_spu(shop, comp_sku, own_sku, own_name, spu_gid=""):
    """跨店聚合：品确认后，其他店同款（同名或同品ID精确匹配）竞品批量写入 SPU 层品对。
    返回聚合关联的竞品数。保守规则：同名精确 或 同竞品品ID，不做模糊聚合。"""
    src = competitor_lookup.get(shop, {}).get(comp_sku)
    if src is None:
        return 0
    src_pid = _safe(src.get("comp_id") or src.get("商品ID"))
    src_name = _safe(src.get("name") or src.get("商品名字") or src.get("清洗后商品名"))
    if not src_pid and not src_name:
        return 0
    oi = own_lookup.get(_safe(own_sku), {})
    spu_gid = spu_gid or own_spu_map.get(_safe(own_sku), "")
    agg = 0
    try:
        from summary_library import load_library as _ll, save_library as _sl, upsert_manual as _um
        spu_df, sku_df = _ll()
        for s, lk in competitor_lookup.items():
            if s == shop:
                continue
            for sku, row in lk.items():
                if sku in aligned_map or sku in progress.get("confirmed", {}).get(s, {}):
                    continue
                pid = _safe(row.get("comp_id") or row.get("商品ID"))
                nm = _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名"))
                if not ((src_pid and pid == src_pid) or (src_name and nm == src_name)):
                    continue
                spu_df = _um(spu_df, "竞品品ID", s, {
                    "竞品品ID": pid, "竞品商品名": nm,
                    "竞品品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
                    "竞品型号": _safe(row.get("型号")),
                    "竞品类目": _safe(row.get("类目三级_归一") or row.get("category3")),
                    "自家品组ID": spu_gid, "自家品名": own_name, "置信度": "",
                })
                progress.setdefault("confirmed", {}).setdefault(s, {})[sku] = {
                    "sku_id": own_sku, "spu_only": True, "agg": True,
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
                agg += 1
        if agg:
            _sl(spu_df, sku_df)
            save_progress()
    except Exception:
        pass
    return agg


@app.route("/api/rename_shop", methods=["POST"])
def api_rename_shop():
    """店铺重命名：同步 标准库文件/试跑目录/多竞店汇总库/两层汇总库/progress/profile。"""
    data = request.get_json(force=True)
    old = (data.get("old_name") or "").strip()
    new = (data.get("new_name") or "").strip()
    if not old or old not in competitor_lookup:
        return jsonify({"ok": False, "msg": "原店铺不存在"}), 400
    if not new:
        return jsonify({"ok": False, "msg": "新店名不能为空"}), 400
    if new in competitor_lookup:
        return jsonify({"ok": False, "msg": "新店名与已有店铺冲突"}), 400

    done = []
    try:
        # 1) 标准库文件
        std_old = os.path.join(OUTPUT_DIR, "竞店标准库", f"竞品_{old}_标准化.xlsx")
        std_new = os.path.join(OUTPUT_DIR, "竞店标准库", f"竞品_{new}_标准化.xlsx")
        if os.path.exists(std_old):
            os.rename(std_old, std_new)
            done.append("标准库文件")
        # 2) archive 快照
        arch_dir = os.path.join(OUTPUT_DIR, "竞店标准库", "archive")
        if os.path.isdir(arch_dir):
            for f in os.listdir(arch_dir):
                if f.startswith(old + "_"):
                    os.rename(os.path.join(arch_dir, f),
                              os.path.join(arch_dir, new + f[len(old):]))
            done.append("archive快照")
        # 3) 试跑目录
        trial_old = os.path.join(OUTPUT_DIR, f"试跑_{old}")
        trial_new = os.path.join(OUTPUT_DIR, f"试跑_{new}")
        if os.path.isdir(trial_old):
            os.rename(trial_old, trial_new)
            done.append("试跑目录")
        # 4) 多竞店汇总库
        sum_f = os.path.join(OUTPUT_DIR, "竞店标准库", "多竞店汇总库.xlsx")
        if os.path.exists(sum_f):
            sdf = pd.read_excel(sum_f, dtype=str)
            col = "竞店名称" if "竞店名称" in sdf.columns else "shop_name"
            if col in sdf.columns:
                sdf[col] = sdf[col].replace(old, new)
                sdf.to_excel(sum_f, index=False)
                done.append("多竞店汇总库")
        # 5) 两层汇总库
        from summary_library import load_library as _ll, save_library as _sl
        spu_df, sku_df = _ll()
        for dfx in (spu_df, sku_df):
            if len(dfx) and "竞店名称" in dfx.columns:
                dfx["竞店名称"] = dfx["竞店名称"].replace(old, new)
        _sl(spu_df, sku_df)
        done.append("两层汇总库")
        # 5b) 竞店数据质量报告（管理栏/选店分析页的店名来源）
        rep_f = os.path.join(OUTPUT_DIR, "竞店标准库", "竞店数据质量报告.xlsx")
        if os.path.exists(rep_f):
            rdf = pd.read_excel(rep_f, dtype=str)
            if "店名" in rdf.columns:
                rdf["店名"] = rdf["店名"].replace(old, new)
                rdf.to_excel(rep_f, index=False)
                done.append("质量报告")
        # 6) progress.json
        for kind in ("confirmed", "no_match"):
            m = progress.get(kind, {})
            if old in m:
                m[new] = m.pop(old)
        save_progress()
        done.append("进度记录")
        # 7) profile
        prof_old = os.path.join(SCRIPTS_DIR, "..", "config", "import_profiles", f"{old}.yaml")
        prof_old = os.path.normpath(prof_old)
        prof_new = os.path.join(os.path.dirname(prof_old), f"{new}.yaml")
        if os.path.exists(prof_old):
            os.rename(prof_old, prof_new)
            done.append("导入配置")
        # 8) 提取产物（output/竞品_标准化_<old>_*.xlsx）
        for f in os.listdir(OUTPUT_DIR):
            if f.startswith("竞品_标准化_" + old + "_"):
                os.rename(os.path.join(OUTPUT_DIR, f),
                          os.path.join(OUTPUT_DIR, f.replace(old, new, 1)))
    except Exception as e:
        return jsonify({"ok": False, "msg": "重命名中途失败（可能部分完成）: " + str(e)}), 500

    # 重载数据
    load_data()
    load_progress()
    _invalidate_stats()          # 旧店名缓存键全部作废
    _warmup_all_shops_bg()       # 后台重建（新店名），避免用户首次访问同步重算
    return jsonify({"ok": True, "msg": f"已改名：{old} → {new}（{'、'.join(done)}）"})


@app.route("/api/own_detail")
def api_own_detail():
    """自家品完整信息（手动对应时点击搜索结果展示详情比对用）。"""
    sku = request.args.get("sku", "").strip()
    info = get_own_info(sku)
    if not info:
        return jsonify({"ok": False, "msg": "自家品不存在"}), 404
    return jsonify({"ok": True, "info": info})


@app.route("/api/manual_confirm", methods=["POST"])
def api_manual_confirm():
    """手动对应：给任意竞品（含暂未匹配/无对应）指定自家品，写两层汇总库人工行+进度。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    sku = (data.get("sku") or "").strip()
    own_sku = (data.get("own_sku") or "").strip()
    note = data.get("note", "")

    if shop not in competitor_lookup or not sku:
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    row = competitor_lookup.get(shop, {}).get(sku)
    if row is None:
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    own_info = get_own_info(own_sku)
    if not own_info:
        return jsonify({"ok": False, "msg": "自家品不存在: " + own_sku}), 400

    # SKU层人工行
    sku_row = {
        "竞品品ID": _safe(row.get("comp_id") or row.get("商品ID")),
        "竞品Sku_id": sku,
        "竞品商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
        "竞品条形码": _safe(row.get("barcode") or row.get("条形码")),
        "自家SPUID": own_info.get("SPUID", ""),
        "自家SKUID": own_info.get("SKUID", ""),
        "自家商品名": own_info.get("商品名称", ""),
        "匹配层": "手动对应",
        "匹配依据": "人工手动指定",
        "置信度": "",
    }
    # SPU层人工行
    spu_gid = own_spu_map.get(own_info.get("SKUID", ""), "")
    spu_row = {
        "竞品品ID": _safe(row.get("comp_id") or row.get("商品ID")),
        "竞品商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
        "竞品品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
        "竞品型号": _safe(row.get("型号")),
        "竞品类目": _safe(row.get("类目三级_归一") or row.get("category3")),
        "自家品组ID": spu_gid or own_info.get("SPUID", ""),
        "自家品名": own_info.get("商品名称", ""),
        "置信度": "",
    }
    from summary_library import load_library as _ll, save_library as _sl, upsert_manual as _um
    spu_df, sku_df = _ll()
    spu_df = _um(spu_df, "竞品品ID", shop, spu_row)
    sku_df = _um(sku_df, "竞品Sku_id", shop, sku_row)
    _sl(spu_df, sku_df)

    # 类目映射顺带确认：竞品类目 -> 自家类目，写回映射表
    try:
        own_cat = _safe(own_info.get("类目归一")) or _safe(own_info.get("类目三级"))
        comp_cat = _safe(row.get("类目三级_归一") or row.get("category3"))
        if own_cat and comp_cat:
            _confirm_category_map(comp_cat, own_cat)
    except Exception:
        pass

    # 学习闭环：手动对应样本（含品牌补全）+ 关键词反哺
    try:
        from learn import record as learn_record, add_keyword as learn_add_kw
        learn_record(shop, sku, _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
                     own_info.get("SKUID", ""), own_info.get("商品名称", ""), source="手动对应",
                     comp_brand=_safe(row.get("品牌（抽取）") or row.get("品牌名称")),
                     comp_model=_safe(row.get("型号")),
                     own_brand=own_info.get("品牌名称", ""), own_model=own_info.get("型号", ""))
        learn_add_kw(own_info.get("商品名称", ""), own_info.get("SPUID", ""),
                     own_info.get("类目三级_归一", "") or own_info.get("美团闪购类目三级", ""))
    except Exception:
        pass

    # 进度记录
    progress.setdefault("confirmed", {}).setdefault(shop, {})[sku] = {
        "sku_id": own_sku,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "manual": True,
    }
    progress.setdefault("no_match", {}).get(shop, {}).pop(sku, None)
    save_progress()

    # 刷新内存映射（aligned_map 供总览状态）
    aligned_map[sku] = {"自家SKUID": own_info.get("SKUID", ""),
                        "自家商品名": own_info.get("商品名称", ""),
                        "匹配层": "手动对应", "匹配依据": "人工手动指定"}
    # 跨店聚合：其他店同款竞品一并写入 SPU 层
    agg_n = 0
    try:
        agg_n = _aggregate_spu(shop, sku, own_sku, own_info.get("商品名称", ""))
    except Exception:
        pass
    return jsonify({"ok": True, "msg": f"已手动对应：{sku} -> {own_sku}" +
                    (f"，并关联 {agg_n} 家店同款" if agg_n else "")})


def _purge_price_hist(shop):
    """把该店的价格观测行从 append-only 历史里摘掉（删店后不留幽灵店）。

    历史文件只追加不覆盖是常态，但删店是用户显式确认的动作：留着已删店的数据，
    首页"变化摘要"与详情页"历史轨迹"会把幽灵店当成有效观测。改写前先落一份带
    时间戳的备份，且原子替换——清错也能还原。
    """
    import time as _tt
    f = os.path.join(OUTPUT_DIR, "竞店标准库", "价格观测历史.jsonl")
    if not os.path.exists(f) or not shop:
        return 0, ""
    # 保险丝一：拿不到在册清单就绝不动文件。
    # （今天真实踩过：脚本未 load_data → 在册清单为空 → 五家在册店全被判幽灵店，历史被清空）
    registered = set(competitor_lookup.keys())
    if not registered:
        print("[删店清理] 在册店清单为空，拒绝改写观测历史")
        return 0, ""
    # 保险丝二：店还在册说明不是"删店"场景，拒绝
    if shop in registered:
        print("[删店清理] 该店仍在册，拒绝清理:", shop)
        return 0, ""
    NL = chr(10)
    keep, drop, total = [], 0, 0
    with open(f, encoding="utf-8") as fh:
        for line in fh:
            t = line.strip()
            if not t:
                continue
            total += 1
            try:
                rec = json.loads(t)
            except Exception:
                keep.append(line.rstrip(NL))
                continue
            if str(rec.get("店") or "") == shop:
                drop += 1
            else:
                keep.append(t)
    if not drop:
        return 0, ""
    # 保险丝三：单次清掉超过两成行数属异常（正常删一家店远小于此），中止不改写
    if drop > total * 0.2:
        print(f"[删店清理] 异常：{shop} 占观测历史 {drop}/{total} 行，拒绝改写")
        return 0, ""
    base = f + ".bak_删店_" + _tt.strftime("%Y%m%d_%H%M%S_%f")
    bak = base
    seq = 0
    while os.path.exists(bak):        # 备份名必须唯一：同秒覆盖会毁掉唯一可还原副本
        seq += 1
        bak = "%s-%d" % (base, seq)
    shutil.copy2(f, bak)
    tmp = f + ".tmp"
    with open(tmp, "w", encoding="utf-8") as fh:
        fh.write(NL.join(keep) + (NL if keep else ""))
    os.replace(tmp, f)
    # 观测缓存按 mtime 自失效；这里显式重置，删店后首个请求即读到干净历史
    globals()["_price_hist_cache"] = None
    _hist_full_cache["mtime"] = None
    _hist_full_cache["data"] = None
    return drop, os.path.basename(bak)


@app.route("/api/delete_shop", methods=["POST"])
def api_delete_shop():
    """删除店铺（含标准库/试跑目录/汇总库/进度/配置），需 confirm=true 二次确认。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    if data.get("confirm") is not True:
        return jsonify({"ok": False, "msg": "请确认删除（confirm=true）"}), 400
    if shop not in competitor_lookup:
        return jsonify({"ok": False, "msg": "店铺不存在"}), 400

    done = []
    try:
        # 试跑目录
        trial = os.path.join(OUTPUT_DIR, f"试跑_{shop}")
        if os.path.isdir(trial):
            shutil.rmtree(trial)
            done.append("试跑目录")
        # 标准库文件 + archive 快照
        lib_dir = os.path.join(OUTPUT_DIR, "竞店标准库")
        std_f = os.path.join(lib_dir, f"竞品_{shop}_标准化.xlsx")
        if os.path.exists(std_f):
            os.remove(std_f)
            done.append("标准库文件")
        arch = os.path.join(lib_dir, "archive")
        if os.path.isdir(arch):
            for f in os.listdir(arch):
                if f.startswith(shop + "_"):
                    os.remove(os.path.join(arch, f))
            done.append("archive快照")
        # 多竞店汇总库：删除该店行
        sum_f = os.path.join(lib_dir, "多竞店汇总库.xlsx")
        if os.path.exists(sum_f):
            sdf = pd.read_excel(sum_f, dtype=str)
            col = "竞店名称" if "竞店名称" in sdf.columns else "shop_name"
            if col in sdf.columns:
                sdf = sdf[sdf[col] != shop]
                if len(sdf):
                    sdf.to_excel(sum_f, index=False)
                else:
                    os.remove(sum_f)
            done.append("多竞店汇总库")
        # 竞店数据质量报告：删除该店行（管理栏/选店分析页的店列表来源）
        rep_f = os.path.join(lib_dir, "竞店数据质量报告.xlsx")
        if os.path.exists(rep_f):
            rdf = pd.read_excel(rep_f, dtype=str)
            if "店名" in rdf.columns:
                rdf = rdf[rdf["店名"] != shop]
                if len(rdf):
                    rdf.to_excel(rep_f, index=False)
                else:
                    os.remove(rep_f)
                done.append("质量报告")
        # 两层汇总库：删除该店行
        from summary_library import load_library as _ll, save_library as _sl
        spu_df, sku_df = _ll()
        spu_df = spu_df[spu_df["竞店名称"] != shop] if len(spu_df) else spu_df
        sku_df = sku_df[sku_df["竞店名称"] != shop] if len(sku_df) else sku_df
        _sl(spu_df, sku_df)
        done.append("两层汇总库")
        # progress
        for kind in ("confirmed", "no_match"):
            progress.get(kind, {}).pop(shop, None)
        save_progress()
        done.append("进度记录")
        # profile
        prof = os.path.join(BASE_DIR, "config", "import_profiles", f"{shop}.yaml")
        if os.path.exists(prof):
            os.remove(prof)
            done.append("导入配置")
        # 提取产物
        for f in os.listdir(OUTPUT_DIR):
            if f.startswith("竞品_标准化_" + shop + "_"):
                os.remove(os.path.join(OUTPUT_DIR, f))
        # 价格观测历史（append-only）：不清就会留幽灵店——变动摘要/历史轨迹会读到已删店
        n_purged, bak = _purge_price_hist(shop)
        if n_purged:
            done.append(f"价格观测 {n_purged} 行（备份 {bak}）")
    except Exception as e:
        return jsonify({"ok": False, "msg": "删除中途失败（可能部分完成）: " + str(e)}), 500

    load_data()
    load_progress()
    _invalidate_stats()          # 被删店的缓存键残留清理
    _warmup_all_shops_bg()       # 剩余店后台重建（删店清了全部店缓存）
    return jsonify({"ok": True, "msg": f"已删除店铺：{shop}（{'、'.join(done)}）"})


@app.route("/api/coarse_queue")
def api_coarse_queue():
    """粗对齐（品级）队列：待确认 = 该店全部未对齐竞品；已确认 = 汇总库SPU层品对品。"""
    shop = (request.args.get("shop") or "").strip()
    q = (request.args.get("q") or "").strip()
    brand = (request.args.get("brand") or "").strip()
    cat = (request.args.get("cat") or "").strip()
    page = request.args.get("page", 1, type=int)
    if shop not in competitor_lookup:
        return jsonify({"ok": False, "msg": "店铺不存在"}), 400

    conf_set = set(progress.get("confirmed", {}).get(shop, {}).keys())
    nm_set = set(progress.get("no_match", {}).get(shop, {}).keys())

    # 待确认：未对齐（不在 aligned_map/confirmed/no_match）的竞品
    # 候选信息：从人工池的品级候选行取（竞品Sku_id -> 候选品组JSON top1）
    spu_cand = {}
    if pool_df is not None and len(pool_df):
        m = pool_df[(pool_df["竞店名称"] == shop) &
                    (pool_df["匹配层"].astype(str) == "品级候选(人工)")]
        for _, r in m.iterrows():
            raw = _safe(r.get("候选自家品"))
            if raw:
                try:
                    cands = json.loads(raw)
                    if cands:
                        c0 = cands[0]
                        spu_cand[_safe(r.get("竞品Sku_id"))] = {
                            "品组ID": _safe(c0.get("品组ID")),
                            "品名": _safe(c0.get("品名")),
                            "置信度": _safe(c0.get("置信度")),
                        }
                except (ValueError, TypeError):
                    pass

    rows = []
    for sku, row in (competitor_lookup.get(shop) or {}).items():
        if sku in conf_set or sku in aligned_map or sku in nm_set:
            continue
        cand = spu_cand.get(sku, {})
        rows.append({
            "Sku_id": sku,
            "商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
            "品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
            "型号": _safe(row.get("型号")),
            "类目": _safe(row.get("类目三级_归一") or row.get("category3") or row.get("美团三级分类")),
            "价格": _safe(row.get("price") or row.get("到手价") or row.get("折后价")),
            "月售": _clean_sales(row.get("sales") or row.get("月销量") or row.get("月售")),
            "图片": _safe(row.get("image_url") or row.get("图片url") or row.get("图片链接")),
            "候选品组": cand.get("品名", ""),
            "候选品组ID": cand.get("品组ID", ""),
            "候选置信度": cand.get("置信度", ""),
        })

    # 筛选
    if q:
        from search_util import kw_hit
        rows = [r for r in rows if kw_hit(r, q, ["商品名", "Sku_id", "品牌"])]
    if brand:
        rows = [r for r in rows if r["品牌"] == brand]
    if cat:
        rows = [r for r in rows if r["类目"] == cat]

    total = len(rows)
    total_pages = max(1, (total + 50 - 1) // 50)
    page = max(1, min(page, total_pages))
    items = rows[(page - 1) * 50: page * 50]

    # 已确认：汇总库 SPU 层该店品对品
    spu_df, _ = load_library()
    aligned = []
    if len(spu_df):
        for _, r in spu_df[spu_df["竞店名称"] == shop].iterrows():
            aligned.append({
                "竞品品ID": _safe(r.get("竞品品ID")),
                "竞品商品名": _safe(r.get("竞品商品名")),
                "自家品组ID": _safe(r.get("自家品组ID")),
                "自家品名": _safe(r.get("自家品名")),
                "置信度": _safe(r.get("置信度")),
                "确认来源": _safe(r.get("确认来源")),
            })

    brands = sorted({r["品牌"] for r in rows if r["品牌"]})
    cats = sorted({r["类目"] for r in rows if r["类目"]})
    resp = {
        "ok": True,
        "shop": shop,
        "items": items,
        "total": total,
        "total_pages": total_pages,
        "page": page,
        "aligned": aligned,
        "aligned_count": len(aligned),
        "brands": brands,
        "cats": cats,
    }
    # 批量大模型判断用：返回筛选范围全部 Sku_id（不分页）
    if request.args.get("all_skus") == "1":
        resp["skus"] = [r["Sku_id"] for r in rows]
    return jsonify(resp)


@app.route("/coarse")
def coarse_page():
    return redirect("/browse")


@app.route("/tools")
def tools_page():
    """工具页：数据导入 / 选店分析·AI / 反向查询 / 分区管理 / 学习管理（低频功能统一收纳）。"""
    shops = list(competitor_lookup.keys())
    zones = []
    own_stores = []
    try:
        zones = _zl.load_zones()
    except Exception:
        zones = []
    try:
        import own_store_lib
        own_stores = own_store_lib.list_stores()
    except Exception:
        own_stores = []
    # 店铺归属映射（收纳语义：竞店 → 所在区），供区管理面板标注
    shop_zone = {}
    for z in zones:
        for s in z.get("comp_shops", []):
            shop_zone[s] = z.get("name", "")
    return render_template("tools.html", shops=shops,
                           shop_count=len(shops), zones=zones,
                           all_shops=shops, own_stores=own_stores,
                           shop_zone=shop_zone)


@app.route("/api/own_store/list")
def api_own_store_list():
    """自家门店列表（多自家门店比对）。"""
    try:
        import own_store_lib
        return jsonify({"ok": True, "stores": own_store_lib.list_stores()})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e), "stores": []})


@app.route("/api/ai/match_judge", methods=["POST"])
def api_ai_match_judge():
    """大模型辅助判断：竞品与候选自家品是否为同一品（可选按钮，点才调用）。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    sku = (data.get("sku") or "").strip()
    if shop not in competitor_lookup or not sku:
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    row = competitor_lookup.get(shop, {}).get(sku)
    if row is None:
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400

    from ai_match import judge as ai_judge, cache_get, cache_set, load_examples

    # 缓存命中直接返回（避免重复花钱）
    cached = cache_get(shop, sku)
    if cached:
        return jsonify({"ok": True, "cached": True, "result": cached})

    comp_name = _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名"))
    comp_brand = _safe(row.get("品牌（抽取）") or row.get("品牌名称"))
    comp_model = _safe(row.get("型号"))
    comp_cat = _safe(row.get("类目三级_归一") or row.get("category3"))

    # 候选自家品：优先用品级候选 JSON；无则类目一致的自家品名 top3
    candidates = []
    if pool_df is not None and len(pool_df):
        m = pool_df[(pool_df["竞店名称"] == shop) & (pool_df["竞品Sku_id"] == sku)]
        for _, pr in m.iterrows():
            raw = _safe(pr.get("候选自家品"))
            if raw:
                try:
                    for c in json.loads(raw)[:3]:
                        if c.get("品名"):
                            candidates.append(_safe(c.get("品名")))
                except (ValueError, TypeError):
                    pass
    if not candidates and own_df is not None:
        cat_col = "类目三级_归一" if "类目三级_归一" in own_df.columns else "美团闪购类目三级"
        if cat_col in own_df.columns and comp_cat:
            sub = own_df[own_df[cat_col].astype(str) == comp_cat]
            for _, orow in sub.head(3).iterrows():
                nm = _safe(orow.get("商品名称"))
                if nm:
                    candidates.append(nm)

    try:
        result = ai_judge(comp_name, comp_brand, comp_model, candidates,
                          examples=load_examples())
        # 建议自家品名称 -> SKUID（供一键采用）
        if result.get("suggest_own") and own_df is not None:
            name_col = "商品名称" if "商品名称" in own_df.columns else None
            if name_col:
                hit = own_df[own_df[name_col].astype(str) == str(result["suggest_own"]).strip()]
                if len(hit):
                    result["suggest_sku"] = _safe(hit.iloc[0].get("SKUID"))
        result.setdefault("suggest_sku", "")
        cache_set(shop, sku, result)
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/ai/prescreen", methods=["POST"])
def api_ai_prescreen():
    """AI 批量预筛：后台线程跑该店人工池待处理条目。
    annotate 模式只标注徽章；triage 模式预筛完自动分流落地（AI预筛行/无对应标记，均可撤销）。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    if not shop:
        return jsonify({"ok": False, "msg": "缺少 shop 参数"}), 400
    try:
        limit = int(data.get("limit") or 100)
    except (ValueError, TypeError):
        limit = 100
    limit = max(1, min(limit, 1000))
    if _PRESCREEN_STATE["running"]:
        return jsonify({"ok": False, "msg": f"预筛进行中（{_PRESCREEN_STATE['shop']} "
                           f"{_PRESCREEN_STATE['done']}/{_PRESCREEN_STATE['total']}），请稍候"}), 409
    # key 预检：未配置/演示模式直接提示，不空跑
    from ai_client import load_config
    cfg = load_config()
    if str(cfg.get("driver") or "real").lower() != "fake" and not (cfg.get("api_key") or "").strip():
        return jsonify({"ok": False, "msg": "未配置 AI 接口密钥：点工具栏「🔑 AI接口」粘贴 DeepSeek key 并测试连通"}), 400
    threading.Thread(target=_run_prescreen_bg, args=(shop, limit), daemon=True).start()
    return jsonify({"ok": True, "msg": f"预筛已启动（{shop}，最多 {limit} 条，后台执行）"})


@app.route("/api/ai/prescreen/status")
def api_ai_prescreen_status():
    """预筛任务状态 + 该店结论分布（前端轮询）。"""
    shop = (request.args.get("shop") or "").strip()
    st = {k: _PRESCREEN_STATE.get(k) for k in
          ("running", "shop", "done", "total", "ok", "fail", "last_err",
           "last_triage", "last_report")}
    dist = {"是": 0, "否": 0, "拿不准": 0}
    total = 0
    if shop:
        for (s, _sku), rec in _prescreen_map().items():
            if s != shop:
                continue
            total += 1
            v = str((rec or {}).get("verdict") or "")
            if v in dist:
                dist[v] += 1
    st["shop_prescreened"] = total
    st["dist"] = dist
    return jsonify({"ok": True, "status": st})


@app.route("/api/learn/status")
def api_learn_status():
    try:
        from learn import learn_status, load_syn_pending
        from learn import BRAND_FILL_FILE, BRAND_COLS, _read_df as _rd
        st = learn_status()
        pending_syn = [p for p in load_syn_pending() if p.get("status") == "待启用"]
        st["synonyms"] = [{"variant": p.get("variant", ""), "standard": p.get("standard", ""),
                           "time": p.get("time", "")} for p in pending_syn]
        bdf = _rd(BRAND_FILL_FILE, BRAND_COLS)
        st["brand_fills"] = []
        if len(bdf):
            for _, r in bdf[bdf["状态"].astype(str) == "待启用"].iterrows():
                st["brand_fills"].append({"sku": _safe(r.get("竞品Sku_id")),
                                          "店": _safe(r.get("竞店名称")),
                                          "品牌": _safe(r.get("推断品牌")),
                                          "型号": _safe(r.get("推断型号")),
                                          "来源": _safe(r.get("来源"))})
        return jsonify({"ok": True, "status": st})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/learn/extract", methods=["POST"])
def api_learn_extract():
    """提炼同义词（大模型，批量一次性）。"""
    try:
        from learn import extract_synonyms
        added = extract_synonyms()
        return jsonify({"ok": True, "msg": f"提炼完成，新增 {len(added)} 个待启用同义词", "added": added})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/learn/synonym", methods=["POST"])
def api_learn_synonym():
    data = request.get_json(force=True)
    variant = (data.get("variant") or "").strip()
    standard = (data.get("standard") or "").strip()
    enable = data.get("enable", True)
    if not variant:
        return jsonify({"ok": False, "msg": "缺少词条"}), 400
    try:
        from learn import enable_synonym
        enable_synonym(variant, standard, enable)
        return jsonify({"ok": True, "msg": "已启用" if enable else "已忽略"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/learn/brand", methods=["POST"])
def api_learn_brand():
    data = request.get_json(force=True)
    sku = (data.get("sku") or "").strip()
    enable = data.get("enable", True)
    if not sku:
        return jsonify({"ok": False, "msg": "缺少竞品编号"}), 400
    try:
        from learn import enable_brand_fill
        enable_brand_fill(sku, enable)
        return jsonify({"ok": True, "msg": "已启用（重跑对齐生效）" if enable else "已忽略"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/advice/generate", methods=["POST"])
def api_advice_generate():
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    mock = data.get("mock") is True
    if not shops:
        return jsonify({"ok": False, "msg": "请至少选择一家店"}), 400
    try:
        from ai_advice import generate_advice
        rows, report_md, path = generate_advice(shops, mock=mock)
        return jsonify({"ok": True, "rows": rows, "report": report_md, "path": path,
                        "msg": f"已生成 {len(rows)} 条建议"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/advice/status", methods=["POST"])
def api_advice_status():
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    sku = (data.get("sku") or "").strip()
    status = (data.get("status") or "待看").strip()
    if status not in ("待看", "已采纳", "忽略"):
        return jsonify({"ok": False, "msg": "状态不合法"}), 400
    try:
        from ai_advice import ADVICE_FILE
        df = pd.read_excel(ADVICE_FILE, dtype=str).fillna("")
        mask = (df["竞店名称"] == shop) & (df["竞品Sku_id"] == sku)
        if mask.any():
            df.loc[mask, "状态"] = status
            df.to_excel(ADVICE_FILE, index=False)
        return jsonify({"ok": True, "msg": "已更新"})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/advice")
def advice_page():
    # 已对齐建议降级为品详情页内一块；生成入口收进工具页。保留路由兼容旧链接。
    return redirect("/tools")


@app.route("/api/pending_skip", methods=["POST"])
def api_pending_skip():
    """低价值批量跳过：把 待确认 且 竞品月售<阈值 的候选行标记为无对应（备注=低价值跳过）。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    threshold = data.get("threshold", 10)
    try:
        threshold = int(threshold)
    except (ValueError, TypeError):
        threshold = 10

    skipped = 0
    if pool_df is not None and len(pool_df):
        fine = pool_df[pool_df["匹配层"].astype(str) != "品级候选(人工)"]
        if shop:
            fine = fine[fine["竞店名称"] == shop]
        for idx in fine.index:
            shop_s = _shop_of(idx)
            sku = _sku_of(idx)
            rec = progress.get("confirmed", {}).get(shop_s, {}).get(sku)
            if rec or _is_no_match(idx):
                continue
            # 注意：Series 不能参与 `or {}`（布尔歧义），先取后判空
            row = competitor_lookup.get(shop_s, {}).get(sku)
            if row is None:
                row = {}
            sales = _extract_num(row.get("sales") or row.get("月销量") or row.get("月售"))
            if sales is not None and sales < threshold:
                progress.setdefault("no_match", {}).setdefault(shop_s, {})[sku] = {
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "note": "低价值跳过",
                }
                skipped += 1
    save_progress()
    return jsonify({"ok": True, "msg": f"已跳过 {skipped} 条低价值候选（月售<{threshold}）",
                    "skipped": skipped})


@app.route("/api/ai/category", methods=["POST"])
def api_ai_category():
    """类目 AI 分析（类层）：映射建议/缺口/价格带/竞争度。"""
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    mock = data.get("mock") is True
    if not shops:
        return jsonify({"ok": False, "msg": "请至少选择一家店"}), 400
    try:
        from ai_category import generate_category_advice
        report = generate_category_advice(shops, mock=mock)
        return jsonify({"ok": True, "report": report})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/confirm_spu_direct", methods=["POST"])
def api_confirm_spu_direct():
    """仅确认此品（粗对齐）：不依赖池行，按 shop+sku+own_sku 直接写 SPU 层。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    sku = (data.get("sku") or "").strip()
    own_sku = (data.get("own_sku") or "").strip()
    if shop not in competitor_lookup or sku not in competitor_lookup.get(shop, {}):
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    row = competitor_lookup[shop][sku]
    oi = own_lookup.get(own_sku) if own_sku else None
    if oi is None:
        return jsonify({"ok": False, "msg": "请先搜索并选择自家品"}), 400

    pid = _safe(row.get("comp_id") or row.get("商品ID"))
    spu_gid = own_spu_map.get(own_sku, "")
    own_name = _safe(oi.get("商品名称"))
    from summary_library import load_library as _ll, save_library as _sl, upsert_manual as _um
    spu_df, sku_df = _ll()
    spu_df = _um(spu_df, "竞品品ID", shop, {
        "竞品品ID": pid, "竞品商品名": _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
        "竞品品牌": _safe(row.get("品牌（抽取）") or row.get("品牌名称")),
        "竞品型号": _safe(row.get("型号")),
        "竞品类目": _safe(row.get("类目三级_归一") or row.get("category3")),
        "自家品组ID": spu_gid, "自家品名": own_name, "置信度": "",
    })
    _sl(spu_df, sku_df)
    progress.setdefault("confirmed", {}).setdefault(shop, {})[sku] = {
        "sku_id": own_sku, "spu_only": True,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "direct": True,
    }
    progress.setdefault("no_match", {}).get(shop, {}).pop(sku, None)
    save_progress()
    # 类目映射顺带确认 + 学习 + 跨店聚合
    try:
        _confirm_category_map(_safe(row.get("类目三级_归一") or row.get("category3")),
                              _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级")))
    except Exception:
        pass
    try:
        from learn import record as learn_record, add_keyword as learn_add_kw
        learn_record(shop, sku, _safe(row.get("name") or row.get("商品名字") or row.get("清洗后商品名")),
                     own_sku, own_name, source="手动确认品",
                     comp_brand=_safe(row.get("品牌（抽取）") or row.get("品牌名称")),
                     comp_model=_safe(row.get("型号")),
                     own_brand=_safe(oi.get("品牌名称")), own_model=_safe(oi.get("型号")))
        learn_add_kw(own_name, _safe(oi.get("SPUID")),
                     _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级")))
    except Exception:
        pass
    agg_n = 0
    try:
        agg_n = _aggregate_spu(shop, sku, own_sku, own_name, spu_gid)
    except Exception:
        pass
    return jsonify({"ok": True, "msg": "已确认此品（粗对齐），规格待补" +
                    (f"，并关联 {agg_n} 家店同款" if agg_n else "")})


@app.route("/api/no_match_direct", methods=["POST"])
def api_no_match_direct():
    """详情页标记无对应（及恢复）。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    sku = (data.get("sku") or "").strip()
    restore = data.get("restore") is True
    if shop not in competitor_lookup or sku not in competitor_lookup.get(shop, {}):
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    if restore:
        progress.setdefault("no_match", {}).get(shop, {}).pop(sku, None)
        save_progress()
        return jsonify({"ok": True, "msg": "已恢复为待处理"})
    progress.setdefault("no_match", {}).setdefault(shop, {})[sku] = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "note": "详情页标记无对应",
    }
    progress.setdefault("confirmed", {}).get(shop, {}).pop(sku, None)
    save_progress()
    return jsonify({"ok": True, "msg": "已标记无对应"})


def _load_library_cached():
    """汇总库两层读取（缓存；save_progress 时失效）。"""
    global _lib_cache
    if _lib_cache is None:
        from summary_library import load_library
        _lib_cache = load_library()
    return _lib_cache


@app.route("/confirm")
def confirm_page():
    # 人工确认已并入竞品库（/browse 状态筛选Tab）；保留路由兼容旧链接
    shop = request.args.get("shop", "").strip()
    status_f = (request.args.get("status") or "").strip()
    view = (request.args.get("view") or "").strip()
    if status_f == "confirmed" or view == "sku":
        status_f = "aligned"
    elif status_f not in ("all", "pending", "coarse", "nomatch"):
        status_f = "all"
    extra = ""
    if shop:
        extra = "?shop=" + quote(shop)
    if status_f != "all":
        extra += ("&" if extra else "?") + "status=" + status_f
    return redirect("/browse" + extra)


def _legacy_confirm_page():
    page = request.args.get("page", 1, type=int)
    status_f = (request.args.get("status") or "all").strip()
    if status_f not in ("all", "pending", "coarse", "confirmed", "nomatch"):
        status_f = "all"
    view = (request.args.get("view") or "pool").strip()
    if view not in ("pool", "spu", "sku"):
        view = "pool"
    sort_f = (request.args.get("sort") or "sales").strip()
    if sort_f not in ("sales", "conf"):
        sort_f = "sales"
    if pool_df is None or not len(pool_df):
        return render_template("index.html", items=[], page=1, total_pages=1, total=0,
                               confirmed=0, coarse=0, no_match=0, pending=0, percent=0,
                               shops=[], current_shop="", status=status_f, view=view,
                               spu_rows=[], sku_rows=[])

    # 店铺统计（两层库口径：与下方"已对齐SPU/SKU"Tab 一致）
    per = _get_stats_all()
    shops = [{"name": s, "total": v["total"], "pending": v["pending"],
              "done": v["total"] - v["pending"], **v} for s, v in per.items()]

    shop = request.args.get("shop", "").strip()
    # 店铺范围：shop 为空 = 全部店铺（含默认店）
    # 细对齐只含第3/4层候选；品级候选行归粗对齐栏目（/coarse），排除以免重复
    pool_fine = pool_df[pool_df["匹配层"].astype(str) != "品级候选(人工)"]
    if shop:
        sub = pool_fine[pool_fine["竞店名称"] == shop]
        sub_idx = sub.index.tolist()
    else:
        sub = pool_fine
        sub_idx = list(range(len(pool_fine)))

    # 按状态筛选（coarse=已粗对齐：品确认过、规格待补；confirmed=已对应：品+规格）
    def _status_of(idx):
        global _pool_status_cache
        if _pool_status_cache is None:
            conf = progress.get("confirmed", {})
            nm = progress.get("no_match", {})
            cache = []
            for i in range(len(_pool_sku_list)):
                shop_s = _pool_shop_list[i]
                sku = _pool_sku_list[i]
                rec = conf.get(shop_s, {}).get(sku)
                if rec:
                    cache.append("coarse" if rec.get("spu_only") else "confirmed")
                elif nm.get(shop_s, {}).get(sku):
                    cache.append("nomatch")
                else:
                    cache.append("pending")
            _pool_status_cache = cache
        return _pool_status_cache[idx]

    if status_f != "all":
        sub_idx = [i for i in sub_idx if _status_of(i) == status_f]

    # 排序：sort=sales 按竞品月售降序（默认，价值优先）；sort=conf 按置信度降序（先做高置信，强化学习）
    if sort_f == "conf":
        global _pool_conf_cache
        if _pool_conf_cache is None:
            if "置信度" in pool_df.columns:
                _pool_conf_cache = (pd.to_numeric(pool_df["置信度"], errors="coerce")
                                    .fillna(0.0).tolist())
            else:
                _pool_conf_cache = [0.0] * len(pool_df)
        sub_idx.sort(key=lambda i: _pool_conf_cache[i], reverse=True)
    else:
        # 价值排序：按竞品月售数字降序（主），价格降序（次）——高价值品优先处理
        def _row_value(idx):
            global _pool_value_cache
            if _pool_value_cache is None:
                vals = []
                for i in range(len(_pool_sku_list)):
                    row = competitor_lookup.get(_pool_shop_list[i], {}).get(_pool_sku_list[i])
                    if row is None:
                        vals.append((-1, -1))
                        continue
                    sales = _extract_num(row.get("sales") or row.get("月销量") or row.get("月售"))
                    price = _extract_num(row.get("price") or row.get("到手价") or row.get("折后价"))
                    vals.append((sales if sales is not None else -1, price if price is not None else -1))
                _pool_value_cache = vals
            return _pool_value_cache[idx]

        sub_idx.sort(key=_row_value, reverse=True)

    total_all = len(sub)          # 统计卡总数（未筛选，当前范围全量）
    total = len(sub_idx)          # 列表分页条数（筛选后）
    total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
    page = max(1, min(page, total_pages))
    start = (page - 1) * PAGE_SIZE
    end = start + PAGE_SIZE

    # 两层已对应数据（SPU 层已粗对齐 / SKU 层已对应；各取前100条展示）
    spu_rows, sku_rows = [], []
    try:
        spu_df_lib, sku_df_lib = _load_library_cached()
        m_spu = spu_df_lib[spu_df_lib["状态"].astype(str) == "有效"] if len(spu_df_lib) else spu_df_lib
        m_sku = sku_df_lib[sku_df_lib["状态"].astype(str) == "有效"] if len(sku_df_lib) else sku_df_lib
        if shop:
            m_spu = m_spu[m_spu["竞店名称"] == shop]
            m_sku = m_sku[m_sku["竞店名称"] == shop]
        for _, r in m_spu.head(100).iterrows():
            spu_rows.append({"竞品商品名": _safe(r.get("竞品商品名")),
                             "自家品名": _safe(r.get("自家品名")),
                             "确认来源": _safe(r.get("确认来源")),
                             "竞店名称": _safe(r.get("竞店名称"))})
        for _, r in m_sku.head(100).iterrows():
            sku_rows.append({"竞品商品名": _safe(r.get("竞品商品名")),
                             "自家商品名": _safe(r.get("自家商品名")),
                             "匹配层": _safe(r.get("匹配层")),
                             "竞店名称": _safe(r.get("竞店名称")),
                             "竞品Sku_id": _safe(r.get("竞品Sku_id"))})
    except Exception:
        pass

    items = []
    for idx in sub_idx[start:end]:
        row = pool_df.iloc[idx]
        st = _status_of(idx)
        if st == "confirmed":
            status_text, status_class = "已对应", "status-confirmed"
        elif st == "coarse":
            status_text, status_class = "已粗对齐", "status-coarse"
        elif st == "nomatch":
            status_text, status_class = "无对应", "status-nomatch"
        else:
            status_text, status_class = "待确认", "status-pending"
        items.append({
            "idx": idx,
            "竞品商品名": _safe(row.get("竞品商品名")),
            "竞品品牌": _safe(row.get("竞品品牌")),
            "竞品类目": _safe(row.get("竞品类目")),
            "自家商品名": _safe(row.get("自家商品名")),
            "匹配层": _safe(row.get("匹配层")),
            "置信度": _safe(row.get("置信度")),
            "status_text": status_text,
            "status_class": status_class,
        })

    # 统计口径 = 与下方"已对齐SPU/SKU"Tab 完全一致：已粗对齐=SPU层有效行数（当前店），
    # 已对应=SKU层有效行数（当前店），无对应=no_match 数，总数=竞品库，待处理=其余
    total_all = _get_stats(shop)["total"]
    coarse_count = len(m_spu)
    confirmed_count = len(m_sku)
    nm_count = sum(len(v) for s2, v in progress.get("no_match", {}).items() if not shop or s2 == shop)
    no_match_count = nm_count
    pending_count = max(total_all - coarse_count - confirmed_count - no_match_count, 0)
    percent = round((coarse_count + confirmed_count + no_match_count) / max(total_all, 1) * 100, 1)

    return render_template("index.html",
                           items=items, page=page, total_pages=total_pages,
                           total=total_all, confirmed=confirmed_count,
                           coarse=coarse_count, no_match=no_match_count,
                           pending=pending_count,
                           percent=percent, shops=shops, current_shop=shop,
                           status=status_f, view=view, sort=sort_f,
                           spu_rows=spu_rows, sku_rows=sku_rows,
                           spu_total=len(m_spu), sku_total=len(m_sku))


@app.route("/item/<int:idx>")
def item_detail(idx):
    if pool_df is None or idx < 0 or idx >= len(pool_df):
        abort(404)
    item = get_pool_item(idx)
    shop = request.args.get("shop", "").strip() or item["竞店名称"]

    # 当前店内翻页（保持店铺上下文）
    shop_rows = pool_df.index[pool_df["竞店名称"] == shop].tolist()
    pos = shop_rows.index(idx) if idx in shop_rows else 0
    next_idx = shop_rows[pos + 1] if pos + 1 < len(shop_rows) else None
    prev_idx = shop_rows[pos - 1] if pos > 0 else None

    # 候选自家的完整信息
    candidate_own = get_own_info(item["自家SKUID"]) if item["自家SKUID"] else {}

    # 已确认状态（按店+竞品Sku_id）
    is_confirmed = _is_confirmed(idx)
    is_no_match = _is_no_match(idx)

    return render_template("item.html",
                           item=item, idx=idx, total=len(pool_df),
                           next_idx=next_idx, prev_idx=prev_idx,
                           candidate_own=candidate_own,
                           own_categories=own_categories,
                           own_brands=own_brands,
                           is_confirmed=is_confirmed,
                           is_no_match=is_no_match,
                           current_shop=shop)


@app.route("/api/search")
def api_search():
    q = request.args.get("q", "").strip()
    category = request.args.get("category", "").strip()
    brand = request.args.get("brand", "").strip()

    df = own_df
    mask = pd.Series([True] * len(df), index=df.index)

    if q:
        # 关键词搜索：空格/分隔符分词，逐词 AND 包含（如「自攻螺丝 M4 18-35mm」）
        from search_util import kw_mask
        kwm = kw_mask(df, q, ["商品名称", "清洗后商品名", "原始商品名", "SKUID", "SPUID", "商品条码"])
        if kwm is not None:
            mask = mask & kwm

    if category:
        cat_col = "类目三级_归一" if "类目三级_归一" in df.columns else "美团闪购类目三级"
        if cat_col in df.columns:
            mask = mask & (df[cat_col].astype(str) == category)

    if brand:
        brand_col = "品牌（抽取）" if "品牌（抽取）" in df.columns else "品牌名称"
        if brand_col in df.columns:
            mask = mask & (df[brand_col].astype(str) == brand)

    result_df = df[mask].head(SEARCH_LIMIT)
    results = []
    for _, row in result_df.iterrows():
        results.append({
            "SKUID": _safe(row.get("SKUID")),
            "商品名称": _safe(row.get("商品名称")),
            "规格名称": _safe(row.get("规格名称")),
            "品牌名称": _safe(row.get("品牌名称")),
            "品牌抽取": _safe(row.get("品牌（抽取）")),
            "类目三级": _safe(row.get("美团闪购类目三级")),
            "类目归一": _safe(row.get("类目三级_归一")),
            "售价": _safe(row.get("美团闪购渠道售价")),
            "月销量": _safe(row.get("月销量")),
            "是否有图片": _safe(row.get("是否有图片")),
            "型号": _safe(row.get("型号")),
        })
    return jsonify({"results": results, "count": len(results)})


@app.route("/api/confirm", methods=["POST"])
def api_confirm():
    data = request.get_json(force=True)
    idx = data.get("idx")
    selected_sku = _safe(data.get("sku_id"))
    note = data.get("note", "")

    if idx is None:
        return jsonify({"ok": False, "msg": "缺少候选ID"}), 400
    idx = int(idx)
    if pool_df is None or idx < 0 or idx >= len(pool_df):
        return jsonify({"ok": False, "msg": "候选ID越界"}), 400
    if not selected_sku:
        return jsonify({"ok": False, "msg": "请先选定一个自家品"}), 400

    pool_row = pool_df.iloc[idx]
    shop = _shop_of(idx)
    own_info = get_own_info(selected_sku)
    if not own_info:
        return jsonify({"ok": False, "msg": "自家SKUID不存在: " + selected_sku}), 400

    # ---- 写两层汇总库（人工确认行，机器重跑不覆盖数据列）----
    sku_row = {
        "竞品品ID": _safe(pool_row.get("竞品品ID")),
        "竞品Sku_id": _safe(pool_row.get("竞品Sku_id")),
        "竞品商品名": _safe(pool_row.get("竞品商品名")),
        "竞品条形码": _safe(pool_row.get("竞品条形码")),
        "自家SPUID": own_info.get("SPUID", ""),
        "自家SKUID": own_info.get("SKUID", ""),
        "自家商品名": own_info.get("商品名称", ""),
        "匹配层": _safe(pool_row.get("匹配层")),
        "匹配依据": _safe(pool_row.get("匹配依据")),
        "置信度": _safe(pool_row.get("置信度")),
    }
    spu_gid = own_spu_map.get(own_info.get("SKUID", ""), "")
    spu_row = {
        "竞品品ID": _safe(pool_row.get("竞品品ID")),
        "竞品商品名": _safe(pool_row.get("竞品商品名")),
        "竞品品牌": _safe(pool_row.get("竞品品牌")),
        "竞品型号": _safe(pool_row.get("竞品型号")),
        "竞品类目": _safe(pool_row.get("竞品类目")),
        "自家品组ID": spu_gid or own_info.get("SPUID", ""),
        "自家品名": own_info.get("商品名称", ""),
        "置信度": _safe(pool_row.get("品级置信度")) or _safe(pool_row.get("置信度")),
    }
    spu_df, sku_df = load_library()
    spu_df = upsert_manual(spu_df, "竞品品ID", shop, spu_row)
    sku_df = upsert_manual(sku_df, "竞品Sku_id", shop, sku_row)
    save_library(spu_df, sku_df)

    # 学习闭环：记录确认样本（样本库 + 品牌补全）
    try:
        from learn import record as learn_record, add_keyword as learn_add_kw
        learn_record(shop, _safe(pool_row.get("竞品Sku_id")), _safe(pool_row.get("竞品商品名")),
                     own_info.get("SKUID", ""), own_info.get("商品名称", ""), source="人工确认",
                     comp_brand=_safe(pool_row.get("竞品品牌")), comp_model=_safe(pool_row.get("竞品型号")),
                     own_brand=own_info.get("品牌名称", ""), own_model=own_info.get("型号", ""))
        learn_add_kw(own_info.get("商品名称", ""), own_info.get("SPUID", ""),
                     own_info.get("类目三级_归一", "") or own_info.get("美团闪购类目三级", ""))
    except Exception:
        pass

    # 更新 pool_df 确认结果
    pool_df.at[idx, "确认结果"] = "确认"
    remark = "确认对齐 -> SKUID:" + selected_sku
    if note:
        remark += " | " + note
    pool_df.at[idx, "人工备注"] = remark
    save_pool_excel()

    # 更新进度（按店+竞品Sku_id）
    _mark_confirmed(idx, selected_sku)
    save_progress()

    return jsonify({"ok": True, "msg": "确认成功，已写入两层汇总库（SKU匹配层+SPU匹配层）"})


@app.route("/api/confirm_spu", methods=["POST"])
def api_confirm_spu():
    """仅确认品级（规格待补）：写 SPU匹配层（人工确认），不写 SKU层。"""
    data = request.get_json(force=True)
    idx = data.get("idx")
    spu_gid = _safe(data.get("spu_gid"))
    spu_name = _safe(data.get("spu_name"))
    note = data.get("note", "")

    if idx is None:
        return jsonify({"ok": False, "msg": "缺少候选ID"}), 400
    idx = int(idx)
    if pool_df is None or idx < 0 or idx >= len(pool_df):
        return jsonify({"ok": False, "msg": "候选ID越界"}), 400
    if not spu_gid:
        return jsonify({"ok": False, "msg": "缺少自家品组ID"}), 400

    pool_row = pool_df.iloc[idx]
    shop = _shop_of(idx)
    spu_row = {
        "竞品品ID": _safe(pool_row.get("竞品品ID")),
        "竞品商品名": _safe(pool_row.get("竞品商品名")),
        "竞品品牌": _safe(pool_row.get("竞品品牌")),
        "竞品型号": _safe(pool_row.get("竞品型号")),
        "竞品类目": _safe(pool_row.get("竞品类目")),
        "自家品组ID": spu_gid,
        "自家品名": spu_name,
        "置信度": _safe(pool_row.get("品级置信度")) or _safe(pool_row.get("置信度")),
    }
    spu_df, sku_df = load_library()
    spu_df = upsert_manual(spu_df, "竞品品ID", shop, spu_row)
    save_library(spu_df, sku_df)

    # 类目映射顺带确认：竞品类目 -> 候选品组所属自家类目，写回映射表
    try:
        own_cat = ""
        raw = _safe(pool_row.get("候选自家品"))
        if raw:
            for c in json.loads(raw):
                if c.get("品组ID") == spu_gid and c.get("SKU列表"):
                    oi = own_lookup.get(_safe((c.get("SKU列表") or [""])[0]), {})
                    own_cat = _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级"))
                    break
        if own_cat:
            _confirm_category_map(_safe(pool_row.get("竞品类目")), own_cat)
    except Exception:
        pass

    # 学习闭环：品级确认样本（无 SKU 级信息，仅记样本）
    try:
        from learn import record as learn_record
        learn_record(shop, _safe(pool_row.get("竞品Sku_id")), _safe(pool_row.get("竞品商品名")),
                     "", spu_name, source="仅确认品级",
                     comp_brand=_safe(pool_row.get("竞品品牌")), comp_model=_safe(pool_row.get("竞品型号")))
    except Exception:
        pass

    # 更新 pool_df 确认结果（品级确认，规格待补）
    pool_df.at[idx, "确认结果"] = "确认"
    remark = "仅确认品级 -> 品组:" + spu_gid
    if note:
        remark += " | " + note
    pool_df.at[idx, "人工备注"] = remark
    save_pool_excel()

    # 更新进度（按店+竞品Sku_id；记录 spu_only 标记）
    progress.setdefault("confirmed", {}).setdefault(shop, {})[_sku_of(idx)] = {
        "sku_id": spu_gid,
        "spu_only": True,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }
    progress.setdefault("no_match", {}).get(shop, {}).pop(_sku_of(idx), None)
    save_progress()

    # 跨店聚合：其他店同款竞品一并写入 SPU 层（标注聚合关联）
    agg_n = 0
    try:
        agg_n = _aggregate_spu(shop, _safe(pool_row.get("竞品Sku_id")), "", spu_name, spu_gid)
    except Exception:
        pass

    return jsonify({"ok": True, "msg": "品级确认成功，已写入汇总库 SPU匹配层（规格待补）" +
                    (f"，并关联 {agg_n} 家店同款" if agg_n else "")})


@app.route("/api/no_match", methods=["POST"])
def api_no_match():
    data = request.get_json(force=True)
    idx = data.get("idx")
    note = data.get("note", "")

    if idx is None:
        return jsonify({"ok": False, "msg": "缺少候选ID"}), 400
    idx = int(idx)
    if pool_df is None or idx < 0 or idx >= len(pool_df):
        return jsonify({"ok": False, "msg": "候选ID越界"}), 400

    pool_row = pool_df.iloc[idx]
    shop = _shop_of(idx)
    comp_info = get_competitor_info(pool_row.get("竞品Sku_id"), shop)

    new_row = {
        "竞店名称": shop,
        "竞品品ID": _safe(pool_row.get("竞品品ID")),
        "竞品Sku_id": _safe(pool_row.get("竞品Sku_id")),
        "竞品商品名": _safe(pool_row.get("竞品商品名")),
        "竞品品牌": _safe(pool_row.get("竞品品牌")),
        "竞品型号": _safe(pool_row.get("竞品型号")),
        "竞品类目": _safe(pool_row.get("竞品类目")),
        "竞品规格键": _safe(pool_row.get("竞品规格键")),
        "竞品价格": comp_info.get("价格", ""),
        "竞品月售": comp_info.get("月售", ""),
        "竞品图片url": comp_info.get("图片url", ""),
        "原候选自家SKUID": _safe(pool_row.get("自家SKUID")),
        "标记时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "备注": note,
    }
    append_to_excel(NO_MATCH_FILE, new_row)

    # 更新 pool_df
    pool_df.at[idx, "确认结果"] = "无对应"
    remark = "标记无对应"
    if note:
        remark += " | " + note
    pool_df.at[idx, "人工备注"] = remark
    save_pool_excel()

    # 更新进度（按店+竞品Sku_id）
    _mark_no_match(idx)
    save_progress()

    return jsonify({"ok": True, "msg": "已标记无对应，写入无对应品清单.xlsx"})


@app.route("/api/progress")
def api_progress():
    shop = request.args.get("shop", "").strip()
    if pool_df is None or not len(pool_df):
        return jsonify({"total": 0, "confirmed": 0, "no_match": 0, "pending": 0, "percent": 0})
    if shop:
        sub = pool_df[pool_df["竞店名称"] == shop]
        conf_set = set(progress.get("confirmed", {}).get(shop, {}).keys())
        nm_set = set(progress.get("no_match", {}).get(shop, {}).keys())
    else:
        sub = pool_df
        conf_set = set()
        nm_set = set()
        for s, m in progress.get("confirmed", {}).items():
            conf_set |= set(m.keys())
        for s, m in progress.get("no_match", {}).items():
            nm_set |= set(m.keys())
    total = len(sub)
    confirmed = sum(1 for sku in sub["竞品Sku_id"] if sku in conf_set)
    no_match = sum(1 for sku in sub["竞品Sku_id"] if sku in nm_set)
    pending = total - confirmed - no_match
    percent = round((confirmed + no_match) / max(total, 1) * 100, 1)
    return jsonify({
        "total": total,
        "confirmed": confirmed,
        "no_match": no_match,
        "pending": pending,
        "percent": percent,
    })


# ---------------------------------------------------------------------------
# 选店分析：勾选参与对比的竞店 -> 运行分析 -> 报告按所选店计算
# ---------------------------------------------------------------------------
ANALYSIS_CONFIG = os.path.join(BASE_DIR, "config", "analysis_shops.yaml")
SHOP_REPORT_FILE = os.path.join(OUTPUT_DIR, "竞店标准库", "竞店数据质量报告.xlsx")
COMPETE_REPORT_FILE = os.path.join(OUTPUT_DIR, "竞争分析报告.xlsx")


def load_analysis_shops():
    """返回 None=未配置(默认全选)；set=已保存的所选店。"""
    if os.path.exists(ANALYSIS_CONFIG):
        try:
            import yaml
            with open(ANALYSIS_CONFIG, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            lst = data.get("shops") or []
            if lst:
                return set(lst)
        except Exception:
            pass
    return None


@app.route("/compare")
def compare_page():
    # 产品分析/top50/对比明细已并入首页榜单；保留路由兼容旧链接
    shop = (request.args.get("shop") or "").strip()
    filter_f = (request.args.get("filter") or "").strip()
    extra = ""
    if shop:
        extra = "?shop=" + quote(shop)
    if filter_f == "opportunity":
        extra += ("&" if extra else "?") + "filter=opp"
    return redirect("/" + extra)


def _legacy_compare_page():
    """同品对比 + 产品分析 + 销量前50人工匹配：view=detail/product/top50。"""
    shop = (request.args.get("shop") or "").strip()
    if not shop and competitor_lookup:
        shop = next(iter(competitor_lookup))
    view = (request.args.get("view") or "product").strip()
    if view not in ("product", "detail", "top50"):
        view = "product"
    filter_f = (request.args.get("filter") or "all").strip()
    if filter_f not in ("all", "cross", "opportunity"):
        filter_f = "all"

    if view == "top50":
        # ---- 竞品销量前50品（人工匹配工作台）----
        try:
            from top50 import build_top50
            key = (shop, "top50")
            if key not in _compare_cache:
                _compare_cache[key] = _derived("top50", shop, lambda: build_top50(shop),
                                               code=_fp_derived())
            prods = _compare_cache[key]
        except Exception as e:
            print("[销量前50] 计算失败:", e)
            prods = []
        # 匹配状态：任一SKU已确认(spu_only)=已匹配；任一SKU无对应=无对应
        from summary_library import load_library as _ll
        try:
            spu_lib, _ = _ll()
            spu_own = dict(zip(
                spu_lib[spu_lib["竞店名称"] == shop]["竞品品ID"].astype(str).str.strip(),
                spu_lib[spu_lib["竞店名称"] == shop]["自家品名"]))
        except Exception:
            spu_own = {}
        matched = 0
        for p in prods:
            pid = p["comp_id"]
            skus = [s["sku_id"] for s in p.get("SKU列表", [])]
            conf = [s for s in skus if progress.get("confirmed", {}).get(shop, {}).get(s)]
            nm = [s for s in skus if progress.get("no_match", {}).get(shop, {}).get(s)]
            if conf:
                p["状态"] = "已匹配"
                p["自家品名"] = spu_own.get(pid, "")
                matched += 1
            elif nm:
                p["状态"] = "无对应"
            else:
                p["状态"] = "待匹配"
        _adv50 = advice_state(shop)
        return render_template("top50.html", products=prods, matched=matched,
                               advice=_adv50["rows"], advice_stale=_adv50["stale"],
                               advice_at=_adv50["at"],
                               current_shop=shop, view=view,
                               shops=[{"name": s} for s in competitor_lookup.keys()])

    if view == "product":
        # ---- 产品分析（品为单位，规则版；结果缓存，数据变更时失效）----
        key = (shop, "product")
        if key not in _compare_cache:
            try:
                # 复用对比明细缓存，避免重复计算
                key_d = (shop, "detail")
                if key_d not in _compare_cache:
                    from build_comparison import build_comparison_data
                    _compare_cache[key_d] = _derived(
                        "detail", shop, lambda: build_comparison_data(shop), code=_fp_derived())
                detail_d, _ = _compare_cache[key_d]
                from build_product_analysis import build_product_analysis
                _compare_cache[key] = _derived(
                    "product", shop,
                    lambda: build_product_analysis(shop, detail=detail_d), code=_fp_derived())
            except Exception as e:
                print("[产品分析] 计算失败:", e)
                _compare_cache[key] = (pd.DataFrame(), pd.DataFrame())
        pdf, cdf = _compare_cache[key]
        if len(pdf):
            if filter_f == "cross":
                pdf = pdf[pdf["在卖品牌数"] >= 1]
            elif filter_f == "opportunity":
                pdf = pdf[pdf["结论"].isin(["机会品", "防守品"])]
        stats = {
            "total": len(pdf),
            "opp": int((pdf["结论"] == "机会品").sum()) if len(pdf) else 0,
            "def": int((pdf["结论"] == "防守品").sum()) if len(pdf) else 0,
        }
        items = pdf.head(200).to_dict("records") if len(pdf) else []
        agg_items = cdf.head(50).to_dict("records") if len(cdf) else []
        return render_template("compare.html", items=items, agg_items=agg_items,
                               stats=stats, current_shop=shop, filter_f=filter_f,
                               view=view, is_product=True,
                               shops=[{"name": s} for s in competitor_lookup.keys()])

    # ---- 对比明细（SKU 级；结果缓存，数据变更时失效）----
    key = (shop, "detail")
    if key not in _compare_cache:
        try:
            from build_comparison import build_comparison_data
            _compare_cache[key] = _derived("detail", shop,
                                           lambda: build_comparison_data(shop), code=_fp_derived())
        except Exception as e:
            print("[同品对比] 计算失败:", e)
            _compare_cache[key] = (pd.DataFrame(), pd.DataFrame())
    detail, agg = _compare_cache[key]
    if len(detail):
        if filter_f == "cross":
            detail = detail[detail["跨品牌"] == "是"]
        elif filter_f == "opportunity":
            mask = detail["结论"].astype(str)
            detail = detail[mask.str.contains("重点|机会|更高", na=False)]
    stats = {
        "total": len(detail),
        "cross": int((detail["跨品牌"] == "是").sum()) if len(detail) else 0,
        "opp": int(detail["结论"].astype(str).str.contains("重点|机会|更高").sum()) if len(detail) else 0,
    }
    items = detail.head(200).to_dict("records") if len(detail) else []
    agg_items = agg.head(50).to_dict("records") if len(agg) else []
    return render_template("compare.html", items=items, agg_items=agg_items,
                           stats=stats, current_shop=shop, filter_f=filter_f,
                           view=view, is_product=False,
                           shops=[{"name": s} for s in competitor_lookup.keys()])


_advice_lib_cache = {"mtime": None, "rows": []}


def _load_advice_rows():
    """读 output/已对齐建议.xlsx（mtime 缓存）。返回 list[dict]；文件不存在返回空。"""
    f = os.path.join(OUTPUT_DIR, "已对齐建议.xlsx")
    if not os.path.exists(f):
        return []
    try:
        mt = os.path.getmtime(f)
        if _advice_lib_cache["mtime"] != mt:
            df = pd.read_excel(f, dtype=str).fillna("")
            _advice_lib_cache["mtime"] = mt
            _advice_lib_cache["rows"] = df.to_dict("records") if len(df) else []
        return _advice_lib_cache["rows"]
    except Exception:
        return []


@app.route("/product_detail")
def product_detail_page():
    """产品分析详情页：某品的自家SKU vs 对应竞品（对比卡片可视化）+ AI对齐建议。"""
    shop = (request.args.get("shop") or "").strip()
    spu = (request.args.get("spu") or "").strip()
    if not shop and competitor_lookup:
        shop = next(iter(competitor_lookup))
    if not spu:
        return redirect("/")
    # 复用对比明细缓存
    key_d = (shop, "detail")
    if key_d not in _compare_cache:
        try:
            from build_comparison import build_comparison_data
            _compare_cache[key_d] = _derived("detail", shop,
                                             lambda: build_comparison_data(shop), code=_fp_derived())
        except Exception as e:
            print("[产品详情] 对比数据失败:", e)
            _compare_cache[key_d] = (pd.DataFrame(), pd.DataFrame())
    detail_d, _ = _compare_cache[key_d]
    try:
        from build_product_analysis import build_product_detail
        data = build_product_detail(shop, spu, detail=detail_d, own_df=own_df)
    except Exception as e:
        print("[产品详情] 失败:", e)
        data = None
    if data is None:
        return redirect("/?shop=" + quote(request.args.get("shop", "")))
    # AI 对齐建议（已对齐建议.xlsx 中匹配自家品名的行）
    _pname = _safe(data.get("品名"))
    advice_rows = [r for r in _load_advice_rows() if _safe(r.get("自家商品名")) == _pname]
    # 各自家门店上架状态（多自家门店比对：该品在哪些店在售）
    store_status = []
    try:
        import own_store_lib
        for st in own_store_lib.list_stores():
            os_ = own_store_lib.get_on_sale(st["name"])
            store_status.append({
                "name": st["name"],
                "on_sale": bool(os_ and spu in os_["spu_set"]),
                "on_sale_spus": st["on_sale_spus"],
            })
    except Exception:
        pass
    return render_template("product_detail.html", data=data, current_shop=shop,
                           advice_rows=advice_rows[:20],
                           store_status=store_status)


@app.route("/api/own_spu_search")
def api_own_spu_search():
    """自家品（SPU级）搜索：销量前50人工匹配右侧用（纯人工搜索，不做机器匹配）。
    支持关键词搜索：空格/分隔符分词，逐词 AND 包含（如「自攻螺丝 M4 18-35mm」）。
    关键词 0 命中时回落**本地小模型语义兜底**（M5·2.5，向量 top10，模型未就绪静默跳过）。"""
    q = (request.args.get("q") or "").strip()
    if not q or own_df is None or not len(own_df):
        return jsonify({"results": []})
    from search_util import kw_mask
    mask = kw_mask(own_df, q, ["商品名称", "清洗后商品名", "原始商品名", "品牌（抽取）"])
    if mask is None:
        return jsonify({"results": []})
    sub = own_df[mask]
    fallback = False
    gid_order = []
    if not len(sub):
        sub, gid_order = _own_vec_sub(q, 10)
        fallback = sub is not None and len(sub) > 0
    if sub is None or not len(sub):
        return jsonify({"results": []})
    results = []
    for spuid, g in sub.groupby("SPUID"):
        g2 = g.copy()
        # P4.3 修正：总库渠道重复行，销量按 SKUID 去重后再聚合（防 ~5 倍虚高）
        g2 = g2.drop_duplicates(subset=["SKUID"], keep="first")
        g2["__s"] = g2["月销量"].apply(_extract_num).fillna(0.0)
        best = g2.loc[g2["__s"].idxmax()]
        prices = [_extract_num(v) for v in g2["美团闪购渠道售价"] if _extract_num(v) is not None]
        results.append({
            "品组ID": _safe(spuid),
            "品名": _safe(best.get("商品名称")) or _safe(best.get("清洗后商品名")),
            "品牌": _safe(best.get("品牌（抽取）")) or _safe(best.get("品牌名称")),
            "类目": _safe(best.get("类目三级_归一")) or _safe(best.get("美团闪购类目三级")),
            "月销合计": round(float(g2["__s"].sum()), 0),
            "规格数": len(g2),
            "售价区间": (round(min(prices), 2), round(max(prices), 2)) if prices else (None, None),
            "图片": _safe(best.get("商品图片url")) or _safe(best.get("图片url")),
            "兜底": "语义" if fallback else "",
        })
    if fallback:
        rank = {g: i for i, g in enumerate(gid_order)}
        results.sort(key=lambda x: rank.get(_safe(x["品组ID"]), 999))
    else:
        results.sort(key=lambda x: -x["月销合计"])
    return jsonify({"results": results[:20], "fallback": "语义" if fallback else ""})


def _own_vec_sub(q, k=10):
    """语义兜底取数（M5·2.5）：q 编码 → 与自家品组向量内积 top-k →
    返回 (这些品组的行子集, 按相似度降序的品组ID列表)。
    与 align._own_group_vectors 共用同一份缓存（output/_向量缓存/own_spu_*.npz，
    键 = 自家表指纹|模型身份|清洗版本 —— 与 align 同键同文件，谁先算谁落盘）。
    模型未就绪 / 任何异常 → 返回 (空表, [])（静默，绝不阻塞搜索主路径）。"""
    import os as _os
    import hashlib as _hl
    import json as _j
    import time as _t
    import numpy as _np
    empty = (own_df.iloc[0:0], []) if own_df is not None else (None, [])
    try:
        import vec_encoder as VE
        import text_norm as TNv
        if own_df is None or not len(own_df) or not VE.is_ready():
            return empty
        if "SPUID" not in own_df.columns:
            return empty
        try:
            fp = f"{_os.path.getmtime(OWN_FILE)}-{_os.path.getsize(OWN_FILE)}"
        except OSError:
            return empty
        tmp = own_df.copy()
        tmp["__s"] = tmp["月销量"].apply(_extract_num).fillna(0.0)
        tmp = tmp.drop_duplicates(subset=["SKUID"], keep="first")
        groups = []
        for gid, g in tmp.groupby(tmp["SPUID"].astype(str).str.strip()):
            g = g.sort_values("__s", ascending=False)
            best = g.iloc[0]
            name = _safe(best.get("商品名称")) or _safe(best.get("清洗后商品名"))
            if name:
                groups.append((_safe(gid), name, _safe(best.get("品牌（抽取）"))))
        if not groups:
            return empty
        texts = [TNv.clean(t, b or "") for _gid, t, b in groups]
        text_fp = _hl.md5("\u0001".join(texts).encode()).hexdigest()[:10]
        cache_dir = _os.path.join(OUTPUT_DIR, "_向量缓存")
        tag = "own_spu_" + _hl.md5(
            f"{fp}|{VE.model_id()}|{TNv.VERSION}".encode()).hexdigest()[:10]
        npz_path = _os.path.join(cache_dir, tag + ".npz")
        meta_path = _os.path.join(cache_dir, tag + ".json")
        gvecs = None
        if _os.path.exists(npz_path) and _os.path.exists(meta_path):
            try:
                meta = _j.load(open(meta_path, encoding="utf-8"))
                if meta.get("n") == len(groups) and meta.get("text_fp") == text_fp:
                    gvecs = _np.load(npz_path)["v"]
            except Exception:
                gvecs = None
        if gvecs is None:
            gvecs = VE.encode(texts)
            try:
                _np.savez_compressed(npz_path, v=gvecs)
                with open(meta_path, "w", encoding="utf-8") as _fm:
                    _j.dump({"n": len(groups), "text_fp": text_fp,
                             "model": VE.model_id(), "clean_version": TNv.VERSION,
                             "at": _t.strftime("%Y-%m-%d %H:%M")},
                            _fm, ensure_ascii=False)
            except Exception:
                pass
        qvec = VE.encode([q])[0]
        sims = gvecs @ qvec
        order = _np.argsort(sims)[::-1][:k]
        gids, sub = [], None
        for j in order:
            if sims[j] < 0.55:
                continue
            gid = groups[j][0]
            gids.append(gid)
        if not gids:
            return empty
        sub = own_df[own_df["SPUID"].astype(str).str.strip().isin(gids)]
        return sub, gids
    except Exception:
        return empty


@app.route("/api/own_spu_detail")
def api_own_spu_detail():
    """自家品组详情（配图）：销量前50匹配右侧点选后展示，含复制SPU。
    q 可选=当前搜索词：记录搜索点击样本（词→品组 人工标注，供关键词库反哺）。
    scene 可选=调用场景（top50/工作台/详情页），写入点击日志便于区分。"""
    spu = (request.args.get("spu") or "").strip()
    q = (request.args.get("q") or "").strip()
    scene = (request.args.get("scene") or "top50").strip() or "top50"
    if not spu or own_df is None or not len(own_df):
        return jsonify({"ok": False, "msg": "参数缺失"})
    sub = own_df[own_df["SPUID"].astype(str).str.strip() == spu]
    if not len(sub):
        return jsonify({"ok": False, "msg": "自家品不存在"})
    g = sub.copy()
    # P4.3 修正：总库同一 SKUID 有 ~5 个渠道行，销量按行求和会虚高 → 按 SKUID 去重
    g = g.drop_duplicates(subset=["SKUID"], keep="first")
    g["__s"] = g["月销量"].apply(_extract_num).fillna(0.0)
    best = g.loc[g["__s"].idxmax()]
    prices = [_extract_num(v) for v in g["美团闪购渠道售价"] if _extract_num(v) is not None]
    costs = [_extract_num(v) for v in g["采购价"] if _extract_num(v) is not None]
    skus = []
    for _, r in g.sort_values("__s", ascending=False).iterrows():
        skus.append({
            "SKUID": _safe(r.get("SKUID")),
            "规格": _safe(r.get("规格名称")) or "（默认）",
            "售价": _safe(r.get("美团闪购渠道售价")),
            "采购价": _safe(r.get("采购价")),
            "月销": round(float(_extract_num(r.get("月销量")) or 0), 0),
            "图片": _safe(r.get("商品图片url")) or _safe(r.get("图片url")),
        })
    # 规格词识别（品内区分度）：q 拆词中，匹配品内"部分"SKU 的 = 规格筛选词（品内选规格用）；
    # 匹配全部 SKU = 品词；匹配 0 = 未命中（可能同义词）。
    spec_words = []
    if q and len(skus) > 1:
        from search_util import split_keywords
        for w in split_keywords(q):
            hit = sum(1 for s in skus if w.lower() in (s["规格"] or "").lower())
            if 0 < hit < len(skus):
                spec_words.append(w)
    # 记录搜索点击样本（词→品 人工标注，供关键词库反哺；scene 区分 top50/工作台/详情页）
    if q:
        _log_search_click(scene, q, _safe(spu), _safe(best.get("商品名称")), len(skus))
    # P4.3 销量趋势：7天日均 vs 30天日均（补货参考回填列，口径见 own_sales_meta.json）
    g["__s7"] = g["7天总销量"].apply(_extract_num).fillna(0.0)
    s30 = float(g["__s"].sum())
    s7 = float(g["__s7"].sum())
    trend = "→平"
    if s30 > 0 and s7 > 0:
        ratio = (s7 / 7.0) / (s30 / 30.0)
        trend = "↑热" if ratio >= 1.2 else ("↓冷" if ratio <= 0.8 else "→平")
    elif s7 > 0:
        trend = "↑热"
    # P4.3 店对店：所选自家店对应仓的 销量（30天/7天）——卡片与对比条同口径，防误读
    own_store_q = (request.args.get("own_store") or "").strip()
    wh30_m, whkw = _wh_map_for_store(own_store_q) if own_store_q else (None, "")
    wh7_m = _wh_map_for_store(own_store_q, col="7天总销量")[0] if own_store_q else None
    wh30_v = round(wh30_m[spu], 0) if (wh30_m is not None and spu in wh30_m) else None
    wh7_v = round(wh7_m[spu], 0) if (wh7_m is not None and spu in wh7_m) else None
    return jsonify({"ok": True, "detail": {
        "品组ID": _safe(spu),
        "品名": _safe(best.get("商品名称")) or _safe(best.get("清洗后商品名")),
        "品牌": _safe(best.get("品牌（抽取）")) or _safe(best.get("品牌名称")),
        "类目": _safe(best.get("类目三级_归一")) or _safe(best.get("美团闪购类目三级")),
        "月销合计": round(s30, 0),
        "7天合计": round(s7, 0),
        "趋势": trend,
        "仓30": wh30_v,
        "仓7": wh7_v,
        "仓名": whkw,
        "规格数": len(g),
        "售价区间": (round(min(prices), 2), round(max(prices), 2)) if prices else (None, None),
        "采购价区间": (round(min(costs), 2), round(max(costs), 2)) if costs else (None, None),
        "图片": _safe(best.get("商品图片url")) or _safe(best.get("图片url")),
        "规格筛选词": spec_words,
        "SKU列表": skus,
    }})


def _comp_skus_of(shop, comp_id):
    """返回该店 商品ID 下的全部 sku 及首个竞品行。"""
    skus, comp_row = [], None
    for sku, row in competitor_lookup.get(shop, {}).items():
        if _safe(row.get("comp_id") or row.get("商品ID")) == comp_id:
            skus.append(sku)
            if comp_row is None:
                comp_row = row
    return skus, comp_row


def _comp_id_of_sku(shop, sku):
    """反向：sku → comp_id（找不到返 ''）。row 可能是 dict 也可能是 pandas Series，
    真值判断必须显式判 None/empty（Series 布尔歧义坑，本轮实测踩中）。"""
    row = competitor_lookup.get(shop, {}).get(sku)
    if row is None:
        return ""
    if hasattr(row, "empty") and row.empty:
        return ""
    return _safe(row.get("comp_id") or row.get("商品ID"))


@app.route("/api/top50_confirm", methods=["POST"])
def api_top50_confirm():
    """销量前50品人工匹配：商品ID ↔ 自家品组（SPU层；纯人工，无机器匹配）。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    comp_id = (data.get("comp_id") or "").strip()
    own_spu = (data.get("own_spu") or "").strip()
    own_name = (data.get("own_name") or "").strip()
    if shop not in competitor_lookup or not comp_id:
        return jsonify({"ok": False, "msg": "参数缺失"}), 400
    if not own_spu:
        return jsonify({"ok": False, "msg": "请先搜索并选择自家品"}), 400
    skus, comp_row = _comp_skus_of(shop, comp_id)
    if not skus or comp_row is None:
        return jsonify({"ok": False, "msg": "竞品不存在"}), 400
    # 写 SPU 匹配层（品对品）
    from summary_library import load_library as _ll, save_library as _sl, upsert_manual as _um
    spu_df, sku_df = _ll()
    spu_df = _um(spu_df, "竞品品ID", shop, {
        "竞品品ID": comp_id,
        "竞品商品名": _safe(comp_row.get("name") or comp_row.get("商品名字") or comp_row.get("清洗后商品名")),
        "竞品品牌": _safe(comp_row.get("品牌（抽取）") or comp_row.get("品牌名称")),
        "竞品型号": _safe(comp_row.get("型号")),
        "竞品类目": _safe(comp_row.get("类目三级_归一") or comp_row.get("category3")),
        "自家品组ID": own_spu, "自家品名": own_name, "置信度": "人工",
    })
    _sl(spu_df, sku_df)
    # 该商品全部规格标记已确认（品对应）
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for sku in skus:
        progress.setdefault("confirmed", {}).setdefault(shop, {})[sku] = {
            "sku_id": "", "spu_only": True, "source": "销量前50人工匹配",
            "time": now,
        }
        progress.setdefault("no_match", {}).get(shop, {}).pop(sku, None)
    save_progress()
    # 学习闭环 + 类目映射（取该自家品组首个SKU）
    try:
        oi = None
        if own_spu in own_df["SPUID"].astype(str).values:
            oi = own_df[own_df["SPUID"].astype(str) == own_spu].iloc[0]
        if oi is not None:
            _confirm_category_map(_safe(comp_row.get("类目三级_归一") or comp_row.get("category3")),
                                  _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级")))
            from learn import record as learn_record, add_keyword as learn_add_kw
            learn_record(shop, skus[0], _safe(comp_row.get("name") or comp_row.get("商品名字")),
                         _safe(oi.get("SKUID")), own_name, source="销量前50人工匹配",
                         comp_brand=_safe(comp_row.get("品牌（抽取）")),
                         comp_model=_safe(comp_row.get("型号")),
                         own_brand=_safe(oi.get("品牌名称")), own_model=_safe(oi.get("型号")))
            learn_add_kw(own_name, _safe(oi.get("SPUID")),
                         _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级")))
    except Exception:
        pass
    return jsonify({"ok": True, "msg": f"已匹配：{own_name}（覆盖 {len(skus)} 个规格）"})


@app.route("/api/top50_nomatch", methods=["POST"])
def api_top50_nomatch():
    """销量前50品：标记无对应 / 恢复待匹配。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    comp_id = (data.get("comp_id") or "").strip()
    restore = data.get("restore") is True
    if shop not in competitor_lookup or not comp_id:
        return jsonify({"ok": False, "msg": "参数缺失"}), 400
    skus, _ = _comp_skus_of(shop, comp_id)
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for sku in skus:
        if restore:
            progress.setdefault("no_match", {}).get(shop, {}).pop(sku, None)
            progress.setdefault("confirmed", {}).get(shop, {}).pop(sku, None)
        else:
            progress.setdefault("no_match", {}).setdefault(shop, {})[sku] = {
                "time": now, "note": "销量前50人工判定无对应",
            }
            progress.setdefault("confirmed", {}).get(shop, {}).pop(sku, None)
    save_progress()
    if restore:
        # 同时移除 SPU 层对应行
        try:
            from summary_library import load_library as _ll, save_library as _sl
            spu_df, sku_df = _ll()
            spu_df = spu_df[~((spu_df["竞店名称"].astype(str) == shop)
                              & (spu_df["竞品品ID"].astype(str).str.strip() == comp_id))]
            _sl(spu_df, sku_df)
        except Exception:
            pass
        return jsonify({"ok": True, "msg": "已恢复为待匹配"})
    return jsonify({"ok": True, "msg": f"已标记无对应（覆盖 {len(skus)} 个规格）"})


@app.route("/api/confirm_undo", methods=["POST"])
def api_confirm_undo():
    """撤销人工对应（应对应错品的情况）。
    body: {shop, comp_id}  → 撤销该品下全部规格的对应
    body: {shop, sku}      → 撤销该单规格的对应
    body: {shop, sku|comp_id, ai: true}  → 撤销 AI 预筛分流（Phase2）：
        只删汇总库 确认来源=AI预筛 的行 + AI 无对应标记（人工行/人工标记不动），
        并把受影响 sku 加拉黑名单（同条不再自动分流），回人工池。
    操作：progress 中该品/规格确认记录删除；汇总库 SPU/SKU 两层该品对应行删；
    无对应标记不动；学习记录不动（保留作为历史）；缓存失效。"""
    try:
        data = request.get_json(force=True) or {}
        shop = (data.get("shop") or "").strip()
        comp_id = (data.get("comp_id") or "").strip()
        sku = (data.get("sku") or "").strip()
        ai_mode = data.get("ai") is True or data.get("source") == "ai"
        if not shop or (not comp_id and not sku):
            return jsonify({"ok": False, "msg": "缺少 shop + comp_id/sku"}), 400
        if shop not in competitor_lookup:
            return jsonify({"ok": False, "msg": f"店铺不存在: {shop}"}), 404

        # 1) 收集要删的 SKU 列表
        if comp_id:
            skus, _ = _comp_skus_of(shop, comp_id)
        else:
            skus = [sku]
        if not skus:
            return jsonify({"ok": False, "msg": "未找到该品/规格"}), 404

        # ===== AI 分流撤销分支（Phase2）=====
        if ai_mode:
            n_lib = 0
            try:
                from summary_library import load_library as _ll, save_library as _sl, CONFIRM_AI
                spu_df, sku_df = _ll()
                spu_pid_col = "竞品品ID" if "竞品品ID" in spu_df.columns else (
                    "竞品商品编号" if "竞品商品编号" in spu_df.columns else "comp_id")
                sku_pid_col = "竞品品ID" if "竞品品ID" in sku_df.columns else (
                    "竞品商品编号" if "竞品商品编号" in sku_df.columns else "comp_id")
                sku_sid_col = "竞品Sku_id" if "竞品Sku_id" in sku_df.columns else (
                    "竞品单品编号" if "竞品单品编号" in sku_df.columns else "sku_id")
                ai_spu = (spu_df["竞店名称"].astype(str) == shop) \
                    & (spu_df["确认来源"].astype(str) == CONFIRM_AI)
                ai_sku = (sku_df["竞店名称"].astype(str) == shop) \
                    & (sku_df["确认来源"].astype(str) == CONFIRM_AI)
                if comp_id:
                    ai_spu &= spu_df[spu_pid_col].astype(str).str.strip() == comp_id
                    ai_sku &= sku_df[sku_pid_col].astype(str).str.strip() == comp_id
                else:
                    ai_spu &= pd.Series(False, index=spu_df.index)  # 单SKU撤销不动SPU行
                    ai_sku &= sku_df[sku_sid_col].astype(str).str.strip() == sku
                n_lib = int(ai_spu.sum() + ai_sku.sum())
                spu_df, sku_df = spu_df[~ai_spu], sku_df[~ai_sku]
                _sl(spu_df, sku_df)
            except Exception as e:
                import traceback
                print("[撤销AI] 汇总库删行失败:", e)
                traceback.print_exc()
            # AI 无对应标记删除（人工标记不动）
            n_nm = 0
            nm = progress.get("no_match", {}).get(shop, {})
            for s in skus:
                rec = nm.get(s)
                if rec and (rec.get("source") == "AI预筛"
                            or "AI预筛" in str(rec.get("note") or "")):
                    nm.pop(s, None)
                    n_nm += 1
            # 拉黑名单：撤销后同条不再自动分流
            try:
                from ai_prescreen import blacklist_add
                blacklist_add(shop, skus)
            except Exception as e:
                print("[撤销AI] 拉黑名单写入失败:", e)
            if n_nm:
                save_progress()
            else:
                _invalidate_stats()
            threading.Thread(target=_warmup_one_shop, args=(shop,), daemon=True).start()
            return jsonify({
                "ok": True,
                "msg": f"已撤销 AI 分流 {len(skus)} 个规格（汇总库AI行删 {n_lib} / AI无对应标记删 {n_nm}），"
                       "已拉黑不再自动分流，回人工池",
                "skus": skus,
            })
        # ===== 人工撤销分支（原有逻辑）=====

        # 2) progress 中删除确认记录
        n_prog = 0
        cm = progress.get("confirmed", {}).get(shop, {})
        for s in skus:
            if s in cm:
                cm.pop(s, None)
                n_prog += 1
        if n_prog:
            save_progress()

        # 3) 汇总库 SPU/SKU 两层删除（异常被吞，确保 progress 删除不被回滚）
        n_lib = 0
        try:
            from summary_library import load_library as _ll, save_library as _sl
            spu_df, sku_df = _ll()
            # 列名兼容
            spu_pid_col = "竞品品ID" if "竞品品ID" in spu_df.columns else (
                "竞品商品编号" if "竞品商品编号" in spu_df.columns else "comp_id")
            sku_pid_col = "竞品品ID" if "竞品品ID" in sku_df.columns else (
                "竞品商品编号" if "竞品商品编号" in sku_df.columns else "comp_id")
            sku_sid_col = "竞品Sku_id" if "竞品Sku_id" in sku_df.columns else (
                "竞品单品编号" if "竞品单品编号" in sku_df.columns else "sku_id")
            if comp_id:
                m = ((spu_df["竞店名称"].astype(str) == shop)
                     & (spu_df[spu_pid_col].astype(str).str.strip() == comp_id))
                n_spu = int(m.sum())
                spu_df = spu_df[~m]
                sku_m = ((sku_df["竞店名称"].astype(str) == shop)
                         & (sku_df[sku_pid_col].astype(str).str.strip() == comp_id))
                n_sku = int(sku_m.sum())
                sku_df = sku_df[~sku_m]
            else:
                n_spu = 0  # 单 SKU 撤销不动 SPU 行（防误伤同品其他规格）
                sku_m = ((sku_df["竞店名称"].astype(str) == shop)
                         & (sku_df[sku_sid_col].astype(str).str.strip() == sku))
                n_sku = int(sku_m.sum())
                sku_df = sku_df[~sku_m]
            _sl(spu_df, sku_df)
            n_lib = n_spu + n_sku
        except Exception as e:
            import traceback
            print("[撤销] 汇总库删行失败:", e)
            traceback.print_exc()

        # 4) 失效缓存
        _invalidate_stats()
        # 5) 后台预热该店
        threading.Thread(target=_warmup_one_shop, args=(shop,), daemon=True).start()

        return jsonify({
            "ok": True,
            "msg": f"已撤销 {len(skus)} 个规格的对应（progress 删 {n_prog} / 汇总库删 {n_lib}）",
            "skus": skus,
        })
    except Exception as e:
        import traceback
        print("[撤销] 顶层异常:", e)
        traceback.print_exc()
        return jsonify({"ok": False, "msg": "撤销异常: " + str(e)}), 500


@app.route("/api/top50/advice", methods=["POST"])
def api_top50_advice():
    """生成 top50 匹配结果的 AI 建议（有对应→优化；无对应→上新）。手动触发才调大模型。"""
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    mock = data.get("mock") is True
    if not shop:
        return jsonify({"ok": False, "msg": "缺少店铺"}), 400
    try:
        from ai_top50 import generate_advice, save_advice_excel
        # 榜单复用面板缓存（build_top50 会重读整张归一化表）；前50 与单独算结果一致
        prods = None
        try:
            key5 = (shop, "top500")
            if key5 not in _compare_cache:
                from top50 import build_top50
                _compare_cache[key5] = _derived("top500", shop,
                                                lambda: build_top50(shop, 500), code=_fp_derived())
            prods = (_compare_cache[key5] or [])[:50]
        except Exception as e:
            print("[AI建议] 取缓存榜单失败，改现算:", e)
        result, infos = generate_advice(shop, mock=mock, prods=prods)
        remember_advice(shop, result)
        path = ""
        try:
            path = save_advice_excel(shop, result)
        except Exception:
            pass
        n_match = sum(1 for a in result if a["状态"] == "已匹配")
        n_new = sum(1 for a in result if a["状态"] == "无对应")
        return jsonify({"ok": True, "advice": result,
                        "msg": f"AI建议已生成（已匹配 {n_match} / 无对应 {n_new} / 待匹配 {len(result) - n_match - n_new}）",
                        "file": path})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/analysis")
def analysis_page():
    return render_template("analysis.html")


# ---------------------------------------------------------------------------
# AI 经营分析（DeepSeek）
# ---------------------------------------------------------------------------
@app.route("/api/ai/config/status")
def api_ai_config_status():
    try:
        from ai_client import load_config, embeddings_config, log_stats, PROVIDER_PRESETS
        cfg = load_config()
        eb = embeddings_config(cfg)
        st = log_stats()
        return jsonify({"ok": True, "configured": bool(cfg.get("api_key")),
                        "model": cfg.get("model", ""), "base_url": cfg.get("base_url", ""),
                        "driver": cfg.get("driver", "real"),
                        "embeddings": bool(eb),
                        "presets": PROVIDER_PRESETS,
                        "spend": {"calls": st.get("calls", 0), "cost_est": st.get("cost_est", 0),
                                  "fail": st.get("fail", 0)}})
    except Exception as e:
        return jsonify({"ok": False, "msg": str(e)})


@app.route("/api/ai/config/test", methods=["POST"])
def api_ai_config_test():
    """连通性测试：用当前配置发一次最小真实调用（约10 token），验证 key/网络/模型。"""
    try:
        from ai_client import call_llm
        reply = call_llm("conn_check",
                         [{"role": "user", "content": "只回答两个字：正常"}],
                         max_tokens=10, timeout=30)
        return jsonify({"ok": True, "msg": "接口连通正常", "reply": str(reply)[:50]})
    except Exception as e:
        return jsonify({"ok": False, "msg": f"接口不通: {e}"}), 500


@app.route("/api/ai/config", methods=["POST"])
def api_ai_config_save():
    """保存 AI 接口配置（OpenAI 兼容可换模型，M7 阶段5）。
    可传 api_key / base_url / model / driver；只改传入字段，保留 yaml 其他段。
    聊天侧禁止改配置（L4 红线）；本路由仅工具页面板调用。"""
    data = request.get_json(force=True) or {}
    api_key = data.get("api_key")
    base_url = data.get("base_url")
    model = data.get("model")
    driver = data.get("driver")
    # 至少要有一项
    if all(x is None or (isinstance(x, str) and not x.strip()) for x in (api_key, base_url, model, driver)):
        return jsonify({"ok": False, "msg": "没有要保存的字段（key/地址/模型/驱动 至少一项）"}), 400
    try:
        from ai_client import save_config_fields
        # key 允许为空串表示“不改 key”——只有显式传了非 None 才写
        save_config_fields(
            api_key=(api_key.strip() if isinstance(api_key, str) else None),
            base_url=(base_url.strip() if isinstance(base_url, str) else None),
            model=(model.strip() if isinstance(model, str) else None),
            driver=(driver.strip() if isinstance(driver, str) else None),
        )
        return jsonify({"ok": True, "msg": "AI 接口配置已保存（仅存本机 config/ai_config.yaml）"})
    except ValueError as e:
        return jsonify({"ok": False, "msg": str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "msg": "保存失败: " + str(e)}), 500


@app.route("/api/ai/analysis", methods=["POST"])
def api_ai_analysis():
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    directions = data.get("directions") or ["hot", "price", "display", "insight"]
    if not shops:
        return jsonify({"ok": False, "msg": "请至少选择一家店"}), 400
    try:
        from ai_analysis import generate_report
        # 生成时只写草稿（面板展示），正式文件由"导出"按钮触发
        report, stats, path, html_path = generate_report(shops, directions)
        report_html = ""
        if os.path.exists(html_path):
            with open(html_path, "r", encoding="utf-8") as f:
                report_html = f.read()
        return jsonify({"ok": True, "report": report, "report_html": report_html,
                        "path": path, "html_path": html_path,
                        "sections": stats.get("sections", []),
                        "structured": {k: v for k, v in stats.items()
                                       if k in ("hot", "gap", "cat_gap", "band", "compare")},
                        "msg": f"报告已生成（{len(shops)} 家店，板块：{'、'.join(stats.get('sections', []))}）。如需保存，请点「导出」"})
    except Exception as e:
        return jsonify({"ok": False, "msg": "生成失败: " + str(e)}), 500


@app.route("/api/ai/export", methods=["POST"])
def api_ai_export():
    """把最近一次生成的报告草稿导出为正式文件（AI经营分析报告.md/.html）。"""
    try:
        from ai_analysis import export_draft_to_report
        md_path, html_path = export_draft_to_report()
        if not md_path:
            return jsonify({"ok": False, "msg": "没有可导出的报告，请先生成 AI 经营分析"}), 400
        return jsonify({"ok": True, "msg": "已导出：AI经营分析报告.html（另含 .md 原文）",
                        "md_path": md_path, "html_path": html_path})
    except Exception as e:
        return jsonify({"ok": False, "msg": "导出失败: " + str(e)}), 500


@app.route("/api/ai/category_diag", methods=["POST"])
def api_ai_category_diag():
    """品类诊断卡片墙：聚合品类数据 + AI 诊断。"""
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    use_ai = data.get("use_ai", True)  # 可选：False 时只返回本地数据（不调 AI）
    if not shops:
        return jsonify({"ok": False, "msg": "请至少选择一家店"}), 400
    try:
        from ai_analysis import build_category_card_data, call_category_diagnosis, load_config

        # Step 1: 构建品类指标数据（纯本地计算，秒级）
        card_data = build_category_card_data(shops)
        if not card_data["cards"]:
            return jsonify({"ok": False, "msg": "所选店铺无品类数据，请先运行传统分析"}), 400

        # Step 2: 调 AI 生成诊断（约 30-60 秒）
        if use_ai:
            config = load_config()
            if config.get("api_key"):
                try:
                    diagnoses = call_category_diagnosis(card_data, config)
                    diag_map = {d["cat"]: d for d in diagnoses if isinstance(d, dict)}
                    for card in card_data["cards"]:
                        diag = diag_map.get(card["cat"], {})
                        card["diagnosis"] = diag.get("diagnosis", "")
                        card["actions"] = diag.get("actions", [])
                except Exception as ai_err:
                    # AI 失败不阻塞：降级为无诊断文本
                    for card in card_data["cards"]:
                        card["diagnosis"] = f"（AI 诊断生成失败：{ai_err}）"
            else:
                for card in card_data["cards"]:
                    card["diagnosis"] = "（未配置 AI 接口密钥，请在上方配置 DeepSeek Key）"

        return jsonify({
            "ok": True,
            "cards": card_data["cards"],
            "summary": card_data["summary"],
            "msg": f"品类诊断已生成（{len(card_data['cards'])} 个品类）",
        })
    except Exception as e:
        return jsonify({"ok": False, "msg": "生成失败: " + str(e)}), 500


@app.route("/api/analysis/shops")
def api_analysis_shops():
    """店列表：质量报告指标 + 待确认数 + 已保存选择。"""
    shops = []
    if os.path.exists(SHOP_REPORT_FILE):
        df = pd.read_excel(SHOP_REPORT_FILE)
        for _, r in df.iterrows():
            shops.append({
                "name": _safe(r.get("店名")),
                "标准条码率%": _safe(r.get("标准条码率%")),
                "类目3率%": _safe(r.get("类目3率%")),
                "质量分": _safe(r.get("质量分")),
                "建议优先级": _safe(r.get("建议优先级")),
                "total": 0, "pending": 0,
            })
    if pool_df is not None and len(pool_df):
        for s in shops:
            g = pool_df[pool_df["竞店名称"] == s["name"]]
            s["total"] = int(len(g))
            conf_set = set(progress.get("confirmed", {}).get(s["name"], {}).keys())
            nm_set = set(progress.get("no_match", {}).get(s["name"], {}).keys())
            done = sum(1 for sku in g["竞品Sku_id"] if sku in conf_set or sku in nm_set)
            s["pending"] = int(len(g) - done)
    saved = load_analysis_shops()
    for s in shops:
        s["selected"] = (s["name"] in saved) if saved is not None else True
    return jsonify({"shops": shops})


@app.route("/api/analysis/save", methods=["POST"])
def api_analysis_save():
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    import yaml
    os.makedirs(os.path.dirname(ANALYSIS_CONFIG), exist_ok=True)
    with open(ANALYSIS_CONFIG, "w", encoding="utf-8") as f:
        yaml.safe_dump({"shops": shops}, f, allow_unicode=True, sort_keys=False)
    return jsonify({"ok": True, "msg": f"已保存 {len(shops)} 家店的选择"})


@app.route("/api/analysis/run", methods=["POST"])
def api_analysis_run():
    data = request.get_json(force=True)
    shops = data.get("shops") or []
    if not shops:
        return jsonify({"ok": False, "msg": "请至少选择一家店"}), 400
    import subprocess
    import sys as _sys
    shop_arg = ",".join(shops)
    try:
        r1 = subprocess.run(
            [_sys.executable, os.path.join(SCRIPTS_DIR, "build_price_matrix.py"),
             "--shops", shop_arg],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=900, cwd=BASE_DIR)
        r2 = subprocess.run(
            [_sys.executable, os.path.join(SCRIPTS_DIR, "build_competition_analysis.py"),
             "--shops", shop_arg],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=1200, cwd=BASE_DIR)
        if r1.returncode != 0 or r2.returncode != 0:
            err = (r1.stderr or "") + (r2.stderr or "")
            return jsonify({"ok": False, "msg": "分析失败: " + err[-500:]}), 500
    except Exception as e:
        return jsonify({"ok": False, "msg": "分析异常: " + str(e)}), 500
    overview = []
    try:
        if os.path.exists(COMPETE_REPORT_FILE):
            ov = pd.read_excel(COMPETE_REPORT_FILE, sheet_name="总览")
            for _, r in ov.iterrows():
                overview.append({"指标": _safe(r.get("指标")), "数值": _safe(r.get("数值"))})
    except Exception:
        pass
    return jsonify({"ok": True, "msg": f"分析完成（{len(shops)} 家店）",
                    "overview": overview,
                    "reports": ["output/竞争分析报告.xlsx", "output/竞品价格矩阵.xlsx"]})


# ---------------------------------------------------------------------------
# 反向查询：本店SKUID -> 已确认对齐的所有竞品品（一对多）
# ---------------------------------------------------------------------------
@app.route("/reverse")
def reverse_page():
    return render_template("reverse.html")


@app.route("/api/reverse")
def api_reverse():
    # 兼容旧参数 sku=xxx（默认 type=sku），新参数 type=xxx&value=xxx
    qtype = (request.args.get("type") or "sku").strip().lower()
    value = (request.args.get("value") or "").strip()
    if not value:
        # 兼容旧调用 /api/reverse?sku=xxx
        legacy_sku = request.args.get("sku", "").strip()
        if legacy_sku:
            value = legacy_sku
            qtype = "sku"

    if not value:
        return jsonify({"count": 0, "msg": "请输入查询内容"})

    if qtype not in ("sku", "spu", "category", "name"):
        return jsonify({"count": 0, "msg": "不支持的查询类型: " + qtype})

    # 实时读取两层汇总库 SKU匹配层（dtype=str 防止 19 位 ID 精度丢失）
    _, sku_df = load_library()
    summary_df = sku_df
    if len(summary_df) == 0:
        return jsonify({"count": 0, "msg": "汇总库暂无数据"})

    # 关键 ID 列统一转字符串
    for col in ["自家SKUID", "竞品Sku_id", "竞品品ID", "自家SPUID"]:
        if col in summary_df.columns:
            summary_df[col] = summary_df[col].apply(_safe)

    own_sku_col = "自家SKUID" if "自家SKUID" in summary_df.columns else None
    if own_sku_col is None:
        return jsonify({"count": 0, "msg": "汇总库缺少 自家SKUID 列"})
    own_spu_col = "自家SPUID" if "自家SPUID" in summary_df.columns else None
    own_name_col = "自家商品名" if "自家商品名" in summary_df.columns else None
    # 类目：SKU层无类目列，从自家表实时反查

    # 构建结果行（统一格式）
    def build_result(row):
        comp_sku = _safe(row.get("竞品Sku_id"))
        shop = _safe(row.get("竞店名称"))
        comp_info = get_competitor_info(comp_sku, shop or None)
        if not comp_info and shop not in competitor_lookup:
            for s in competitor_lookup:
                comp_info = get_competitor_info(comp_sku, s)
                if comp_info:
                    shop = s
                    break
        own_sku = _safe(row.get(own_sku_col)) if own_sku_col else ""
        oi = own_lookup.get(own_sku, {}) if own_sku else {}
        return {
            "竞店名": shop or comp_info.get("店名", ""),
            "竞品品ID": _safe(row.get("竞品品ID")),
            "竞品Sku_id": comp_sku,
            "竞品商品名": _safe(row.get("竞品商品名")),
            "竞品图片url": comp_info.get("图片url", ""),
            "竞品到手价": comp_info.get("价格", ""),
            "竞品月售": comp_info.get("月售", ""),
            "竞品类目": comp_info.get("三级类", "") or _safe(row.get("竞品类目")),
            "匹配依据": _safe(row.get("匹配依据")),
            "匹配层": _safe(row.get("匹配层")),
            "置信度": _safe(row.get("置信度")),
            "确认来源": _safe(row.get("确认来源")),
            "状态": _safe(row.get("状态")),
            "自家SPUID": _safe(row.get(own_spu_col)) if own_spu_col else "",
            "自家SKUID": own_sku,
            "自家商品名": _safe(row.get(own_name_col)) if own_name_col else "",
            "自家类目": _safe(oi.get("类目三级_归一")) or _safe(oi.get("美团闪购类目三级")),
            # 自家商品对比信息（售价/采购价/月销/规格/条码/有图）
            "自家售价": _safe(oi.get("美团闪购渠道售价")) or _safe(oi.get("售价")),
            "自家采购价": _safe(oi.get("采购价")),
            "自家月销": _safe(oi.get("月销量")) or _safe(oi.get("月售")),
            "自家规格": _safe(oi.get("规格名称")) or _safe(oi.get("规格")),
            "自家条码": _safe(oi.get("商品条码")),
            "自家有图": _safe(oi.get("是否有图片")),
            "自家品牌": _safe(oi.get("品牌名称")) or _safe(oi.get("品牌（抽取）")),
        }

    # ----- SKU 模式：精确 + 模糊匹配（contains） -----
    if qtype == "sku":
        mask = summary_df[own_sku_col].astype(str).str.contains(value, case=False, na=False)
        matched = summary_df[mask]
        if len(matched) == 0:
            return jsonify({"count": 0, "msg": "该SKU暂无已确认对齐的竞品品"})
        results = [build_result(row) for _, row in matched.iterrows()]
        return jsonify({
            "count": len(results),
            "results": results,
            "query_type": "sku",
            "query_value": value,
        })

    # ----- SPU 模式：返回该SPU下所有SKU对应的竞品品，按SKU分组 -----
    if qtype == "spu":
        if own_spu_col is None:
            return jsonify({"count": 0, "msg": "汇总库缺少 自家SPUID 列"})
        mask = summary_df[own_spu_col].astype(str).str.contains(value, case=False, na=False)
        matched = summary_df[mask]
        if len(matched) == 0:
            return jsonify({"count": 0, "msg": "该SPU暂无已确认对齐的竞品品"})

        results = [build_result(row) for _, row in matched.iterrows()]
        # 按 自家SKUID 分组
        groups = {}
        group_order = []
        for r in results:
            sku = r["自家SKUID"] or "(无SKUID)"
            if sku not in groups:
                groups[sku] = []
                group_order.append(sku)
            groups[sku].append(r)
        grouped = [{
            "自家SKUID": sku,
            "自家商品名": groups[sku][0].get("自家商品名", ""),
            "自家SPUID": groups[sku][0].get("自家SPUID", ""),
            "items": groups[sku],
        } for sku in group_order]
        return jsonify({
            "count": len(results),
            "group_count": len(grouped),
            "groups": grouped,
            "results": results,
            "query_type": "spu",
            "query_value": value,
        })

    # ----- 类目模式：返回该类目下所有已确认对齐记录 + 统计 -----
    if qtype == "category":
        # 精确匹配类目：先从自家表取出该类目全部 SKUID（类目映射表里是三级类目全名）
        cat_col_own = None
        if own_df is not None:
            cat_col_own = ("类目三级_归一" if "类目三级_归一" in own_df.columns else
                           ("美团闪购类目三级" if "美团闪购类目三级" in own_df.columns else None))
        if cat_col_own is None:
            return jsonify({"count": 0, "msg": "自家表缺少类目列"})
        own_sub = own_df[own_df[cat_col_own].astype(str) == value]
        total_in_own = own_sub["SKUID"].nunique() if "SKUID" in own_sub.columns else len(own_sub)
        cat_skus = set(own_sub["SKUID"].astype(str)) if "SKUID" in own_sub.columns else set()
        matched = summary_df[summary_df[own_sku_col].astype(str).isin(cat_skus)]
        results = [build_result(row) for _, row in matched.iterrows()]

        # 已对齐数 = matched 中 distinct 自家SKUID
        aligned_sku_count = matched[own_sku_col].nunique() if len(matched) > 0 else 0
        unaligned_count = max(total_in_own - aligned_sku_count, 0)
        align_rate = round(aligned_sku_count / total_in_own * 100, 1) if total_in_own else 0.0

        return jsonify({
            "count": len(results),
            "results": results,
            "stats": {
                "total_in_own": int(total_in_own),
                "aligned_sku_count": int(aligned_sku_count),
                "unaligned_count": int(unaligned_count),
                "align_rate": align_rate,
            },
            "query_type": "category",
            "query_value": value,
        })

    # ----- 商品名模式：自家商品名 contains 关键词 -----
    if qtype == "name":
        # 关键词搜索：空格分词逐词 AND（如「自攻螺丝 M4」）；在 自家商品名/竞品商品名 匹配
        from search_util import kw_mask
        kwm = kw_mask(summary_df, value, ["自家商品名", "竞品商品名"])
        if kwm is None:
            return jsonify({"count": 0, "msg": "请输入查询内容"})
        matched = summary_df[kwm]
        if len(matched) == 0:
            return jsonify({"count": 0, "msg": "未找到商品名含这些关键词的已对齐记录"})
        results = [build_result(row) for _, row in matched.iterrows()]
        return jsonify({
            "count": len(results),
            "results": results,
            "query_type": "name",
            "query_value": value,
        })

    return jsonify({"count": 0, "msg": "未知查询类型"})


# ---------------------------------------------------------------------------
# 导入映射页：上传新竞品表 -> 猜映射 -> 人工确认 -> 提取列输出标准格式
# ---------------------------------------------------------------------------
@app.route("/import")
def import_page():
    return render_template("import.html")


def _load_import_records():
    """读取导入记录 {文件md5: {last:{time,file_name}, imported:{time,source_name}或None}}。"""
    if os.path.exists(IMPORT_RECORDS_FILE):
        try:
            with open(IMPORT_RECORDS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_import_records(rec):
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    with open(IMPORT_RECORDS_FILE, "w", encoding="utf-8") as f:
        json.dump(rec, f, ensure_ascii=False, indent=2)


def _file_md5(path):
    import hashlib
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def _check_duplicate(saved_path):
    """同一文件（md5 相同）之前成功导入过则返回记录 dict，否则 None。"""
    try:
        fh = _file_md5(saved_path)
        records = _load_import_records()
        dup = records.get(fh, {}).get("imported")
        if dup:
            return {
                "time": dup.get("time", ""),
                "source_name": dup.get("source_name", ""),
                "file_name": dup.get("file_name", ""),
            }
    except Exception:
        pass
    return None


@app.route("/import/own_store", methods=["POST"])
def import_own_store():
    """自家门店导出表导入（独立分支：不进竞品对齐，SPUID 与总库同源直接 join）。
    方案见 docs/多自家门店比对方案.md。"""
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "未收到文件"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"ok": False, "msg": "文件名为空"}), 400
    orig = f.filename
    if not (orig.lower().endswith(".xlsx") or orig.lower().endswith(".xls")):
        return jsonify({"ok": False, "msg": "仅支持 .xlsx / .xls 文件"}), 400
    ext = ".xls" if orig.lower().endswith(".xls") else ".xlsx"
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    tmp = os.path.join(UPLOAD_DIR, "ownstore_" + uuid.uuid4().hex[:12] + ext)
    f.save(tmp)
    try:
        import own_store_lib
        result = own_store_lib.import_store(tmp)
        if result.get("ok"):
            own_store_lib.invalidate()
        return jsonify(result)
    except Exception as e:
        return jsonify({"ok": False, "msg": "导入失败: " + str(e)}), 500
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


@app.route("/import/upload", methods=["POST"])
def import_upload():
    """上传 Excel 文件 -> scan_columns + guess_mapping -> 返回映射预览。"""
    if "file" not in request.files:
        return jsonify({"ok": False, "msg": "未收到文件"}), 400
    f = request.files["file"]
    if not f.filename:
        return jsonify({"ok": False, "msg": "文件名为空"}), 400
    source_name = (request.form.get("source_name") or "").strip()
    if not source_name:
        source_name = "未命名来源_" + datetime.now().strftime("%Y%m%d_%H%M%S")

    # 批量导入模式：未填店名时自动从文件名提取（如 "竞品_鑫宗源五金_美团导出.xlsx" -> "鑫宗源五金"）
    if request.form.get("auto_shop") == "1" and source_name.startswith("未命名来源_"):
        try:
            from build_comp_library import extract_shop_name
            auto_name = extract_shop_name(f.filename)
            if auto_name:
                source_name = auto_name
        except Exception:
            pass

    # 先用原始文件名校验扩展名（secure_filename 会剥离中文导致扩展名丢失）
    orig_name = f.filename
    orig_lower = orig_name.lower()
    if not (orig_lower.endswith(".xlsx") or orig_lower.endswith(".xls")):
        return jsonify({"ok": False, "msg": "仅支持 .xlsx / .xls 文件"}), 400
    ext = ".xls" if orig_lower.endswith(".xls") else ".xlsx"

    # 保存到 config/import_uploads/<uuid>.<ext>（避免中文文件名导致 secure_filename 异常）
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file_id = uuid.uuid4().hex[:12]
    safe_name = file_id + ext
    saved_path = os.path.join(UPLOAD_DIR, safe_name)
    f.save(saved_path)

    # 记录文件指纹（重复导入检测）
    file_hash = _file_md5(saved_path)
    records = _load_import_records()
    records.setdefault(file_hash, {})
    records[file_hash]["last"] = {
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_name": orig_name,
    }
    _save_import_records(records)

    try:
        columns = import_adapter.scan_columns(saved_path)
    except Exception as e:
        return jsonify({"ok": False, "msg": "扫描列失败: " + str(e)}), 500

    mapping = import_adapter.guess_mapping(columns)

    # 尝试加载已有同名 profile 预填
    existing = import_adapter.load_profile(source_name)
    if existing and existing.get("mapping"):
        # 已有配置覆盖猜映射结果
        for col in mapping:
            if col in existing["mapping"]:
                mapping[col] = existing["mapping"][col]

    # 构造列预览（每列附是否猜中）
    col_preview = []
    for col in columns:
        field = mapping.get(col, import_adapter.UNSPECIFIED)
        col_preview.append({
            "excel_col": _safe(col),
            "field": field,
            "guessed": field != import_adapter.UNSPECIFIED,
        })

    hit_count = sum(1 for c in col_preview if c["guessed"])

    return jsonify({
        "ok": True,
        "file_id": file_id,
        "file_name": safe_name,
        "file_hash": file_hash,
        "source_name": source_name,
        "columns": col_preview,
        "hit_count": hit_count,
        "total_count": len(col_preview),
        "standard_fields": [
            {"value": f, "label": import_adapter.FIELD_CN.get(f, f)}
            for f in import_adapter.STANDARD_FIELDS
        ],
    })


def _resolve_upload_path(file_id, file_name):
    """根据 file_id + file_name 还原上传文件路径（兼容旧格式 file_id_name 和新格式 file_id.ext）。"""
    if not file_name:
        return None
    # 新格式：file_name 已是完整保存名（file_id + ext）
    p1 = os.path.join(UPLOAD_DIR, file_name)
    if os.path.exists(p1):
        return p1
    # 旧格式兼容：file_id + _ + file_name
    if file_id:
        p2 = os.path.join(UPLOAD_DIR, f"{file_id}_{file_name}")
        if os.path.exists(p2):
            return p2
    return None


@app.route("/import/confirm", methods=["POST"])
def import_confirm():
    """接收用户确认的映射 -> save_profile + extract_columns。"""
    data = request.get_json(force=True)
    file_id = (data.get("file_id") or "").strip()
    file_name = (data.get("file_name") or "").strip()
    source_name = (data.get("source_name") or "").strip()
    mapping = data.get("mapping")  # {excel_col: 标准字段或"待指定"或"忽略"}

    if not file_id or not file_name:
        return jsonify({"ok": False, "msg": "缺少文件标识，请重新上传"}), 400
    if not source_name:
        return jsonify({"ok": False, "msg": "请填写来源名"}), 400
    if not isinstance(mapping, dict) or not mapping:
        return jsonify({"ok": False, "msg": "映射数据为空"}), 400

    # 还原上传文件路径
    saved_path = _resolve_upload_path(file_id, file_name)
    if not saved_path:
        return jsonify({"ok": False, "msg": "上传文件已失效，请重新上传"}), 400

    # 重复导入检测：同一文件（md5 相同）之前成功导入过
    duplicate = _check_duplicate(saved_path)

    # 过滤掉 "忽略" 的列（视为待指定，不写入 profile，extract 也不保留）
    clean_mapping = {
        col: (field if field != "忽略" else import_adapter.UNSPECIFIED)
        for col, field in mapping.items()
    }

    # 保存 profile
    try:
        profile_path = import_adapter.save_profile(source_name, clean_mapping)
    except Exception as e:
        return jsonify({"ok": False, "msg": "保存配置失败: " + str(e)}), 500

    # 提取列输出标准格式表
    output_name = f"竞品_标准化_{source_name}_{file_id[:8]}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_name)
    try:
        result_path = import_adapter.extract_columns(saved_path, clean_mapping, output_path)
    except ValueError as e:
        return jsonify({"ok": False, "msg": "提取失败: " + str(e)}), 400
    except Exception as e:
        return jsonify({"ok": False, "msg": "提取异常: " + str(e)}), 500

    # 读取输出表的行数和列
    try:
        out_df = pd.read_excel(result_path, dtype=str)
        row_count = len(out_df)
        out_cols = list(out_df.columns)
    except Exception:
        row_count = -1
        out_cols = []

    # 同店检测：导入的是已有店铺的数据更新？
    same_shop = None
    try:
        from build_comp_library import _detect_existing_shop
        sum_f = os.path.join(OUTPUT_DIR, "竞店标准库", "多竞店汇总库.xlsx")
        if os.path.exists(sum_f):
            sdf = pd.read_excel(sum_f, dtype=str)
            sdf = sdf.rename(columns={k: v for k, v in import_adapter.CN_STD.items() if k in sdf.columns})
            same_shop = _detect_existing_shop(source_name, sdf)
    except Exception:
        same_shop = None

    # ---- 自动入库 + 对齐（新店=首次对齐；同店=数据更新重跑对齐）----
    # 两层汇总库/人工池由对齐引擎自动同步；用户无需再手动跑命令
    import_msg, align_rows = "", 0
    sales_health = None
    try:
        from build_comp_library import add_one
        add_result = add_one(saved_path, force=True, shop=source_name)
        align_rows = len(add_result) if add_result is not None else 0
        import_msg = f"已入库并完成对齐（{align_rows} 行，店铺: {source_name}）"
        # 重新加载数据（新店人工池/汇总库已生成，确认页即时可见）
        load_data()
        load_progress()
        # 缓存失效：旧 _compare_cache / _lib_cache / _stats_cache 全部丢弃，
        # 避免首页/竞品库/browse 显示陈旧结论（用户反馈"引入新店后旧店结论清零"的根因）。
        _invalidate_stats()
        # 后台预热全部店（不止新导入店）：_invalidate_stats 清的是全部店的缓存，
        # 只预热新店会让其他店首次访问同步重算（top500 2-7s/店）——
        # 导入后用户点任何店都应秒开（性能目标：导入慢可以，用页面慢不行）。
        _warmup_all_shops_bg()
        # 销量健康度：sales 列是否被正确归一化（用户反馈"热销 50 榜单没有销量"）。
        try:
            from top50 import health_check as _hc
            sales_health = _hc(source_name)
        except Exception as _e:
            sales_health = {"ok": False, "msg": "销量检查失败: " + str(_e)}
        # 关键词库自动反哺（机器学习逻辑持续优化）：导入后跑一次，
        # 把新产生的搜索点击/品名高频词写回关键词库，无人工干预。
        threading.Thread(target=_auto_apply_keywords, daemon=True).start()
    except Exception as e:
        import_msg = "入库/对齐未完成: " + str(e)

    # 记录本次成功导入（供重复导入检测）
    # 注：此前用未定义变量 fh 致 NameError 被吞，imported 记录从未写入，
    #     重复导入检测（_check_duplicate）永久失效。此处直接对已还原的上传
    #     文件重算 md5，不依赖前端回传。
    try:
        file_hash = _file_md5(saved_path)
        records = _load_import_records()
        records.setdefault(file_hash, {})["imported"] = {
            "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source_name": source_name,
            "file_name": file_name,
        }
        _save_import_records(records)
    except Exception:
        pass

    # P1.1：本次 vs 上次 变动摘要（同店更新时给出 降价/涨价/新品/下架/被打价 计数；
    # 首次导入=基线建立，note 说明）。jsonl 已被 add_one 追加新快照，diff 即最新两期。
    changes = None
    try:
        import price_diff as _pdiff
        _rows, _meta = _pdiff.diff_shop(source_name, with_undercut=True)
        _cnt = {}
        for _r in _rows:
            _k = _r.get("变动") or "?"
            _cnt[_k] = _cnt.get(_k, 0) + 1
        changes = {"counts": _cnt, "rows": len(_rows),
                   "prev_t": _meta.get("prev_t", ""), "cur_t": _meta.get("cur_t", ""),
                   "note": _meta.get("note", "")}
    except Exception as _e:
        print("[导入] 变动摘要失败:", _e)

    return jsonify({
        "ok": True,
        "msg": "导入成功",
        "profile_path": profile_path,
        "output_path": result_path,
        "output_name": output_name,
        "row_count": row_count,
        "columns": out_cols,
        "same_shop": same_shop,
        "duplicate": duplicate,
        "align_msg": import_msg,
        "align_rows": align_rows,
        "sales_health": sales_health,
        "changes": changes,
    })


# ---------------------------------------------------------------------------
# 类目列表 API：从类目映射表读取
# ---------------------------------------------------------------------------
@app.route("/api/categories")
def api_categories():
    if not os.path.exists(CATEGORY_MAP_FILE):
        return jsonify({"ok": False, "msg": "类目映射表不存在", "categories": []})
    try:
        cat_df = pd.read_excel(CATEGORY_MAP_FILE, dtype=str)
    except Exception as e:
        return jsonify({"ok": False, "msg": "读取类目映射表失败: " + str(e), "categories": []})

    cat_col = "三级类目" if "三级类目" in cat_df.columns else (
        cat_df.columns[0] if len(cat_df.columns) > 0 else None
    )
    if cat_col is None:
        return jsonify({"ok": False, "msg": "类目映射表无列", "categories": []})

    cats = [_safe(c) for c in cat_df[cat_col].dropna().unique().tolist() if _safe(c)]
    cats = sorted(cats)
    return jsonify({"ok": True, "categories": cats, "count": len(cats)})


# ---------------------------------------------------------------------------
# 启动
# ---------------------------------------------------------------------------
PANEL_PORT = int(os.environ.get("PANEL_PORT", "5000"))
# 代码指纹：web_panel.py + templates/*.html 内容 md5 前 12 位。
# 模板也纳入（Flask 生产模式不自动重载模板，指纹不含模板会导致改模板后旧进程不替换）。
def _panel_version():
    # 算法只有一处定义：version_stamp.panel_version（覆盖 scripts/*.py + templates/*.html）
    from version_stamp import panel_version
    return panel_version()


PANEL_VERSION = _panel_version()


@app.route("/api/version")
def api_version():
    """面板版本指纹（启动器健康检查用）。"""
    return jsonify({"version": PANEL_VERSION, "pid": os.getpid(),
                    "port": PANEL_PORT})


# ---------------------------------------------------------------------------
# M1 小助手：核心分析场景（本地聚合 → 反空谈提示词 → DeepSeek 增量洞察）
# 数据源全部复用既有缓存/引擎产物（品卡/类目统计/变动明细/对齐明细）。
# ---------------------------------------------------------------------------

# 取品卡/对齐明细缓存；未命中时若预热在跑则返回 None(202)，否则前台同步算一次
# （预热失败会退避 10 分钟，不能让助手端点一直 202——同步算兜底，慢但能出结果）
def _asst_cached(shop, key):
    if (shop, key) in _compare_cache:
        return _compare_cache[(shop, key)]
    st, ts = _warmup_state.get(shop, ("", 0))
    if st == "running":
        return None   # 该店正在现算，别并发再算一遍
    try:
        if key == "product":
            from build_product_analysis import build_product_analysis
            return _derived("product", shop, lambda: build_product_analysis(shop),
                            code=_fp_derived())
        if key == "detail":
            from build_comparison import build_comparison_data
            return _derived("detail", shop, lambda: build_comparison_data(shop),
                            code=_fp_derived())
        return None
    except Exception as e:
        print("[助手] 同步构建失败:", shop, key, e)
        return None


def _asst_payload(scene, shop):
    """按场景聚合给 AI 的数据。返回 (rows, cols, note)。
    只读缓存引擎产物，不含 dashboard 展示层逻辑。"""
    from assistant_brain import SCENE_HEAD  # noqa
    if scene == "top":
        got = _asst_cached(shop, "product")
        if got is None:
            return None, [], "产品分析计算中（约30秒，后台预热），稍后再点"
        pdf = got[0]
        cols = ["品名", "结论", "竞品总月售", "竞品有效月售", "竞品收藏量",
                "自家月销", "在卖品牌数", "竞品品牌", "建议"]
        rows = pdf.head(22).to_dict("records")
        return rows, cols, f"{shop} · 产品分析品卡前22行（结论引擎已排序）"
    if scene == "cat":
        got = _asst_cached(shop, "product")
        if got is None:
            return None, [], "类目统计计算中（约30秒，后台预热），稍后再点"
        cdf = got[1]
        if not len(cdf):
            return [], [], f"{shop} 无类目统计（可能无有效对应品）"
        cols = ["类目", "品数", "机会品", "防守品", "竞品总月售", "竞品有效月售",
                "竞品收藏量", "自家月销合计"]
        rows = cdf.head(24).to_dict("records")
        return rows, cols, f"{shop} · 类目聚合（按竞品总月售降序，前24类）"
    if scene == "price":
        try:
            import price_diff
            rows, meta = price_diff.diff_shop(shop, with_undercut=True)
        except Exception as e:
            return None, [], f"变动计算失败：{e}"
        if not rows:
            obs = _shop_obs_meta().get(shop) or {}
            snaps = int(obs.get("snaps") or 0)
            until = str(obs.get("last_t") or "?")[:16]
            if snaps < 2:
                return [], [], (f"{shop} 只有 {snaps} 次价格观测——变动类结论要等同一家竞店"
                                f"再导入一次才能算（当前数据截至 {until}）。这是数据没到，不是功能坏了。")
            return [], [], f"{shop} 两次观测之间没有价格/上下架变动（当前数据截至 {until}）"
        cols = ["变动", "品名", "前值", "现值", "幅度", "月销量", "备注"]
        return rows[:24], cols, (f"{shop} · 变动明细（{meta.get('prev_t', '?')} → "
                                 f"{meta.get('cur_t', '?')}）前24行，已按 警报/降价优先排序")
    if scene == "match":
        got = _asst_cached(shop, "detail")
        if got is None:
            return None, [], "对齐明细计算中（约30秒，后台预热），稍后再点"
        df = got[0]
        if df is None or not len(df):
            return [], [], f"{shop} 暂无已对应明细"
        # 判据与探测器⑤同口径：销量按 有效销量 = 成交 + λ×收藏 折算。
        # 这里以前写 float(竞品月售)，而该列是「月售35 / 35人想买」混排的文本，
        # 解析异常被 continue 静默吞掉 → 示例竞店B店 3089 行全丢、页面对 4/5 家店报"体检干净"。
        import insight_scan as _ISc
        _th_c, _th_o = _ISc.TH["diverge_comp"], _ISc.TH["diverge_own"]
        cand, n_hit, n_from_intent = [], 0, 0
        for r in df.to_dict("records"):
            raw = r.get("竞品月售")
            xb = str(r.get("跨品牌") or "").strip() == "是"
            diverge = (_ST.eff_sales_ge(raw, _th_c)
                       and _ST.num(r.get("自家月销")) <= _th_o)
            if _ST.is_intent(raw) and diverge:
                n_from_intent += 1              # 靠收藏折算才够门槛：照样进候选，但要标出来
            if xb or diverge:
                n_hit += 1
                if len(cand) < 30:
                    cand.append(r)
        _intent_note = (f"；其中 {n_from_intent} 行靠收藏折算过门槛"
                        f"（λ={_ST.INTENT_WEIGHT:g}，收藏算得轻但不是不算）"
                        if n_from_intent else "")
        if not cand:
            return [], [], (f"{shop} 已对应 {len(df)} 行：没有跨品牌对应，也没有"
                            f"「有效销量≥{_th_c} 而我家≤{_th_o}」的销量背离行，体检干净")
        cols = ["竞品商品名", "自家商品名", "竞品品牌", "自家品牌",
                "竞品月售", "自家月销", "跨品牌", "语义层级", "置信度", "结论"]
        return cand, cols, (f"{shop} · 可疑对应抽样（命中 {n_hit} 行，取前 {len(cand)} 行）："
                            f"跨品牌 或 竞品有效销量≥{_th_c} 而自家零销" + _intent_note)
    if scene == "plan":
        import insight_store as IS
        live = [i for i in IS.list_items()["items"]
                if i.get("present", True) and i.get("status") == "new"
                and (i.get("scope") in (shop, "跨店", "全部", ""))]
        if not live:
            return [], [], (f"{shop} 的洞察台账没有未处理条目——要么已处理完，"
                            f"要么还没点「重新扫描」。台账存在 output/AI洞察.json。")
        rows = [{"结论": i.get("title"), "层级": i.get("layer"), "范围": i.get("scope"),
                 "探测器": i.get("detector"),
                 "数字证据": "；".join(f"{k}={v}" for k, v in (i.get("metrics") or {}).items()
                                    if k != "算式")[:170],
                 "连续未处理期数": i.get("periods", 1)}
                for i in live[:20]]
        cols = ["结论", "层级", "范围", "数字证据", "连续未处理期数"]
        return rows, cols, (f"洞察台账未处理条目 {len(rows)} 条（规则层跑全量筛过，"
                            f"数字可复算；这是给模型的取舍材料，不是原始明细）")
    return None, [], "未知场景"


@app.route("/api/assistant/meta")
def api_assistant_meta():
    """场景清单 + 可用店铺 + AI 配置状态（前端渲染入口卡）。"""
    from assistant_brain import SCENES
    try:
        from ai_client import load_config as _ac_cfg
        cfg = _ac_cfg()
        ai_ready = str(cfg.get("driver") or "real").lower() != "fake" \
            and bool(cfg.get("api_key"))
    except Exception:
        ai_ready = False
    return jsonify({"ok": True, "scenes": SCENES,
                    "shops": sorted(competitor_lookup.keys()),
                    "ai_ready": ai_ready})


@app.route("/api/assistant/analyze", methods=["POST"])
def api_assistant_analyze():
    """跑一个核心分析场景。body={scene, shop}；scene 在免锁白名单（只读+长耗时）。"""
    data = request.get_json(force=True, silent=True) or {}
    scene = (data.get("scene") or "").strip()
    shop = (data.get("shop") or "").strip() or DEFAULT_SHOP
    mock = bool(data.get("mock"))
    if scene not in ("top", "cat", "price", "match", "plan"):
        return jsonify({"ok": False, "msg": f"未知场景: {scene}"}), 400
    if shop not in competitor_lookup:
        return jsonify({"ok": False, "msg": f"店铺不存在: {shop}"}), 400
    rows, cols, note = _asst_payload(scene, shop)
    if rows is None:
        return jsonify({"ok": False, "msg": note or "数据尚未就绪，稍后再试",
                        "scene": scene, "shop": shop}), 202
    if not rows:
        return jsonify({"ok": True, "scene": scene, "shop": shop, "mock": mock,
                        "result": {"summary": note, "insights": []}})
    from assistant_brain import analyze as _asst_analyze
    result = _asst_analyze(scene, {"rows": rows, "cols": cols,
                                   "scope": note, "shop": shop}, mock=mock)
    return jsonify({"ok": True, "scene": scene, "shop": shop, "mock": mock,
                    "result": result})


# ---------------------------------------------------------------------------
# 补货参考：导入向导 + 店↔仓指认（docs/任务规划表_20260904.md · T1）
# 原则：识别与预览不写库；写库必须"预览→确认"；对应关系永远人工优先（manual 不被覆盖）；
#      未指认的店绝不允许静默用多仓合计冒充本店销量。
# ---------------------------------------------------------------------------
REPL_DIR = os.path.join(BASE_DIR, "每日补货参考")
_repl_lock = threading.Lock()
_repl_run = {"state": "idle", "log": "", "started": "", "file": ""}


def _repl_snapshots(limit=12):
    """每日补货参考/ 下的候选快照（按修改时间倒序，最新在前）。"""
    try:
        names = [f for f in os.listdir(REPL_DIR)
                 if f.startswith("前置仓补货参考") and f.lower().endswith(".xlsx")
                 and not f.startswith("~$")]
    except OSError:
        names = []
    out = []
    for n in sorted(names, key=lambda f: -os.path.getmtime(os.path.join(REPL_DIR, f)))[:limit]:
        p = os.path.join(REPL_DIR, n)
        out.append({"name": n, "path": p,
                    "mtime": time.strftime("%Y-%m-%d %H:%M",
                                           time.localtime(os.path.getmtime(p))),
                    "size_kb": round(os.path.getsize(p) / 1024)})
    return out


def _repl_inspect(path):
    """读一份快照 → 它含哪几个仓、各多少行、当前指认到哪家店、系统推定建议是谁。

    读法必须走 bind_own_sales.read_snapshot（该表是二级表头，两处各写一份迟早会读歪）；
    识别结果按文件指纹落快照——同一份快照重复打开不必再解析 7MB 大表。
    """
    import own_store_lib
    from bind_own_sales import read_snapshot

    def _build():
        df = read_snapshot(path, cols=["门店/仓编码", "门店/仓名称", "商品SKU", "商品UPC"])
        return {"total_rows": int(len(df)),
                "buckets": [[_safe(code) or _safe(name), _safe(name), int(len(g))]
                            for (code, name), g in df.groupby(
                        ["门店/仓编码", "门店/仓名称"], dropna=False, sort=False)]}

    raw, _hit = _snap.remember("repl_inspect_" + os.path.basename(path), [path], _build)
    stores = own_store_lib.list_stores()
    mapped = _store_wh.load()
    rows = []
    for key, name, n in raw.get("buckets") or []:
        if not key:
            continue
        w = mapped.get(key) or {}
        sug = _store_wh.guess(stores, name, key)
        rows.append({"code": key, "wh_name": name, "rows": n,
                     "mapped_store": _safe(w.get("store_name")),
                     "source": _safe(w.get("source")),
                     "suggest_store": sug[0] if sug else "",
                     "suggest_code": sug[1] if sug else ""})
    return {"file": os.path.basename(path), "total_rows": raw.get("total_rows", 0),
            "warehouses": rows}


@app.route("/api/replenish/status")
def api_replenish_status():
    """向导首屏：可选快照 + 当前生效口径 + 每家自家店的指认状态。"""
    import own_store_lib
    try:
        with open(os.path.join(OUTPUT_DIR, "own_sales_meta.json"), encoding="utf-8") as f:
            meta = json.load(f)
    except Exception:
        meta = {}
    stores = []
    for s in own_store_lib.list_stores():
        st = _wh_status(s.get("name"))
        stores.append({"name": s.get("name"), "code": _safe(s.get("code")),
                       "state": st["state"], "usable": st["usable"],
                       "wh_name": st["wh_name"], "wh_code": st["code"]})
    return jsonify({"ok": True, "snap_dir": REPL_DIR, "snapshots": _repl_snapshots(),
                    "meta": meta, "stores": stores, "map": _store_wh.summary(),
                    "present_wh": sorted(_wh_sales().keys()), "run": _repl_run})


@app.route("/api/replenish/inspect", methods=["POST"])
def api_replenish_inspect():
    body = request.get_json(force=True, silent=True) or {}
    path = _safe(body.get("path"))
    if not path or not os.path.exists(path):
        return jsonify({"ok": False, "msg": "快照文件不存在，请重新选择"}), 400
    try:
        return jsonify(dict(ok=True, **_repl_inspect(path)))
    except Exception as e:
        return jsonify({"ok": False, "msg": f"读取失败：{e}"}), 500


@app.route("/api/replenish/map", methods=["POST"])
def api_replenish_map():
    """把一个仓指认给一家自家店（人工确认）；store_name 传空 = 取消指认。"""
    import own_store_lib
    body = request.get_json(force=True, silent=True) or {}
    code, wh_name = _safe(body.get("code")), _safe(body.get("wh_name"))
    store = _safe(body.get("store_name"))
    if not code:
        return jsonify({"ok": False, "msg": "缺少仓标识"}), 400
    names = {s.get("name") for s in own_store_lib.list_stores()}
    if store and store not in names:
        return jsonify({"ok": False, "msg": f"自家店清单里没有「{store}」"}), 400
    dup = [c for c, v in _store_wh.load().items()
           if store and v.get("store_name") == store and c != code]
    if dup:
        return jsonify({"ok": False,
                        "msg": f"「{store}」已指认给仓 {dup[0]}，一店一仓，请先取消那条"}), 400
    if store:
        rec = next((s for s in own_store_lib.list_stores() if s.get("name") == store), {})
        _store_wh.set_mapping(code, wh_name, store, _safe(rec.get("code")),
                              source="manual", note="导入向导人工指认")
    else:
        _store_wh.clear_mapping(code)
    return jsonify({"ok": True, "map": _store_wh.summary(),
                    "stores": [{"name": s.get("name"), **{"state": _wh_status(s.get("name"))["state"]}}
                               for s in own_store_lib.list_stores()]})


@app.route("/api/replenish/apply", methods=["POST"])
def api_replenish_apply():
    """后台子进程跑回填（dry_run=True 只做影响预览）。不阻塞界面，前端轮询 status 的 run 字段。"""
    body = request.get_json(force=True, silent=True) or {}
    path = _safe(body.get("path"))
    dry = bool(body.get("dry_run"))
    if not path or not os.path.exists(path):
        return jsonify({"ok": False, "msg": "快照文件不存在，请重新选择"}), 400
    try:
        unmapped = [f"{w['code']} {w['wh_name']}" for w in _repl_inspect(path)["warehouses"]
                    if not w["mapped_store"]]
    except Exception:
        unmapped = []
    if not dry and unmapped:
        return jsonify({"ok": False, "msg": "这些仓还没指认给自家店：" + "、".join(unmapped)
                        + "。指认完再执行，否则销量口径不完整。"}), 400
    with _repl_lock:
        if _repl_run.get("state") == "running":
            return jsonify({"ok": False, "msg": "已有回填任务在跑，请等它结束"}), 409
        _repl_run.update({"state": "running", "file": os.path.basename(path),
                          "started": time.strftime("%H:%M:%S"), "log": ""})

    def _run():
        import subprocess
        cmd = [sys.executable, os.path.join(SCRIPTS_DIR, "bind_own_sales.py"), "--file", path]
        if dry:
            cmd.append("--dry-run")
        env = dict(os.environ, PYTHONIOENCODING="utf-8")   # 否则子进程中文日志按 gbk 出来会变乱码
        try:
            r = subprocess.run(cmd, cwd=BASE_DIR, capture_output=True, timeout=1800,
                               encoding="utf-8", errors="replace", env=env)
            txt = ((r.stdout or "") + "\n" + (r.stderr or "")).strip()
            _repl_run.update({"state": "done" if r.returncode == 0 else "error",
                              "log": txt[-4000:], "rc": r.returncode})
        except Exception as e:
            _repl_run.update({"state": "error", "log": str(e), "rc": -1})
        if not dry:
            try:
                _reload_own_data()
                _invalidate_stats()
            except Exception as e:
                _repl_run["log"] += f"\n[面板重载失败，请重启面板] {e}"
    threading.Thread(target=_run, daemon=True).start()
    return jsonify({"ok": True, "started": True, "dry_run": dry, "unmapped": unmapped})


# ---------------------------------------------------------------------------
# T4 洞察台账：规则层负责发现（零成本、可复算），大模型只负责解释与排序
# ---------------------------------------------------------------------------
_insight_run = {"state": "idle", "log": "", "step": "", "done": 0, "total": 0}


def _insight_frames(shop):
    """取一店的四份结论（内存→磁盘快照→现算），预热与扫描共用一套取数。"""
    pdf, cdf = _derived("product", shop, lambda: __import__("build_product_analysis").build_product_analysis(shop),
                        code=_fp_derived())
    det, _agg = _derived("detail", shop, lambda: __import__("build_comparison").build_comparison_data(shop),
                         code=_fp_derived())
    key5 = (shop, "top500")
    if key5 not in _compare_cache:
        _derived("top500", shop, lambda: __import__("top50").build_top50(shop, 500), code=_fp_derived())
    return pdf, cdf, det, _compare_cache.get(key5) or []


@app.route("/api/insight/list")
def api_insight_list():
    import insight_store as IS
    try:
        d = IS.list_items(layer=(request.args.get("layer") or "").strip(),
                          status=(request.args.get("status") or "").strip(),
                          shop=(request.args.get("shop") or "").strip(),
                          present_only=request.args.get("gone") != "1")
    except Exception as e:
        return jsonify({"ok": False, "msg": f"读取台账失败：{e}"}), 500
    d.update({"ok": True, "run": _insight_run})
    return jsonify(d)


@app.route("/api/insight/refresh", methods=["POST"])
def _insight_scan_start():
    """启动洞察台账后台扫描（供 路由 与 聊天动作工具 共用）。
    返回 {"ok", "started", "shops"} 或 {"ok": False, "msg"}。"""
    if _insight_run.get("state") == "running":
        return {"ok": False, "msg": "已在扫描中，请稍候"}
    shops = list(competitor_lookup.keys())
    _insight_run.update({"state": "running", "step": "准备结论", "done": 0,
                         "total": len(shops), "log": ""})

    def _run():
        import insight_scan as _ISc
        import insight_store as _IS
        frames = {}
        for s in shops:
            _insight_run["step"] = f"计算/读取 {s[:10]} 的结论"
            try:
                frames[s] = _insight_frames(s)
            except Exception as e:
                _insight_run["log"] += f"\n[{s}] 取帧失败: {e}"
            _insight_run["done"] += 1
        _insight_run["step"] = "跑探测器"
        try:
            found = _ISc.scan(lambda sh: frames[sh], list(frames))
            stamp = _IS.caliber_stamp()
            # 请求线程的锁在返回时就释放了，写台账这一步必须由线程自己持锁
            with _data_lock_ctx(timeout=_LOCK_TIMEOUT):
                res = _IS.merge(found, stamp)
                _IS.record_scan(_ISc.coverage(), {"found": res["found"],
                                                   "added": res["added"],
                                                   "updated": res["updated"],
                                                   "total": res["total"]})
            _insight_run.update({"state": "done", "step": "", "found": res["found"],
                                 "added": res["added"], "updated": res["updated"],
                                 "total_all": res["total"], "caliber": stamp.get("text", ""),
                                 "coverage": _ISc.coverage()})
        except Exception as e:
            _insight_run.update({"state": "error",
                                 "log": _insight_run.get("log", "") + f"\n扫描失败: {e}"})
    threading.Thread(target=_run, daemon=True).start()
    return {"ok": True, "started": True, "shops": len(shops)}


def api_insight_refresh():
    """后台跑规则扫描：先逐店确保结论快照就位，再探测。全程不调大模型（0 元）。"""
    return jsonify(_insight_scan_start())


@app.route("/api/insight/state", methods=["POST"])
def api_insight_state():
    """标记处理状态——这是人的判断，扫描重跑也不得覆盖。
    同时写入采纳追踪台账（M6·3.3，append-only 流水）。"""
    import insight_store as IS
    import adopt_log as AL
    body = request.get_json(force=True, silent=True) or {}
    key, st = _safe(body.get("key")), _safe(body.get("status"))
    if not key:
        return jsonify({"ok": False, "msg": "缺少条目 key"}), 400
    ok = IS.set_status(key, st, _safe(body.get("note")))
    if ok:
        item = next((it for it in IS.load().get("items", []) if it.get("key") == key), {})
        act = {"adopted": AL.ACTION_ADOPT, "done": AL.ACTION_DONE,
               "ignored": AL.ACTION_IGNORE, "new": "撤销"}.get(st, st)
        stamp = IS.caliber_stamp()
        AL.record(AL.TYPE_INSIGHT, key,
                  obj={"subject": item.get("subject", ""), "detector": item.get("detector", "")},
                  action=act, note=_safe(body.get("note")),
                  data_date=stamp.get("data_date", ""), source="台账页",
                  shop=item.get("scope", ""))
    return jsonify({"ok": ok, "msg": "" if ok else "条目不存在或状态非法"}), (200 if ok else 400)


@app.route("/api/insight/deep", methods=["POST"])
def api_insight_deep():
    """每周深报（M6·3.2）：规则层产出当周结构化深报（0 元）；可选 AI 解读（需 key）。
    body: {weeks: 1, ai: false}。落盘 output/深报_<日期>.md。"""
    import weekly_report as WR
    body = request.get_json(force=True, silent=True) or {}
    weeks = int(body.get("weeks") or 1)
    try:
        res = WR.build(weeks=max(1, min(weeks, 4)))
    except Exception as e:
        return jsonify({"ok": False, "msg": f"深报生成失败：{e}"}), 500
    ai_text = ""
    if body.get("ai") is True:
        ai_text = WR.ai_interpret(res)
    md = WR.to_markdown(res, ai_text=ai_text)
    try:
        fp = os.path.join(OUTPUT_DIR, f"深报_{time.strftime('%Y%m%d')}.md")
        with open(fp, "w", encoding="utf-8") as f:
            f.write(md)
    except OSError:
        fp = ""
    return jsonify({"ok": True, "markdown": md, "structured": res, "file": fp,
                    "ai": bool(ai_text)})


@app.route("/api/insight/adopt")
def api_insight_adopt():
    """采纳追踪面板（M6·3.3）：计数 + 到期待回看清单（每条带数据可得性自述）。"""
    import adopt_log as AL
    st = AL.stats()
    pend = []
    for p in AL.pending_review():
        p["obs2"] = AL.has_shop_obs2(p.get("shop"))
        p["data_note"] = ("竞品已有第二期观测，可回看" if p["obs2"]
                          else "竞品仅 1 期观测——需要同店再导入一次才能回看效果")
        pend.append(p)
    return jsonify({"stats": st, "pending_review": pend[:30]})


def _md_to_html(md):
    """极简 markdown → HTML（够摘要用：标题/粗斜体/行内码/列表/分隔线）。

    服务端渲染而不是丢给浏览器：这一页要能直接打印、截图、发给老板，
    不依赖别的页面的 JS。规则与 analysis.html 的 renderMd 保持一致。
    """
    out = []
    for raw in str(md or "").split("\n"):
        t = (raw.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
        t = re.sub(r"\*\*(.+?)\*\*", r"<b>\1</b>", t)
        t = re.sub(r"(?<!\*)\*(?!\*)(.+?)\*", r"<i>\1</i>", t)
        t = re.sub(r"`(.+?)`", r"<code>\1</code>", t)
        s = t.strip()
        if s.startswith("### "):
            out.append("<h4>%s</h4>" % s[4:])
        elif s.startswith("## "):
            out.append("<h3>%s</h3>" % s[3:])
        elif s.startswith("# "):
            out.append("<h2>%s</h2>" % s[2:])
        elif s.startswith("- "):
            out.append('<div class="li">%s</div>' % s[2:])
        elif re.match(r"^\d+\. ", s):
            out.append('<div class="li">%s</div>' % s)
        elif set(s) == {"-"} and len(s) >= 3:
            out.append("<hr>")
        elif not s:
            out.append('<div class="gap"></div>')
        else:
            out.append("<p>%s</p>" % t)
    return "\n".join(out)


@app.route("/api/insight/brief")
def api_insight_brief():
    """给老板的一页摘要（标题下方就写数据截至与口径）。

    fmt=html 单页可直接打印/截图（默认，给链接用）；fmt=md 纯文本；fmt=json 给前端。
    以前只有 JSON，页面直链过去就是一屏 \\u 转义码。
    """
    import insight_store as IS
    try:
        top = max(3, min(30, int(request.args.get("top") or 12)))
    except ValueError:
        top = 12
    md = IS.export_brief(top=top)
    fmt = (request.args.get("fmt") or "html").strip().lower()
    if fmt == "md":
        return Response(md, mimetype="text/markdown; charset=utf-8")
    if fmt == "json":
        return jsonify({"ok": True, "markdown": md,
                        "caliber": IS.caliber_stamp().get("text", "")})
    return render_template("brief.html", body=_md_to_html(md), top=top,
                           at=time.strftime("%Y-%m-%d %H:%M"))


@app.route("/export/insight.xlsx")
def export_insight_xlsx():
    """给自己的：整本台账（含数字证据、算式、来源凭据、状态、连续期数）。"""
    import insight_store as IS
    path, _n = IS.export_xlsx()
    if not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name="AI洞察台账.xlsx")


@app.route("/api/insight/interpret", methods=["POST"])
def api_insight_interpret():
    """让大模型只做"解释 + 排序 + 把动作说具体"；数字一律沿用规则层给的。约 ¥0.03/次。"""
    import insight_store as IS
    body = request.get_json(force=True, silent=True) or {}
    items = [i for i in IS.list_items()["items"] if i.get("status") != "ignored"][:14]
    if not items:
        return jsonify({"ok": False, "msg": "台账为空，先跑一次扫描"}), 400
    from assistant_brain import interpret as _interp
    res = _interp(items, mock=bool(body.get("mock")))
    if not res.get("ok"):
        return jsonify({"ok": False, "msg": res.get("msg", "解读失败")}), 500
    return jsonify({"ok": True, "noted": IS.set_ai_note(items, res["notes"]),
                    "overall": res.get("overall", ""), "cost_est": res.get("cost_est", 0)})


def _insight_ctx(item, max_rows=12):
    """把一条洞察的 refs 展成真实数据行——追问的上下文只能是这些行，模型不许另加数字。"""
    rows = []
    for r in (item.get("refs") or []):
        ids = {str(x) for x in (r.get("ids") or [])}
        shop = r.get("shop") or (item.get("scope")
                                 if item.get("scope") in competitor_lookup else "")
        if not ids or shop not in competitor_lookup:
            continue
        try:
            pdf, cdf, det, p500 = _insight_frames(shop)
        except Exception as e:
            print("[追问] 取帧失败", shop, e)
            continue
        try:
            if r.get("kind") == "竞品SKU":
                got = det[det["竞品Sku_id"].astype(str).isin(ids)]
            elif r.get("kind") == "自家SKU":
                got = det[det["自家SKUID"].astype(str).isin(ids)]
            elif r.get("kind") == "品组":
                got = pdf[pdf["品组ID"].astype(str).isin(ids)]
            elif r.get("kind") == "类目":
                got = pdf[pdf["类目"].astype(str).isin(ids)]
            elif r.get("kind") == "竞品品":
                got = pd.DataFrame([p for p in (p500 or []) if str(p.get("comp_id")) in ids])
            else:
                got = pd.DataFrame()
            if hasattr(got, "to_dict"):
                rows += got.head(max_rows).to_dict("records")
        except Exception as e:
            print("[追问] 凭据取行失败", r.get("kind"), e)
        if len(rows) >= max_rows:
            break
    return rows[:max_rows]


@app.route("/api/assistant/home")
def api_assistant_home():
    """抽屉首屏：先给台账（0 元、可回溯），要解释再点场景或追问。"""
    import insight_store as IS
    import insight_scan as ISc
    only_shop = _safe(request.args.get("shop"))
    try:
        d = IS.list_items()
    except Exception as e:
        return jsonify({"ok": False, "msg": f"台账读取失败：{e}"}), 500
    live = [i for i in d["items"] if i.get("present", True)]
    if only_shop:
        # 只留与这家店有关的：本店 scope + 跨店共卖（全局层，本就跨多家店）
        live = [i for i in live
                if i.get("scope") in (only_shop, "跨店", "全部", "")]
    open_items = [i for i in live if i.get("status") == "new"]
    cov = ISc.coverage() or {}          # 本进程跑过的扫描才有；重启后是空
    cov_at, cov_stale = "", False
    if not cov:
        persisted = IS.scan_info()
        cov = persisted.get("coverage") or {}
        cov_at = persisted.get("at") or ""
        cov_stale = bool(cov_at)
    blocked = [{"detector": k, "why": (v.get("blocked") or [""])[0], "shops": v.get("shops", 0)}
               for k, v in cov.items()
               if not v.get("fired") and (v.get("blocked") or [""])[0]]
    obs = _shop_obs_meta()
    snaps = sorted({int(v.get("snaps") or 0) for v in obs.values()} or {0})
    return jsonify({
        "ok": True, "total": len(live), "open": len(open_items),
        "by_layer": {"全局": sum(1 for i in live if i.get("layer") == "全局"),
                     "细节": sum(1 for i in live if i.get("layer") == "细节")},
        "top3": [{"key": i["key"], "title": i["title"], "layer": i["layer"],
                  "scope": i["scope"], "detector": i["detector"],
                  "periods": i.get("periods", 1),
                  "action": i.get("ai_action") or i.get("hint", "")}
                 for i in sorted(open_items, key=lambda x: -float(x.get("severity") or 0))[:3]],
        "blocked": blocked, "caliber": d.get("caliber", ""),
        "scope_shop": only_shop,
        "coverage_at": cov_at, "coverage_stale": cov_stale,
        "data_date": d.get("data_date", ""), "updated": d.get("updated", ""),
        "obs_note": ("竞品观测期数：" + "/".join(str(s) for s in snaps) + " 次") if snaps else "尚无价格观测记录",
        "ledger": os.path.join(OUTPUT_DIR, "AI洞察.json")})


@app.route("/api/assistant/ask", methods=["POST"])
def api_assistant_ask():
    """就着某条洞察追问。上下文=该条凭据的真实行（不喂全表，模型不许造数）。"""
    import insight_store as IS
    body = request.get_json(force=True, silent=True) or {}
    key, q = _safe(body.get("key")), _safe(body.get("question"))
    item = next((i for i in IS.list_items()["items"] if i.get("key") == key), None)
    if not item:
        return jsonify({"ok": False, "msg": "这条已不在台账里（数据可能已变），请重新扫描"}), 400
    rows = _insight_ctx(item)
    from assistant_brain import ask as _ask
    res = _ask(item, rows, q, mock=bool(body.get("mock")))
    res.update({"key": key, "title": item.get("title"), "rows_used": len(rows)})
    return jsonify(res)


# ---------------------------------------------------------------------------
# 聊天式 AI 小助手（assistant_chat）：一个入口、能动手。
# 工具实现全部走面板当期缓存（0 成本、0 写库），数字只准来自工具返回——
# 模型只负责组织语言，不负责产生数字（红线见 assistant_chat.py 头注释）。
# ---------------------------------------------------------------------------
def _chat_ctx(shop):
    """构造聊天助手的工具上下文。返回 {"shop","shops","caliber","tools":{name:fn}}。"""
    import sales_text as _STc
    from search_util import kw_hit as _kwh
    import adopt_log as _ALm
    import insight_store as _ISm

    def _act_ledger(otype, key, obj, action, note=""):
        """动作留痕（M6·3.3 采纳台账）：确认/撤销/标记全部 append-only 落流水。"""
        stamp = _ISm.caliber_stamp()
        _ALm.record(otype, key, obj=obj, action=action, note=note,
                    data_date=stamp.get("data_date", ""), source="聊天助手", shop=shop)

    def _own_group_card(gid):
        """自家品组卡（确认对应预览用）：品名/类目/月销合计/规格数/代表SKU。"""
        gid_s = str(gid or "").strip()
        if not gid_s or own_df is None or not len(own_df):
            return None
        sub = own_df[own_df["SPUID"].astype(str).str.strip() == gid_s]
        if not len(sub):
            return None
        g = sub.drop_duplicates(subset=["SKUID"], keep="first").copy()
        g["__s"] = g["月销量"].apply(_extract_num).fillna(0.0)
        best = g.loc[g["__s"].idxmax()]
        return {"品组ID": gid_s,
                "品名": _safe(best.get("商品名称")) or _safe(best.get("清洗后商品名")),
                "类目": _safe(best.get("类目三级_归一")),
                "月销合计": round(float(g["__s"].sum()), 1),
                "规格数": int(len(g)),
                "代表SKUID": _safe(best.get("SKUID"))}

    def _real_spuids():
        return set(own_df["SPUID"].astype(str).str.strip()) \
            if own_df is not None and len(own_df) and "SPUID" in own_df.columns else set()

    def _top500(shop_):
        key5 = (shop_, "top500")
        if key5 not in _compare_cache:
            from top50 import build_top50
            _compare_cache[key5] = _derived("top500", shop_,
                                            lambda: build_top50(shop_, 500),
                                            code=_fp_derived())
        return list(_compare_cache.get(key5) or [])

    def _bridge(shop_):
        br = _bridge_cache.get(shop_)
        if br:
            return br
        if (shop_, "detail") not in _compare_cache or (shop_, "product") not in _compare_cache:
            threading.Thread(target=_warmup_one_shop, args=(shop_,), daemon=True).start()
        br = _build_bridge(shop_)
        if br[3]:
            _bridge_cache[shop_] = br
        return br

    def _pair_status(shop_, comp_id, skus, spu_own, sku2own, own_sku_map):
        """品级 对应状态 + 自家品名（单店口径，与首页榜单同链路）。"""
        conf_map = progress.get("confirmed", {}).get(shop_, {})
        nm_map = progress.get("no_match", {}).get(shop_, {})
        conf = any(conf_map.get(s) for s in skus)
        nm = any(nm_map.get(s) for s in skus)
        gid, gname = spu_own.get(str(comp_id) or "", ("", ""))
        if not gid or gid not in _real_spuids():
            for s in skus:
                own_sk = sku2own.get(s)
                if own_sk and own_sk in own_sku_map:
                    g2, n2 = own_sku_map[own_sk][0], own_sku_map[own_sk][1]
                    if g2:
                        gid, gname = gid or g2, (gname or n2)
                        break
        status = "已匹配" if conf else ("无对应" if nm else "待匹配")
        return gid, gname, status

    # ---- 工具 1：竞店品级销量榜（shop 可空 = 全部店合并榜，复用区榜并查集） ----
    def _t_rank_global(filt, n):
        """全部店合并榜：zone=__all__ 同款机器（绑定组>条码>品名归一 并查集合并），
        结论按 区分支同款简化口径（任一店已确认=已匹配；无对应且有效≥30=引进候选）。"""
        try:
            zrank = _zone_rank_cached({"name": "__all__", "comp_shops": list(competitor_lookup.keys()),
                                       "own_stores": [], "pseudo": True}, [])
        except Exception as e:
            return {"note": f"合并榜计算失败：{e}", "rows": [], "links": []}
        items = zrank.get("items") or []
        shops_involved = sorted({m.get("shop") for it in items
                                 for m in (it.get("店铺明细") or []) if m.get("shop")})
        bridges = {s: _bridge(s) for s in shops_involved if s in competitor_lookup}
        rows = []
        for it in items:
            gid, gname, matched = "", "", False
            for m in (it.get("店铺明细") or []):
                br = bridges.get(m.get("shop"))
                if br is None:
                    continue
                _osm, spu_own, _s2o, _cm = br
                if progress.get("confirmed", {}).get(m.get("shop"), {}).get(m.get("sku_id") or ""):
                    matched = True
                g2, n2 = spu_own.get(str(m.get("comp_id") or ""), ("", ""))
                if g2 and not gid:
                    gid, gname = g2, n2
            concl = ""
            if gid:
                for br in bridges.values():
                    if gid in br[3]:
                        concl = str(br[3][gid][0] or "")
                        break
            _ev = it.get("有效月售")
            ms = _STc.num(_ev) if _ev not in (None, "") else _STc.num(it.get("总月售"))
            if ms >= 30 and not gid:
                concl = "引进候选"
            if filt == "opp" and concl not in ("机会品", "引进候选", "上架候选"):
                continue
            if filt == "deflead" and concl not in ("防守品", "领先品"):
                continue
            shops_txt = "、".join(str(m.get("shop"))[:8]
                                  for m in (it.get("店铺明细") or []))[:60]
            _rep = next((m for m in (it.get("店铺明细") or []) if m.get("sku_id")), None)
            rows.append({"品名": it.get("品名") or "", "类目": it.get("类目") or "",
                         "品牌": it.get("品牌") or "",
                         "总月售": it.get("总月售") or 0, "有效月售": ms,
                         "在售店数": len(it.get("店铺明细") or []),
                         "在售店": shops_txt, "价格带": it.get("价格") or "",
                         "结论": concl or "观察品", "状态": "已匹配" if matched else "待匹配",
                         "自家品名": gname or "", "_cid": str(it.get("comp_id") or ""),
                         "_sku": str((_rep or {}).get("sku_id") or ""),
                         "_mshop": str((_rep or {}).get("shop") or "")
                                    or it.get("主店") or (shops_involved[0] if shops_involved else "")})
            if len(rows) >= n:
                break
        note = (f"全部店（{len(shops_involved)} 家）合并榜 {len(rows)} 行（filter={filt}；"
                "同品跨店自动合并：绑定组>条码>品名归一）")
        return {"note": note,
                "cols": ["品名", "总月售", "有效月售", "在售店数", "结论", "状态", "自家品名"],
                "rows": rows,
                "links": [{"label": "首页·全部竞店榜", "url": "/?zone=__all__" +
                           ("&filter=opp" if filt == "opp" else "")}]}

    def t_query_rank(a):
        filt = str(a.get("filter") or "hot")
        try:
            n = max(1, min(int(a.get("n") or 15), 30))
        except (TypeError, ValueError):
            n = 15
        shop_ = str(a.get("shop") or "").strip()   # 传空 = 全部店（语义见 schema）
        if not shop_ or shop_ in ("全部", "全部店", "全部竞店", "__all__"):
            return _t_rank_global(filt, n)
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}。可用店：{ '、'.join(sorted(competitor_lookup)[:8]) }（shop 传空=全部店合并）",
                    "rows": [], "links": []}
        prods = _top500(shop_)
        own_sku_map, spu_own, sku2own, concl_map = _bridge(shop_)
        rows = []
        for p in prods:
            skus = [str(s.get("sku_id") or "") for s in p.get("SKU列表", [])]
            gid, gname, status = _pair_status(shop_, p.get("comp_id"), skus,
                                              spu_own, sku2own, own_sku_map)
            concl = str(concl_map.get(gid, ("", ""))[0] or "") if gid else ""
            ev = p.get("有效销量")
            ms = _STc.num(ev) if ev not in (None, "") \
                else _STc.eff_of(p.get("月售"), p.get("意向"))
            if ms >= 30 and not gid:
                concl = "引进候选"
            if filt == "opp" and concl not in ("机会品", "引进候选", "上架候选"):
                continue
            if filt == "deflead" and concl not in ("防守品", "领先品"):
                continue
            rows.append({
                "品名": p.get("品名") or p.get("name") or "", "类目": p.get("类目") or "",
                "品牌": p.get("品牌") or "", "竞品价": p.get("价格") or "",
                "月售": p.get("月售") or 0, "有效销量": ms,
                "结论": concl or "观察品", "状态": status,
                "自家品名": gname or "", "自家月销": p.get("自家月销") if p.get("自家月销") != "" else "",
                "_sku": skus[0] if skus else ""})
            if len(rows) >= n:
                break
        note = f"{shop_} 榜单 {len(rows)} 行（filter={filt}，月售降序，当期缓存）"
        links = [{"label": f"首页榜单·{shop_[:10]}", "url": "/?shop=" + quote(shop_)},
                 {"label": "竞品总览", "url": "/browse?shop=" + quote(shop_)}]
        return {"note": note, "cols": ["品名", "竞品价", "月售", "有效销量",
                                       "结论", "状态", "自家品名", "自家月销"],
                "rows": rows, "links": links}

    # ---- 工具 2：按关键词查竞店的竞品 ----
    def t_query_product(a):
        kw = str(a.get("keyword") or "").strip()
        shop_ = str(a.get("shop") or shop or "").strip()
        if not kw:
            return {"note": "请给关键词（空格分隔多词 AND）", "rows": [], "links": []}
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}", "rows": [], "links": []}
        own_sku_map, spu_own, sku2own, _cm = _bridge(shop_)
        rows, seen = [], set()
        # 先查榜单缓存（有结论/状态，最快），不足再扫该店归一表
        for p in _top500(shop_):
            skus = [str(s.get("sku_id") or "") for s in p.get("SKU列表", [])]
            keys = " | ".join(x for x in (p.get("品名"), p.get("品牌"), p.get("类目")) if x)
            if not _kwh({"_": keys}, kw, ["_"]):
                continue
            gid, gname, status = _pair_status(shop_, p.get("comp_id"), skus,
                                              spu_own, sku2own, own_sku_map)
            concl = str(_cm.get(gid, ("", ""))[0] or "") if gid else ""
            rows.append({"品名": p.get("品名") or "", "价格": p.get("价格") or "",
                         "月售": p.get("月售") or 0,
                         "结论": concl or "观察品",
                         "状态": status, "自家品名": gname or "",
                         "comp_id": str(p.get("comp_id") or ""), "_sku": skus[0] if skus else ""})
            seen.add(str(p.get("comp_id") or ""))
            if len(rows) >= 8:
                break
        if len(rows) < 8:
            lk = competitor_lookup.get(shop_) or {}
            for sku, row in lk.items():
                cid = str(row.get("comp_id") or "")
                if cid in seen:
                    continue
                keys = " | ".join(str(row.get(c) or "") for c in
                                  ("name", "清洗后商品名", "原始商品名", "品牌（抽取）",
                                   "category3", "类目三级_归一"))
                if not _kwh({"_": keys}, kw, ["_"]):
                    continue
                skus = [s2 for s2, r2 in lk.items() if str(r2.get("comp_id") or "") == cid]
                gid, gname, status = _pair_status(shop_, cid, skus,
                                                  spu_own, sku2own, own_sku_map)
                concl = str(_cm.get(gid, ("", ""))[0] or "") if gid else ""
                rows.append({"品名": str(row.get("name") or ""), "价格": str(row.get("price") or ""),
                             "月售": str(row.get("sales") or ""), "结论": concl or "观察品",
                             "状态": status, "自家品名": gname or "",
                             "comp_id": cid, "_sku": sku})
                seen.add(cid)
                if len(rows) >= 8:
                    break
        if not rows:
            return {"note": f"{shop_} 没搜到含「{kw}」的竞品（多词 AND，全含才命中）",
                    "rows": [], "links": []}
        links = [{"label": (r["品名"] or "")[:14] or "详情",
                  "url": "/comp_detail?shop=" + quote(shop_) + "&sku=" + quote(r["_sku"])}
                 for r in rows[:6] if r.get("_sku")]
        return {"note": f"{shop_} 关键词「{kw}」命中 {len(rows)} 个品（前 8 行）",
                "cols": ["品名", "价格", "月售", "结论", "状态", "自家品名"],
                "rows": rows, "links": links}

    # ---- 工具 3：最近两次观测的变动（shop 可空 = 全部店汇总） ----
    def t_query_changes(a):
        from price_diff import diff_shop
        from collections import Counter
        shop_ = str(a.get("shop") or "").strip()   # 传空 = 全部店汇总
        if not shop_ or shop_ in ("全部", "全部店", "全部竞店"):
            summary, alarms = [], []
            for s in sorted(competitor_lookup.keys()):
                try:
                    rows, meta = diff_shop(s)
                except Exception:
                    continue
                if not rows:
                    summary.append({"店": s[:12], "观测期数": (meta or {}).get("snapshots", 1),
                                    "状态": "仅 1 次观测（再导入一次才能算变动）"})
                    continue
                cnt = Counter(str(r.get("变动") or "") for r in rows)
                summary.append({"店": s[:12], "观测期数": 2,
                                "变动": "、".join(f"{k}{v}" for k, v in cnt.most_common())})
                for r in rows:
                    if str(r.get("变动") or "") in ("undercut警报", "⚡undercut警报") and len(alarms) < 8:
                        alarms.append({"店": s[:10], "品名": str(r.get("品名") or r.get("商品名称") or "")[:24],
                                       "前值→现值": f"{r.get('前值')} → {r.get('现值')}",
                                       "月售": r.get("月销量") or r.get("月售") or ""})
            return {"note": f"全部 {len(summary)} 店变动汇总（各家最近两次观测）",
                    "cols": ["店", "观测期数", "变动", "状态"], "rows": summary + alarms,
                    "links": [{"label": "首页变动摘要", "url": "/"}]}
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}", "rows": [], "links": []}
        try:
            rows, meta = diff_shop(shop_)
        except Exception as e:
            return {"note": f"变动读取失败：{e}", "rows": [], "links": []}
        if not rows:
            return {"note": (meta or {}).get("note")
                    or f"{shop_} 只有 1 次观测——再导入一次同店数据才能算变动（数据没到，不是功能坏了）",
                    "rows": [], "links": []}
        out = [{"变动": r.get("变动"), "品名": r.get("品名") or r.get("商品名称") or "",
                "前值→现值": f"{r.get('前值')} → {r.get('现值')}",
                "月售": r.get("月销量") or r.get("月售") or ""}
               for r in rows[:10]]
        from collections import Counter
        cnt = Counter(str(r.get("变动") or "") for r in rows)
        note = (f"{shop_} 变动（{(meta or {}).get('prev_t', '?')} → {(meta or {}).get('cur_t', '?')}）："
                + "、".join(f"{k}{v}" for k, v in cnt.most_common()) + "；前 10 行如下")
        return {"note": note, "cols": ["变动", "品名", "前值→现值", "月售"],
                "rows": out,
                "links": [{"label": "首页变动摘要", "url": "/?shop=" + quote(shop_)}]}

    # ---- 工具 4：洞察台账 ----
    def t_query_insights(a):
        import insight_store as IS
        layer = str(a.get("layer") or "")
        det = str(a.get("detector") or "")
        shop_ = str(a.get("shop") or "").strip()
        try:
            n = max(1, min(int(a.get("n") or 10), 20))
        except (TypeError, ValueError):
            n = 10
        d = IS.list_items(layer=layer)
        items = [i for i in d["items"] if i.get("present", True)]
        if det:
            items = [i for i in items if det in str(i.get("detector") or "")]
        if shop_:
            items = [i for i in items if i.get("scope") in (shop_, "跨店", "全部", "")]
        items.sort(key=lambda x: -float(x.get("severity") or 0))
        out = []
        for i in items[:n]:
            met = {k: v for k, v in (i.get("metrics") or {}).items() if k != "算式"}
            out.append({"结论": i.get("title"), "层级": i.get("layer"), "范围": i.get("scope"),
                        "探测器": i.get("detector"), "连续未处理期数": i.get("periods", 1),
                        "数字证据": "；".join(f"{k}={v}" for k, v in list(met.items())[:6]),
                        "建议": str(i.get("ai_action") or i.get("hint") or "")[:80]})
        note = (f"台账未处理相关条目 {len(items)} 条，按严重度取前 {len(out)} 条"
                + ("（本店相关）" if shop_ else "（全店）"))
        return {"note": note, "cols": ["结论", "层级", "范围", "数字证据", "建议"],
                "rows": out,
                "links": [{"label": "看完整台账/凭据", "url": "/analysis"}]}

    # ---- 工具 5：自家总库查品 ----
    def t_query_own(a):
        kw = str(a.get("keyword") or "").strip()
        if not kw:
            return {"note": "请给关键词", "rows": [], "links": []}
        if own_df is None or not len(own_df):
            return {"note": "自家总库未加载", "rows": [], "links": []}
        name_col = "商品名称" if "商品名称" in own_df.columns else (
            "name" if "name" in own_df.columns else None)
        if not name_col:
            return {"note": "自家表缺商品名称列", "rows": [], "links": []}
        names = own_df[name_col].astype(str).tolist()
        m = [i for i, nm in enumerate(names) if _kwh({"_": nm}, kw, ["_"])]
        fallback = False
        if not m:
            # 语义兜底（M5·2.5）：关键词 0 命中 → 本地小模型向量 top8（未就绪静默跳过）
            fsub, _gids = _own_vec_sub(kw, 8)
            if fsub is not None and len(fsub):
                sub_fb = fsub.copy()
                sub_fb["_sk"] = sub_fb["SKUID"].astype(str).str.strip()
                sub_fb = sub_fb.drop_duplicates("_sk")
                sub_fb["_sp"] = sub_fb["SPUID"].astype(str).str.strip()
                sales_col_fb = "月销量" if "月销量" in sub_fb.columns else None
                price_col_fb = "美团闪购渠道售价" if "美团闪购渠道售价" in sub_fb.columns else None

                def _ffb(v):
                    x = _STc.num(v)
                    return float(x) if isinstance(x, (int, float)) else 0.0

                rows = []
                for sp, g in sub_fb.groupby("_sp"):
                    gs = g.sort_values(sales_col_fb, key=lambda c: c.map(_ffb),
                                       ascending=False) if sales_col_fb else g
                    r0 = gs.iloc[0]
                    ms = sum(_ffb(v) for v in g[sales_col_fb]) if sales_col_fb else 0.0
                    prices = [_ffb(v) for v in g[price_col_fb] if _ffb(v) > 0] \
                        if price_col_fb else []
                    rows.append({"自家品名": str(r0.get(name_col) or ""),
                                 "月销合计(30天)": round(ms, 1),
                                 "规格数": int(len(g)),
                                 "售价区间": (f"{min(prices):g}-{max(prices):g}" if prices else ""),
                                 "SPUID": sp})
                if rows:
                    return {"note": (f"自家总库没搜到「{kw}」的共有词，以下为本地小模型"
                                     f"语义近似兜底 {len(rows)} 个品组（可信度次一级，仅供参考）"),
                            "cols": ["自家品名", "月销合计(30天)", "规格数", "售价区间"],
                            "rows": rows,
                            "fallback": "语义",
                            "links": [{"label": "竞品库里搜同款", "url": "/browse?q=" + quote(kw)}]}
            return {"note": f"自家总库没搜到「{kw}」", "rows": [], "links": []}
        sub = own_df.iloc[m].copy()
        sub["_sk"] = sub["SKUID"].astype(str).str.strip()
        sub = sub.drop_duplicates("_sk")   # 渠道重复行去重（月销不得按行求和）
        sub["_sp"] = sub["SPUID"].astype(str).str.strip()
        sales_col = "月销量" if "月销量" in sub.columns else None
        price_col = "美团闪购渠道售价" if "美团闪购渠道售价" in sub.columns else None

        def _f(v):
            x = _STc.num(v)
            return float(x) if isinstance(x, (int, float)) else 0.0

        rows = []
        for sp, g in sub.groupby("_sp"):
            gs = g.sort_values(sales_col, key=lambda c: c.map(_f), ascending=False) \
                if sales_col else g
            r0 = gs.iloc[0]
            ms = sum(_f(v) for v in g[sales_col]) if sales_col else 0.0
            prices = [_f(v) for v in g[price_col] if _f(v) > 0] if price_col else []
            rows.append({"自家品名": str(r0.get(name_col) or ""),
                         "月销合计(30天)": round(ms, 1),
                         "规格数": int(len(g)),
                         "售价区间": (f"{min(prices):g}-{max(prices):g}" if prices else ""),
                         "SPUID": sp})
            if len(rows) >= 8:
                break
        rows.sort(key=lambda x: -x["月销合计(30天)"])
        return {"note": f"自家总库「{kw}」命中 {len(rows)} 个品组（月销为补货参考快照口径）",
                "cols": ["自家品名", "月销合计(30天)", "规格数", "售价区间"],
                "rows": rows,
                "links": [{"label": "竞品库里搜同款", "url": "/browse?q=" + quote(kw)}]}

    # ---- 工具 7：全局搜品（shop 可空 = 跨全部店聚合；这就是"全局搜索品"） ----
    def t_global_search(a):
        kw = str(a.get("keyword") or "").strip()
        if not kw:
            return {"note": "请给关键词（空格分隔多词 AND）", "rows": [], "links": []}
        shop_one = str(a.get("shop") or "").strip()
        if shop_one and shop_one not in ("全部", "全部店", "全部竞店"):
            return t_query_product({"shop": shop_one, "keyword": kw})
        try:
            from zone_lib import _norm_merge_name
        except Exception:
            def _norm_merge_name(s):
                return str(s or "").strip().lower()
        merged = {}
        for s in sorted(competitor_lookup.keys()):
            own_sku_map, spu_own, sku2own, cm = _bridge(s)
            hits = []
            for p in _top500(s):
                keys = " | ".join(x for x in (p.get("品名"), p.get("品牌"), p.get("类目")) if x)
                if _kwh({"_": keys}, kw, ["_"]):
                    hits.append((str(p.get("品名") or ""), p.get("价格") or "", p.get("月售") or 0,
                                 p.get("有效销量"), str(p.get("类目") or ""), str(p.get("品牌") or "")))
            if not hits:
                lk = competitor_lookup.get(s) or {}
                seen_c = set()
                for sku, row in lk.items():
                    cid = str(row.get("comp_id") or "")
                    if cid in seen_c:
                        continue
                    keys = " | ".join(str(row.get(c) or "") for c in
                                      ("name", "清洗后商品名", "品牌（抽取）", "类目三级_归一"))
                    if _kwh({"_": keys}, kw, ["_"]):
                        ev = _STc.eff_of(row.get("sales"), None)
                        hits.append((str(row.get("name") or ""), str(row.get("price") or ""),
                                     _STc.num(row.get("sales")) or 0, ev,
                                     str(row.get("类目三级_归一") or row.get("category3") or ""),
                                     str(row.get("品牌（抽取）") or "")))
                        seen_c.add(cid)
                        if len(hits) >= 4:
                            break
            gid, gname = "", ""
            # 该店对应关系：用关键词命中的代表 comp 的 spu 映射
            for sku, row in (competitor_lookup.get(s) or {}).items():
                cid = str(row.get("comp_id") or "")
                keys = " | ".join(str(row.get(c) or "") for c in
                                  ("name", "清洗后商品名", "品牌（抽取）", "类目三级_归一"))
                if _kwh({"_": keys}, kw, ["_"]):
                    gid, gname = spu_own.get(cid, ("", ""))
                    break
            for (nm, price, raw_sales, ev, cat, brand) in hits:
                key = _norm_merge_name(nm)[:40] or nm[:20]
                m = merged.setdefault(key, {"品名": nm, "在售店": [], "价格数值": [],
                                            "总月售": 0.0, "总有效月售": 0.0,
                                            "类目": cat, "品牌": brand, "自家品名": "",
                                            "_shop": s})
                m["在售店"].append(s[:10])
                m["总月售"] += float(raw_sales or 0)
                m["总有效月售"] += float(ev or 0)
                pv = _STc.num(price)
                if pv:
                    m["价格数值"].append(float(pv))
                if gid and not m["自家品名"]:
                    m["自家品名"] = gname
        if not merged:
            return {"note": f"全部店都没搜到含「{kw}」的竞品", "rows": [], "links": []}
        rows = []
        for m in merged.values():
            prices = sorted(m["价格数值"])
            band = f"{prices[0]:g}-{prices[-1]:g}" if prices else ""
            concl = "引进候选" if (not m["自家品名"] and m["总有效月售"] >= 30) else (
                "已对应" if m["自家品名"] else "观察品")
            rows.append({"品名": m["品名"][:30], "在售店数": len(m["在售店"]),
                         "在售店": "、".join(dict.fromkeys(m["在售店"]))[:50],
                         "总有效月售": round(m["总有效月售"], 1),
                         "跨店价格带": band, "结论": concl,
                         "自家品名": m["自家品名"] or "", "_sku_shop": m["_shop"]})
        rows.sort(key=lambda x: -x["总有效月售"])
        rows = rows[:8]
        links = [{"label": (r["品名"] or "")[:12] or "查看",
                  "url": "/browse?q=" + quote(kw)} for r in rows[:4]]
        return {"note": (f"全局搜「{kw}」：{len(merged)} 个品（跨店按品名归一聚合，"
                         "有效月售=成交+λ×收藏），按总有效月售取前 8"),
                "cols": ["品名", "在售店数", "总有效月售", "跨店价格带", "结论", "自家品名"],
                "rows": rows, "links": links}

    # ---- 工具 8：值得购入分析（底板=产品分析品卡，规则层已算好） ----
    def t_analyze_buy(a):
        shop_ = str(a.get("shop") or shop or "").strip() or DEFAULT_SHOP
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}", "rows": [], "links": []}
        got = _asst_cached(shop_, "product")
        if got is None:
            threading.Thread(target=_warmup_one_shop, args=(shop_,), daemon=True).start()
            return {"note": f"{shop_} 产品分析底板计算中（约30秒，后台预热后重问即可）",
                    "rows": [], "links": []}
        pdf = got[0]
        if pdf is None or not len(pdf):
            return {"note": f"{shop_} 无品卡底板（可能没有有效对应数据）", "rows": [], "links": []}
        mode = str(a.get("mode") or "机会类")
        allowed = {"机会类": ("机会品", "引进候选", "上架候选"),
                   "防守类": ("防守品",), "领先类": ("领先品",)}
        concls = allowed.get(mode)
        cat = str(a.get("category") or "").strip()
        sub = pdf
        if concls:
            sub = sub[sub["结论"].isin(list(concls))]
        if cat:
            sub = sub[sub["类目"].astype(str).str.contains(cat, regex=False)]
        if not len(sub):
            return {"note": f"{shop_} 在「{mode}」口径下没有候选（可试 mode=全部）",
                    "rows": [], "links": []}
        try:
            n = max(1, min(int(a.get("n") or 15), 30))
        except (TypeError, ValueError):
            n = 15
        sub = sub.sort_values("竞品有效月售", ascending=False).head(n)
        rows = [{"品名": str(r.get("品名") or "")[:30], "类目": str(r.get("类目") or ""),
                 "竞品品牌": str(r.get("竞品品牌") or ""),
                 "竞品价格带": str(r.get("竞品价格带") or ""),
                 "竞品有效月售": r.get("竞品有效月售"), "自家月销": r.get("自家月销"),
                 "自家规格数": r.get("自家规格数"), "竞品规格数": r.get("竞品规格数"),
                 "结论": str(r.get("结论") or ""),
                 "规则建议": str(r.get("建议") or "")[:90],
                 "品组ID": str(r.get("品组ID") or ""),
                 "_own_gid": str(r.get("品组ID") or "")}
                for _, r in sub.iterrows()]
        total = len(sub)
        return {"note": (f"{shop_} 购入分析底板：口径=「{mode}」结论类目"
                         + (f"、类目含「{cat}」" if cat else "")
                         + f"，共 {total} 个候选（按竞品有效月售降序取前 {len(rows)}）。"
                           "数字来自产品分析品卡（成交+λ×收藏），你做取舍与排序即可"),
                "cols": ["品名", "竞品有效月售", "竞品价格带", "自家月销", "结论", "规则建议"],
                "rows": rows,
                "links": [{"label": f"首页机会榜·{shop_[:8]}",
                           "url": "/?shop=" + quote(shop_) + "&filter=opp"},
                          {"label": "竞品总览", "url": "/browse?shop=" + quote(shop_)}]}

    # ---- 工具 9：店铺全景底板（统计+机会Top+类目缺口+变动，规则层拼装） ----
    def t_analyze_shop(a):
        shop_ = str(a.get("shop") or shop or "").strip() or DEFAULT_SHOP
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}", "rows": [], "links": []}
        st = _get_stats(shop_)
        rows = [{"项": "竞品总数", "值": st.get("total", 0)},
                {"项": "待确认", "值": st.get("pending", 0)},
                {"项": "已粗对齐(品)", "值": st.get("coarse", 0)},
                {"项": "已对应(品+规格)", "值": st.get("aligned", 0)},
                {"项": "无对应", "值": st.get("nomatch", 0)}]
        got = _asst_cached(shop_, "product")
        if got and len(got[0]):
            pdf = got[0]
            top = pdf.sort_values("竞品有效月售", ascending=False).head(5)
            for _, r in top.iterrows():
                rows.append({"项": f"机会Top｜{str(r.get('品名'))[:18]}",
                             "值": f"有效月售{r.get('竞品有效月售')}·{r.get('结论')}"
                                   f"·自家月销{r.get('自家月销')}"})
            cdf = got[1]
            if cdf is not None and len(cdf):
                ctop = cdf.sort_values("竞品有效月售", ascending=False).head(3)
                for _, r in ctop.iterrows():
                    rows.append({"项": f"类目缺口｜{str(r.get('类目'))[:14]}",
                                 "值": f"品数{r.get('品数')}·机会{r.get('机会品')}"
                                       f"·竞品有效月售{r.get('竞品有效月售')}"})
        try:
            from price_diff import diff_shop as _ds
            drows, dmeta = _ds(shop_)
            if drows:
                from collections import Counter as _Ct
                cnt = _Ct(str(r.get("变动") or "") for r in drows)
                rows.append({"项": "最近变动", "值": "、".join(
                    f"{k}{v}" for k, v in cnt.most_common())[:60]})
            else:
                rows.append({"项": "最近变动", "值": "仅1次观测，变动待下次导入"})
        except Exception:
            pass
        return {"note": f"{shop_} 全景底板（统计+机会Top+类目缺口+变动，数字全部本地规则层）",
                "cols": ["项", "值"], "rows": rows,
                "links": [{"label": f"首页·{shop_[:8]}", "url": "/?shop=" + quote(shop_)}]}

    # ---- 工具 10：经营分析底板（复用 ai_analysis.build_summary，规则层结构化数据） ----
    def t_analyze_business(a):
        shops_arg = str(a.get("shop") or shop or "").strip()
        if not shops_arg or shops_arg in ("全部", "全部店", "全部竞店"):
            shops_l = sorted(competitor_lookup.keys())[:5]
        elif shops_arg in competitor_lookup:
            shops_l = [shops_arg]
        else:
            return {"note": f"没有这家店：{shops_arg}", "rows": [], "links": []}
        dirs = [d for d in str(a.get("directions") or "").split(",") if d.strip()]
        from ai_analysis import build_summary as _bs
        try:
            _text, stt = _bs(shops_l, dirs or None)
        except Exception as e:
            return {"note": f"经营分析底板构建失败：{e}", "rows": [], "links": []}
        full = {}
        rows = []
        blocks = [("热销与引进", "hot"), ("缺口品", "gap"), ("类目缺口", "cat_gap"),
                  ("价格带", "band"), ("同款比价", "compare")]
        for label, key in blocks:
            data = stt.get(key) or []
            if isinstance(data, list) and data:
                full[label] = [{k: (str(v)[:60] if v is not None else "")
                                for k, v in r.items()}
                               for r in data[:60] if isinstance(r, dict)]
                for r in data[:8]:
                    if isinstance(r, dict):
                        rr = {"板块": label}
                        rr.update({k: (str(v)[:40]) for k, v in list(r.items())[:5]})
                        rows.append(rr)
        if not rows:
            return {"note": "经营分析底板为空（检查是否已导入数据）", "rows": [], "links": []}
        return {"note": (f"经营分析底板（店：{'、'.join(s[:8] for s in shops_l)}）——"
                         "数字全部来自本地规则层；你按板块做解读与排序，不新增数字；"
                         "完整版（全部板块全量行）在导出里"),
                "cols": ["板块", "品名/对象", "关键数字"],
                "rows": rows, "links": [{"label": "完整报告页", "url": "/analysis"}],
                "full": full}

    # ---- 工具 10b：前50逐品底板（复用 ai_top50 摘要层） ----
    def t_analyze_top50(a):
        shop_ = str(a.get("shop") or shop or "").strip() or DEFAULT_SHOP
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}", "rows": [], "links": []}
        prods50 = _top500(shop_)[:50]
        if not prods50:
            return {"note": f"{shop_} 榜单数据未就绪（稍后再问）", "rows": [], "links": []}
        from ai_top50 import build_summary as _t50s
        try:
            _summary, infos = _t50s(shop_, prods=prods50)
        except Exception as e:
            return {"note": f"前50底板构建失败：{e}", "rows": [], "links": []}
        infos = list(infos or [])
        if not infos:
            return {"note": f"{shop_} 前50底板为空", "rows": [], "links": []}
        from collections import Counter
        cnt = Counter(str(i.get("状态") or "") for i in infos)
        rows = sorted(infos, key=lambda x: -_STc.num(x.get("竞品月售")) or 0)[:15]
        rows = [{"序": i.get("id"), "品名": str(i.get("品名") or "")[:26],
                 "状态": i.get("状态"), "竞品月售": i.get("竞品月售"),
                 "销量口径": i.get("销量口径"), "自家品名": i.get("自家品名") or ""}
                for i in rows]
        return {"note": (f"{shop_} 前50逐品底板（已匹配 {cnt.get('已匹配', 0)}"
                         f"/无对应 {cnt.get('无对应', 0)}/待匹配 {cnt.get('待匹配', 0)}）——"
                         "你按状态给逐品建议：已匹配=优化，无对应=是否引进+参考价，待匹配=先去匹配"),
                "cols": ["品名", "状态", "竞品月售", "销量口径", "自家品名"],
                "rows": rows,
                "links": [{"label": f"首页·{shop_[:8]}",
                           "url": "/?shop=" + quote(shop_) + "&filter=opp"}],
                "full": {"前50逐品": infos}}

    # ---- 工具 10c：老板摘要底板（台账+口径戳，同 export_brief 的选条逻辑） ----
    def t_brief_boss(a):
        import insight_store as IS
        try:
            n = max(2, min(int(a.get("n") or 12), 20))
        except (TypeError, ValueError):
            n = 12
        d = IS.list_items()
        live = [i for i in d["items"] if i.get("status") != "ignored"]
        per = max(2, n // 2)
        items = ([i for i in live if i.get("layer") == "全局"][:per]
                 + [i for i in live if i.get("layer") == "细节"][:n - per])
        if not items:
            return {"note": "台账为空：本期规则层未发现问题（宁缺毋滥），老板摘要会如实说",
                    "rows": [], "links": []}
        rows = []
        for i in items:
            met = {k: v for k, v in (i.get("metrics") or {}).items()
                   if k != "算式"}
            rows.append({"层级": i.get("layer"), "结论": i.get("title"),
                         "范围": i.get("scope"),
                         "依据": "；".join(f"{k}={v}" for k, v in list(met.items())[:5]),
                         "动作": str(i.get("ai_action") or i.get("hint") or "")[:70],
                         "连续未处理期数": i.get("periods", 1)})
        return {"note": (f"老板摘要底板：全局 {per} 条 + 细节 {n - per} 条（按严重度），"
                         "口径照抄台账；你把标题/依据/动作转成给老板的一段话即可，不新增数字"),
                "cols": ["层级", "结论", "范围", "依据", "动作"],
                "rows": rows,
                "links": [{"label": "老板摘要页(可打印)", "url": "/api/insight/brief"}],
                "full": {"老板摘要": rows}}

    # ---- 工具 6：口径与新鲜度 ----
    def t_query_caliber(_a):
        import insight_store as IS
        from ai_client import log_stats
        try:
            d = IS.list_items()
            caliber = d.get("caliber", "")
            data_date = d.get("data_date", "")
        except Exception:
            caliber, data_date = "", ""
        obs = _shop_obs_meta()
        snaps = sorted({int(v.get("snaps") or 0) for v in obs.values()} or {0})
        st = log_stats()
        rows = [
            {"项": "竞品数据截至", "值": data_date or caliber or "（未知）"},
            {"项": "竞品观测期数", "值": "/".join(str(s) for s in snaps) + " 次" if snaps else "尚无"},
            {"项": "收藏折算权重λ", "值": _STc.INTENT_WEIGHT},
            {"项": "AI累计调用", "值": f"{st['calls']} 次 / 估算 ¥{st['cost_est']}"},
        ]
        return {"note": caliber or "口径未知",
                "cols": ["项", "值"], "rows": rows, "links": []}

    # ---- 动作类工具（M6·3.4，两段式：预览 → confirm=true 执行；全部留痕采纳台账） ----
    def t_mark_insight(a):
        key = _safe(a.get("key"))
        st = _safe(a.get("status")) or "采纳"
        note = _safe(a.get("note"))
        confirm = a.get("confirm") is True
        smap = {"采纳": "adopted", "已做": "done", "忽略": "ignored"}
        item = next((it for it in _ISm.load().get("items", []) if it.get("key") == key), None)
        if not item:
            return {"note": "台账没有该条目——先用 query_insights 拿到 key 再标记"}
        if st not in smap:
            return {"note": "status 只支持 采纳/已做/忽略"}
        if not confirm:
            return {"step": "preview",
                    "note": (f"待标记：{str(item.get('title') or '')[:60]} → 状态「{st}」。"
                             "用户明确同意后，再次调用并传 confirm=true。")}
        with _data_lock_ctx(timeout=_LOCK_TIMEOUT):
            ok = _ISm.set_status(key, smap[st], note)
        if not ok:
            return {"note": "标记失败（条目不存在）"}
        _act_ledger(_ALm.TYPE_INSIGHT, key,
                    {"subject": item.get("subject", ""), "detector": item.get("detector", "")},
                    st, note or "聊天标记")
        return {"note": f"已标记「{st}」✓（写回台账与采纳追踪流水）",
                "rows": [{"key": str(key)[:36], "状态": st}],
                "links": [{"label": "台账", "url": "/analysis"}]}

    def t_confirm_match(a):
        shop_a = _safe(a.get("shop")) or shop
        comp_id = _safe(a.get("comp_id"))
        gid = _safe(a.get("own_gid"))
        confirm = a.get("confirm") is True
        if not shop_a or not comp_id or not gid:
            return {"note": "需要 shop + comp_id（竞品品ID）+ own_gid（自家品组ID）"}
        if shop_a not in competitor_lookup:
            return {"note": f"店铺不存在：{shop_a}"}
        skus, r0 = _comp_skus_of(shop_a, comp_id)
        if not skus or r0 is None:
            return {"note": "未找到该竞品品"}
        own_card = _own_group_card(gid)
        if not own_card:
            return {"note": "own_gid 不存在（own_gid 必须来自 query_own / global_search 返回）"}
        skus = [s for s in skus if s]
        conf = [s for s in skus if progress.get("confirmed", {}).get(shop_a, {}).get(s)]
        if not confirm:
            note = (f"待确认（品级）：「{str(r0.get('商品名字') or '')[:26]}」"
                    f" ↔「{own_card['品名'][:26]}」｜竞品月售{str(r0.get('月销量') or '')}"
                    f" vs 自家月销{own_card['月销合计']}｜写入「人工确认·等同级」（可撤销）"
                    + ("｜⚠ 该品已有对应，请先撤销" if conf else ""))
            return {"step": "preview", "note": note}
        if conf:
            return {"note": "该品已有人工对应——先 undo_match 再确认新的"}
        from summary_library import upsert_manual, load_library as _ll2, save_library as _sl2, LEVEL_EQ
        rep = own_card.get("代表SKUID") or ""
        with _data_lock_ctx(timeout=_LOCK_TIMEOUT):
            spu_df, sku_df = _ll2()
            spu_df = upsert_manual(spu_df, "竞品品ID", shop_a, {
                "竞品品ID": comp_id,
                "竞品商品名": _safe(r0.get("商品名字")),
                "竞品品牌": _safe(r0.get("品牌（抽取）")),
                "竞品型号": _safe(r0.get("型号")),
                "竞品类目": _safe(r0.get("类目三级_归一")),
                "自家品组ID": gid, "自家品名": own_card["品名"],
                "置信度": 1.0, "对应级别": LEVEL_EQ, "来源": "人工"})
            save_library(spu_df, sku_df)
            cm = progress.setdefault("confirmed", {}).setdefault(shop_a, {})
            for sid in skus:
                cm[sid] = {"sku_id": rep, "spu_only": True,
                           "source": "人工确认(聊天)",
                           "time": __import__("time").strftime("%Y-%m-%d %H:%M:%S")}
            save_progress()
            _invalidate_stats()
        _act_ledger(_ALm.TYPE_ACTION, f"confirm:{shop_a}:{comp_id}",
                    {"竞品": str(r0.get("商品名字") or "")[:40], "自家品组": own_card["品名"],
                     "own_gid": gid, "skus": len(skus)},
                    "确认对应", "人工确认·等同级（聊天两段式）")
        return {"note": f"已确认对应 ✓（{len(skus)} 个规格记到品级；页面刷新可见，可撤销）",
                "rows": [{"竞品品ID": comp_id, "自家品组": gid}],
                "links": [{"label": "竞品库", "url": f"/browse?shop={shop_a}"}]}

    def t_undo_match(a):
        shop_a = _safe(a.get("shop")) or shop
        comp_id = _safe(a.get("comp_id"))
        sku = _safe(a.get("sku"))
        confirm = a.get("confirm") is True
        if not shop_a or (not comp_id and not sku):
            return {"note": "需要 shop + comp_id 或 sku"}
        if shop_a not in competitor_lookup:
            return {"note": f"店铺不存在：{shop_a}"}
        skus, _ = (_comp_skus_of(shop_a, comp_id) if comp_id else ([str(sku)], None))
        skus = [s for s in skus if s]
        if not skus:
            return {"note": "未找到该品/规格"}
        cm = progress.get("confirmed", {}).get(shop_a, {})
        have = {s: cm[s] for s in skus if s in cm}
        if not confirm:
            if not have:
                return {"note": "该品/规格当前没有人工确认记录，无需撤销"}
            own_names = []
            for s, rec in list(have.items())[:3]:
                osk = str(rec.get("sku_id") or "")
                nm = ""
                if osk and own_df is not None:
                    hit = own_df[own_df["SKUID"].astype(str) == osk]
                    nm = _safe(hit.iloc[0].get("商品名称")) if len(hit) else ""
                own_names.append(str(nm)[:24])
            return {"step": "preview",
                    "note": (f"待撤销：该品 {len(have)}/{len(skus)} 个规格的人工对应"
                             f"（{('；'.join(own_names)) or '（无自家名）'}）。"
                             "用户明确同意后传 confirm=true。")}
        with _data_lock_ctx(timeout=_LOCK_TIMEOUT):
            n_prog = 0
            for s in skus:
                if s in cm:
                    cm.pop(s, None)
                    n_prog += 1
            if n_prog:
                save_progress()
            from summary_library import load_library as _ll2, save_library as _sl2
            spu_df, sku_df = _ll2()
            if comp_id:
                m = ((spu_df["竞店名称"].astype(str) == shop_a)
                     & (spu_df["竞品品ID"].astype(str).str.strip() == comp_id))
                m_sku = ((sku_df["竞店名称"].astype(str) == shop_a)
                         & (sku_df["竞品品ID"].astype(str).str.strip() == comp_id))
            else:
                m = ((spu_df["竞店名称"].astype(str) == shop_a)
                     & (spu_df["竞品品ID"].astype(str).str.strip() == comp_id)
                     & False)   # 单SKU撤销不动 SPU 行（与页面撤销语义一致）
                m_sku = ((sku_df["竞店名称"].astype(str) == shop_a)
                         & (sku_df["竞品Sku_id"].astype(str).str.strip() == sku))
            n_lib = int(m.sum() + m_sku.sum())
            spu_df, sku_df = spu_df[~m], sku_df[~m_sku]
            _sl2(spu_df, sku_df)
            save_progress()
            _invalidate_stats()
        _act_ledger(_ALm.TYPE_ACTION, f"undo:{shop_a}:{comp_id or sku}",
                    {"skus": n_prog}, "撤销对应", f"聊天撤销（删库行 {n_lib} 条）")
        return {"note": f"已撤销 ✓（progress {n_prog} 条 / 汇总库 {n_lib} 行；已退回人工池）",
                "rows": [{"撤销规格": n_prog, "库行": n_lib}],
                "links": [{"label": "竞品库", "url": f"/browse?shop={shop_a}"}]}

    def t_rescan(a):
        confirm = a.get("confirm") is True
        info = _ISm.scan_info() or {}
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：洞察台账重新扫描（0 元，每店首次约半分钟，后台跑）。"
                             f"上次扫描 {info.get('at') or '从未跑过'}。"
                             "用户明确同意后传 confirm=true。")}
        res = _insight_scan_start()
        if not res.get("ok"):
            return {"note": res.get("msg") or "启动失败"}
        _act_ledger(_ALm.TYPE_ACTION, "rescan", {}, "登记", "聊天触发台账重扫")
        return {"note": "已启动后台扫描 ✓（0 元；台账页可看进度，完成后自动更新）",
                "links": [{"label": "台账", "url": "/analysis"}]}

    def t_deep_report(a):
        """每周深报一键化（M7 阶段1·用户点名）：预览（0 元）→ confirm=true 生成导出。"""
        import weekly_report as _WR
        weeks = a.get("weeks") or 1
        try:
            weeks = max(1, min(int(weeks or 1), 4))
        except (TypeError, ValueError):
            weeks = 1
        confirm = a.get("confirm") is True
        with_ai = a.get("ai") is True
        try:
            res = _WR.build(weeks=weeks)
        except Exception as e:
            return {"note": f"深报数据准备失败：{e}"}
        a2 = res["adopt"]
        rows = [{"项": f"近{weeks}周新增确认", "值": res["new_confirms"]},
                {"项": "采纳流水", "值": f"{a2['total']} 条（采纳 {a2['adopted']} / 已做 {a2['done']} / 忽略 {a2['ignored']}）"},
                {"项": "待回看（满30天）", "值": a2["pending_review"]},
                {"项": "当前未处理台账", "值": res["open_total"]}]
        if not confirm:
            return {"step": "preview",
                    "note": ("每周深报预览（规则层算数，0 元）。用户明确同意后传 confirm=true "
                             "生成并导出 markdown；要 AI 解读加 ai=true（约 ¥0.03）。"),
                    "cols": ["项", "值"], "rows": rows,
                    "links": [{"label": "台账页", "url": "/analysis"}]}
        ai_text = _WR.ai_interpret(res) if with_ai else ""
        md = _WR.to_markdown(res, ai_text=ai_text)
        fp = ""
        try:
            fp = os.path.join(OUTPUT_DIR, f"深报_{time.strftime('%Y%m%d')}.md")
            with open(fp, "w", encoding="utf-8") as _rf:
                _rf.write(md)
        except OSError:
            fp = ""
        _act_ledger(_ALm.TYPE_ACTION, "deep_report", {},
                    "登记", f"聊天触发每周深报（{weeks}周，AI={'是' if ai_text else '否'}）")
        return {"note": f"深报已生成并导出 ✓（{fp.split(chr(92))[-1] if fp else '仅预览'}）",
                "cols": ["项", "值"], "rows": rows,
                "links": [{"label": "台账页", "url": "/analysis"}]}

    # ---- M7 阶段2：查询面补全（全 risk=0，focus 契约） ----
    def t_query_slow(a):
        store = str(a.get("store") or "").strip()
        tier = str(a.get("tier") or "all").strip()
        try:
            n = max(1, min(int(a.get("n") or 15), 30))
        except (TypeError, ValueError):
            n = 15
        import slow_lib as _sl2
        import own_store_lib as _osl
        spu_set = None
        stores = _osl.list_stores()
        if store and any(s["name"] == store for s in stores):
            os_ = _osl.get_on_sale(store)
            if os_:
                spu_set = os_.get("spu_set")
        elif store:
            return {"note": f"没有这家自家门店：{store}。可用：{ '、'.join(s['name'][:10] for s in stores[:8]) or '（未导入门店表）' }",
                    "rows": [], "links": []}
        lib_sku_df = None
        try:
            _, lib_sku_df = _load_library_cached()
        except Exception:
            pass
        key_s = (store, id(spu_set))
        if key_s not in _slow_rank_cache:
            if len(_slow_rank_cache) > 8:
                _slow_rank_cache.clear()
            _slow_rank_cache[key_s] = _sl2.build_slow_rank(
                own_df, spu_set=spu_set, comp_lookup=competitor_lookup, lib_sku_df=lib_sku_df)
        res = _slow_rank_cache[key_s]
        items = res.get("items") or []
        if tier == "zero":
            items = [i for i in items if i.get("层级") == "零动销"]
        elif tier == "slow":
            items = [i for i in items if i.get("层级") == "慢动销"]
        rows = []
        for it in items[:n]:
            rows.append({
                "品名": str(it.get("品名") or "")[:28], "层级": it.get("层级"),
                "月销": it.get("月销"), "占压资金": it.get("占压"),
                "库存": it.get("库存"), "采购价": it.get("采购价"),
                "处置建议": str(it.get("处置") or it.get("建议") or "")[:40],
                "SPUID": str(it.get("SPUID") or ""), "_own_gid": str(it.get("SPUID") or "")})
        scope = store or "总库"
        return {"note": f"{scope} 滞销榜（占压资金降序，tier={tier}）共 {len(items)} 条，取前 {len(rows)}",
                "cols": ["品名", "层级", "月销", "占压资金", "处置建议"],
                "rows": rows,
                "links": [{"label": "滞销分析页", "url": "/slow" + (f"?own_store={quote(store)}" if store else "")}]}

    def t_query_pool(a):
        shop_ = str(a.get("shop") or "").strip()
        cat = str(a.get("category") or "").strip()
        try:
            n = max(1, min(int(a.get("n") or 15), 30))
        except (TypeError, ValueError):
            n = 15
        if pool_df is None or not len(pool_df):
            return {"note": "候选池是空的（可在首页/详情页把引进候选加入池）",
                    "rows": [], "links": [{"label": "候选池页", "url": "/pool"}]}
        df = pool_df
        if shop_ and "竞店名称" in df.columns:
            df = df[df["竞店名称"].astype(str) == shop_]
        if cat:
            cat_cols = [c for c in ("类目", "类目三级", "类目三级_归一", "category3") if c in df.columns]
            if cat_cols:
                m = False
                for c in cat_cols:
                    m = m | df[c].astype(str).str.contains(cat, regex=False)
                df = df[m]
        if not len(df):
            return {"note": f"候选池在该筛选下无条目（shop={shop_ or '全部'} cat={cat or '无'}）",
                    "rows": [], "links": [{"label": "候选池页", "url": "/pool"}]}
        rows = []
        for _, r in df.head(n).iterrows():
            sku = _safe(r.get("竞品Sku_id"))
            sh = _safe(r.get("竞店名称"))
            rows.append({
                "品名": _safe(r.get("竞品商品名") or r.get("商品名称") or r.get("品名"))[:28],
                "竞店": sh[:12], "竞品价": _safe(r.get("竞品价格") or r.get("到手价")),
                "月售": _safe(r.get("月销量") or r.get("月售")),
                "自家对应": _safe(r.get("自家商品名") or r.get("绑定自家品名")),
                "竞品Sku_id": sku, "_sku": sku, "_mshop": sh})
        return {"note": f"候选池 {len(df)} 条（shop={shop_ or '全部'}），取前 {len(rows)}",
                "cols": ["品名", "竞店", "竞品价", "月售", "自家对应"],
                "rows": rows,
                "links": [{"label": "候选池页", "url": "/pool"}]}

    def t_query_onlist(a):
        import own_store_lib as _osl
        store = str(a.get("store") or "").strip()
        kw = str(a.get("keyword") or "").strip()
        spuid = str(a.get("spuid") or "").strip()
        stores = _osl.list_stores()
        if not stores:
            return {"note": "尚未导入自家门店表（在「导入」页上传门店表后可用）",
                    "rows": [], "links": []}
        if spuid:
            rows = []
            for st in stores:
                os_ = _osl.get_on_sale(st["name"])
                on = bool(os_ and spuid in (os_.get("spu_set") or set()))
                rows.append({"门店": st["name"][:16], "SPUID": spuid,
                             "上架": "在售" if on else "未上架",
                             "门店在售数": st.get("on_sale_spus") or ""})
            on_n = sum(1 for r in rows if r["上架"] == "在售")
            return {"note": f"SPUID {spuid}：{on_n}/{len(rows)} 家门店在售",
                    "cols": ["门店", "上架", "门店在售数"], "rows": rows,
                    "links": [{"label": "滞销/门店分析", "url": "/slow"}]}
        if not store:
            rows = [{"门店": st["name"][:16], "在售SPU": st.get("on_sale_spus") or 0,
                     "文件时间": st.get("mtime") or st.get("updated") or ""}
                    for st in stores]
            return {"note": f"自家门店 {len(stores)} 家（在售数为该店导入表口径）",
                    "cols": ["门店", "在售SPU"], "rows": rows,
                    "links": [{"label": "滞销分析页", "url": "/slow"}]}
        if not any(s["name"] == store for s in stores):
            return {"note": f"没有这家门店：{store}。可用：{ '、'.join(s['name'][:10] for s in stores[:8]) }",
                    "rows": [], "links": []}
        os_ = _osl.get_on_sale(store)
        if not os_:
            return {"note": f"{store} 在售集合读取失败", "rows": [], "links": []}
        spu_set = os_.get("spu_set") or set()
        if own_df is None or not len(own_df):
            return {"note": "自家总库未加载", "rows": [], "links": []}
        sub = own_df[own_df["SPUID"].astype(str).str.strip().isin(spu_set)]
        if kw:
            from search_util import kw_mask as _km
            mask = _km(sub, kw, ["商品名称", "清洗后商品名", "原始商品名"])
            if mask is not None:
                sub = sub[mask]
        if not len(sub):
            return {"note": f"{store} 没有符合条件的在售品（kw={kw or '无'}）",
                    "rows": [], "links": []}
        # 按 SPU 聚合
        sub = sub.copy()
        sub["_sp"] = sub["SPUID"].astype(str).str.strip()
        rows = []
        for sp, g in list(sub.groupby("_sp"))[:15]:
            r0 = g.iloc[0]
            nm = _safe(r0.get("商品名称") or r0.get("清洗后商品名"))
            ms = 0.0
            if "月销量" in g.columns:
                try:
                    ms = float(g["月销量"].map(lambda v: _extract_num(v) or 0).sum())
                except Exception:
                    ms = 0.0
            rows.append({"品名": nm[:28], "SPUID": sp, "规格数": int(len(g)),
                         "月销合计": round(ms, 1), "_own_gid": sp})
        return {"note": f"{store} 在售 {len(spu_set)} SPU，展示前 {len(rows)}（kw={kw or '无'}）",
                "cols": ["品名", "SPUID", "规格数", "月销合计"], "rows": rows,
                "links": [{"label": "滞销分析", "url": "/slow?own_store=" + quote(store)}]}

    def t_query_history(a):
        shop_ = str(a.get("shop") or shop or "").strip()
        sku = str(a.get("sku") or "").strip()
        kw = str(a.get("keyword") or "").strip()
        # 复用面板已有的完整观测历史加载器
        try:
            hist = _load_hist_full()   # {店: {sku: [(t, price, sales, stock), ...]}}
        except Exception as e:
            return {"note": f"观测历史读取失败：{e}", "rows": [], "links": []}
        if not hist:
            return {"note": "尚无价格观测历史（导入同店第二次数据后会开始积累）",
                    "rows": [], "links": []}
        # 无店无sku：全店期数概览
        if not shop_ or shop_ in ("全部", "全部店", "全部竞店"):
            if not sku and not kw:
                rows = []
                for s, m in sorted(hist.items()):
                    periods = max((len(v) for v in m.values()), default=0)
                    rows.append({"店": s[:14], "观测SKU数": len(m), "期数上限": periods,
                                 "状态": "可算变动" if periods >= 2 else "仅1期，再导入一次才有变动"})
                return {"note": f"观测历史覆盖 {len(hist)} 店", "cols": ["店", "观测SKU数", "期数上限", "状态"],
                        "rows": rows, "links": [{"label": "首页", "url": "/"}]}
            shop_ = next(iter(hist))
        if shop_ not in hist:
            return {"note": f"没有 {shop_} 的观测历史。可用店：{ '、'.join(list(hist)[:6]) }",
                    "rows": [], "links": []}
        m = hist[shop_]
        # 定位 sku
        targets = []
        if sku and sku in m:
            targets = [sku]
        elif kw:
            lk = competitor_lookup.get(shop_) or {}
            for s, row in lk.items():
                keys = " | ".join(str(row.get(c) or "") for c in
                                  ("name", "清洗后商品名", "原始商品名"))
                if _kwh({"_": keys}, kw, ["_"]) and s in m:
                    targets.append(s)
                    if len(targets) >= 5:
                        break
        elif not sku:
            # 无定位：给有 ≥2 期的若干 SKU 概览
            multi = [s for s, pts in m.items() if len(pts) >= 2][:8]
            rows = []
            for s in multi:
                pts = m[s]
                p0, p1 = pts[0], pts[-1]
                rows.append({"Sku_id": s[:20],
                             "首期价→末期价": f"{p0[1]} → {p1[1]}",
                             "首期月售→末期": f"{p0[2]} → {p1[2]}",
                             "期数": len(pts)})
            return {"note": f"{shop_} 观测 {len(m)} 个SKU，其中 {sum(1 for v in m.values() if len(v)>=2)} 个有≥2期",
                    "cols": ["Sku_id", "首期价→末期价", "首期月售→末期", "期数"], "rows": rows,
                    "links": [{"label": "首页变动", "url": "/?shop=" + quote(shop_)}]}
        if not targets:
            return {"note": f"{shop_} 没找到对应观测轨迹（sku={sku or '无'} kw={kw or '无'}）",
                    "rows": [], "links": []}
        rows, focus = [], []
        for s in targets[:5]:
            pts = m[s]
            name = ""
            lk = competitor_lookup.get(shop_) or {}
            row = lk.get(s) or {}
            name = str(row.get("name") or "")[:24]
            for t, price, sales, stock in pts[-6:]:
                rows.append({"品名": name, "Sku_id": s[:18], "观测日": str(t)[:10],
                             "到手价": price, "月售": sales, "库存": stock,
                             "_sku": s, "_mshop": shop_})
            focus.append({"kind": "comp", "shop": shop_, "sku": s,
                          "name": name or s[:16],
                          "url": f"/comp_detail?shop={quote(shop_)}&sku={quote(s)}"})
        return {"note": f"{shop_} 轨迹（{len(targets)} 品，每品最多最近 6 期）",
                "cols": ["品名", "观测日", "到手价", "月售", "库存"], "rows": rows,
                "links": [{"label": "首页变动", "url": "/?shop=" + quote(shop_)}],
                "focus": focus}

    def t_query_health(a):
        only_bad = a.get("only_bad")
        if only_bad is None:
            only_bad = True
        try:
            res = health_summary()   # 复用面板 10 分钟缓存，不在聊天路径重扫 14MB
        except Exception as e:
            return {"note": f"体检执行失败：{e}", "rows": [], "links": []}
        if not res:
            return {"note": "体检模块无结果", "rows": [], "links": [{"label": "体检", "url": "/tools#health"}]}
        all_checks = res.get("items") or []
        checks = all_checks
        if only_bad:
            checks = [c for c in all_checks if str(c.get("level") or "") in ("bad", "warn")]
        rows = [{"项": c.get("item"), "级别": c.get("level"),
                 "说明": str(c.get("msg") or "")[:60],
                 "处置": str(c.get("fix") or "")[:40]} for c in checks[:20]]
        bad_n = res.get("bad", sum(1 for c in all_checks if c.get("level") == "bad"))
        warn_n = res.get("warn", sum(1 for c in all_checks if c.get("level") == "warn"))
        note = f"体检：红 {bad_n} / 黄 {warn_n} / 共 {res.get('total') or len(all_checks)} 项" \
               + ("，仅展示异常项" if only_bad else "")
        return {"note": note, "cols": ["项", "级别", "说明", "处置"], "rows": rows,
                "links": [{"label": "完整体检", "url": "/tools#health"}]}

    # ---- M7 阶段3：动作面 ----
    def t_query_learn(a):
        try:
            from learn import learn_status, load_syn_pending
            from learn import BRAND_FILL_FILE, BRAND_COLS, _read_df as _rd
            st = learn_status()
            pending_syn = [p for p in load_syn_pending() if p.get("status") == "待启用"]
            bdf = _rd(BRAND_FILL_FILE, BRAND_COLS)
            brand_pend = []
            if len(bdf):
                for _, r in bdf[bdf["状态"].astype(str) == "待启用"].iterrows():
                    brand_pend.append({"sku": _safe(r.get("竞品Sku_id")),
                                       "店": _safe(r.get("竞店名称")),
                                       "品牌": _safe(r.get("推断品牌")),
                                       "型号": _safe(r.get("推断型号"))})
        except Exception as e:
            return {"note": f"学习状态读取失败：{e}", "rows": [], "links": []}
        rows = [{"项": "待启用同义词", "值": len(pending_syn),
                 "示例": "；".join(f"{p.get('variant')}→{p.get('standard')}"
                                   for p in pending_syn[:5])},
                {"项": "待启用品牌补全", "值": len(brand_pend),
                 "示例": "；".join(f"{p.get('sku')}:{p.get('品牌')}"
                                   for p in brand_pend[:5])},
                {"项": "学习样本", "值": st.get("samples", 0), "示例": ""},
                {"项": "关键词库", "值": f"{st.get('keywords', 0)} 条 / {st.get('keyword_groups', 0)} 组",
                 "示例": ""}]
        # 明细（供 enable_* 预览）
        for p in pending_syn[:8]:
            rows.append({"项": "同义词待办", "值": p.get("variant"),
                         "示例": f"标准={p.get('standard')}",
                         "variant": p.get("variant"), "standard": p.get("standard")})
        for p in brand_pend[:8]:
            rows.append({"项": "品牌待办", "值": p.get("sku"),
                         "示例": f"{p.get('店','')[:8]} {p.get('品牌')} {p.get('型号') or ''}",
                         "sku": p.get("sku")})
        return {"note": f"学习待办：同义词 {len(pending_syn)} / 品牌 {len(brand_pend)}；"
                        f"样本 {st.get('samples', 0)}",
                "cols": ["项", "值", "示例"], "rows": rows,
                "links": [{"label": "工具页·学习", "url": "/tools"}]}

    def t_query_catmap(a):
        kw = str(a.get("keyword") or "").strip()
        try:
            n = max(1, min(int(a.get("n") or 15), 30))
        except (TypeError, ValueError):
            n = 15
        try:
            if not os.path.exists(CATEGORY_MAP_FILE):
                return {"note": "类目映射表不存在", "rows": [], "links": []}
            df = pd.read_excel(CATEGORY_MAP_FILE, dtype=str).fillna("")
        except Exception as e:
            return {"note": f"类目映射表读取失败：{e}", "rows": [], "links": []}
        if "三级类目" not in df.columns:
            return {"note": "映射表缺「三级类目」列", "rows": [], "links": []}
        if kw:
            df = df[df["三级类目"].astype(str).str.contains(kw, regex=False)]
        sug_col = "建议映射" if "建议映射" in df.columns else None
        rows = []
        pending = 0
        for _, r in df.head(n * 2).iterrows():
            cc = _safe(r.get("三级类目"))
            oc = _safe(r.get(sug_col)) if sug_col else ""
            if not cc:
                continue
            if not oc:
                pending += 1
            rows.append({"竞品类目": cc, "自家映射": oc or "（待确认）",
                         "竞品计数": _safe(r.get("竞品计数")), "自家计数": _safe(r.get("自家计数")),
                         "comp_cat": cc, "own_cat": oc})
            if len(rows) >= n:
                break
        total_pend = int((df[sug_col].astype(str).str.strip() == "").sum()) if sug_col else 0
        return {"note": f"类目映射共 {len(df)} 条，待确认约 {total_pend}（展示前 {len(rows)}）",
                "cols": ["竞品类目", "自家映射", "竞品计数", "自家计数"], "rows": rows,
                "links": [{"label": "竞品库·按类目", "url": "/browse"}]}

    def t_enable_synonym(a):
        variant = str(a.get("variant") or "").strip()
        standard = str(a.get("standard") or "").strip()
        enable = a.get("enable") is not False
        confirm = a.get("confirm") is True
        if not variant:
            return {"note": "缺少 variant（变体词）"}
        # 预览：从待启用列表补 standard
        try:
            from learn import load_syn_pending, enable_synonym as _en
            pend = load_syn_pending()
            rec = next((p for p in pend if p.get("variant") == variant), None)
            if not standard and rec:
                standard = str(rec.get("standard") or "")
        except Exception as e:
            return {"note": f"学习模块读取失败：{e}"}
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：{'启用' if enable else '忽略'}同义词 "
                             f"「{variant}」→「{standard or '?'}」。"
                             "启用会写入 synonyms.yaml（归一化生效）。用户明确同意后传 confirm=true。"),
                    "cols": ["变体", "标准", "动作"],
                    "rows": [{"变体": variant, "标准": standard or rec and rec.get("standard") or "",
                              "动作": "启用" if enable else "忽略"}]}
        try:
            _en(variant, standard, enable)
        except Exception as e:
            return {"note": f"写入失败：{e}"}
        _act_ledger(_ALm.TYPE_ACTION, f"syn:{variant}",
                    {"standard": standard}, "启用同义词" if enable else "忽略同义词",
                    "聊天助手")
        return {"note": f"已{'启用' if enable else '忽略'}同义词「{variant}」→「{standard}」 ✓",
                "rows": [{"变体": variant, "标准": standard, "动作": "已启用" if enable else "已忽略"}],
                "links": [{"label": "工具页·学习", "url": "/tools"}]}

    def t_enable_brand_fill(a):
        sku = str(a.get("sku") or "").strip()
        enable = a.get("enable") is not False
        confirm = a.get("confirm") is True
        if not sku:
            return {"note": "缺少 sku（竞品Sku_id）"}
        try:
            from learn import BRAND_FILL_FILE, BRAND_COLS, _read_df as _rd
            from learn import enable_brand_fill as _eb
            bdf = _rd(BRAND_FILL_FILE, BRAND_COLS)
            rec = bdf[bdf["竞品Sku_id"].astype(str) == sku] if len(bdf) else bdf
            if not len(rec):
                return {"note": f"品牌补全记录里没有 sku={sku}"}
            r0 = rec.iloc[0]
            brand, model, shop_b = _safe(r0.get("推断品牌")), _safe(r0.get("推断型号")), _safe(r0.get("竞店名称"))
        except Exception as e:
            return {"note": f"品牌补全读取失败：{e}"}
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：{'启用' if enable else '忽略'}品牌补全 "
                             f"sku={sku} → {brand} {model}（{shop_b}）。重跑对齐后生效。"
                             "用户明确同意后传 confirm=true。"),
                    "cols": ["Sku", "品牌", "型号", "店"],
                    "rows": [{"Sku": sku, "品牌": brand, "型号": model, "店": shop_b}]}
        try:
            _eb(sku, enable)
        except Exception as e:
            return {"note": f"写入失败：{e}"}
        _act_ledger(_ALm.TYPE_ACTION, f"brand:{sku}",
                    {"brand": brand}, "启用品牌补全" if enable else "忽略品牌补全",
                    "聊天助手")
        return {"note": f"已{'启用' if enable else '忽略'}品牌补全 {sku}→{brand} ✓（重跑对齐生效）",
                "rows": [{"Sku": sku, "品牌": brand, "动作": "已启用" if enable else "已忽略"}],
                "links": [{"label": "工具页", "url": "/tools"}]}

    def t_confirm_catmap(a):
        comp_cat = str(a.get("comp_cat") or "").strip()
        own_cat = str(a.get("own_cat") or "").strip()
        confirm = a.get("confirm") is True
        if not comp_cat or not own_cat:
            return {"note": "需要 comp_cat（竞品三级类目）和 own_cat（自家类目）"}
        cur = cat_map_cache.get(comp_cat, "")
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：类目映射「{comp_cat}」→「{own_cat}」"
                             + (f"（当前：{cur}）" if cur else "（当前无映射）")
                             + "。写映射表+缓存，收敛层级即时生效。用户明确同意后传 confirm=true。"),
                    "cols": ["竞品类目", "自家类目", "当前"],
                    "rows": [{"竞品类目": comp_cat, "自家类目": own_cat, "当前": cur or "（无）"}]}
        try:
            _confirm_category_map(comp_cat, own_cat)
        except Exception as e:
            return {"note": f"写入失败：{e}"}
        _act_ledger(_ALm.TYPE_ACTION, f"catmap:{comp_cat}",
                    {"own": own_cat}, "确认类目映射", "聊天助手")
        return {"note": f"已确认类目映射 ✓「{comp_cat}」→「{own_cat}」",
                "cols": ["竞品类目", "自家类目"],
                "rows": [{"竞品类目": comp_cat, "自家类目": own_cat}],
                "links": [{"label": "竞品库", "url": "/browse"}]}

    def t_start_prescreen(a):
        shop_ = str(a.get("shop") or shop or "").strip()
        status_only = a.get("status_only") is True
        confirm = a.get("confirm") is True
        try:
            limit = max(1, min(int(a.get("limit") or 100), 1000))
        except (TypeError, ValueError):
            limit = 100
        st = {k: _PRESCREEN_STATE.get(k) for k in
              ("running", "shop", "done", "total", "ok", "fail", "last_err", "last_report")}
        if status_only or (st.get("running") and not confirm):
            dist = {"是": 0, "否": 0, "拿不准": 0}
            total = 0
            if shop_:
                for (s, _sku), rec in _prescreen_map().items():
                    if s != shop_:
                        continue
                    total += 1
                    v = str((rec or {}).get("verdict") or "")
                    if v in dist:
                        dist[v] += 1
            note = (f"预筛进行中：{st.get('shop')} {st.get('done')}/{st.get('total')}"
                    if st.get("running") else "预筛未在跑")
            note += f"；{shop_} 已预筛 {total} 条（是{dist['是']}/否{dist['否']}/拿不准{dist['拿不准']}）"
            return {"note": note,
                    "cols": ["项", "值"],
                    "rows": [{"项": "运行中", "值": st.get("running")},
                             {"项": "进度", "值": f"{st.get('done')}/{st.get('total')}"},
                             {"项": "成功/失败", "值": f"{st.get('ok')}/{st.get('fail')}"},
                             {"项": f"{shop_} 已预筛", "值": total},
                             {"项": "结论分布", "值": f"是{dist['是']} 否{dist['否']} 拿不准{dist['拿不准']}"}],
                    "links": [{"label": "竞品库·预筛", "url": "/browse?shop=" + quote(shop_)}]}
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}"}
        # 待处理条数（预览用）
        pending_n = 0
        try:
            if pool_df is not None and len(pool_df):
                sub = pool_df[pool_df["竞店名称"].astype(str) == shop_]
                for idx in sub.index:
                    sku = _sku_of(idx)
                    if progress.get("confirmed", {}).get(shop_, {}).get(sku):
                        continue
                    if _is_no_match(idx):
                        continue
                    pending_n += 1
        except Exception:
            pending_n = -1
        if not confirm:
            # key 预检
            from ai_client import load_config as _lcfg
            cfg = _lcfg()
            has_key = str(cfg.get("driver") or "real").lower() == "fake" or bool((cfg.get("api_key") or "").strip())
            return {"step": "preview",
                    "note": (f"待执行：启动 {shop_} AI 批量预筛（最多 {limit} 条，后台跑，约 0.1 分/条）。"
                             f"待处理约 {pending_n} 条。"
                             + ("" if has_key else "⚠ 未配置 AI key，启动会失败。")
                             + "用户明确同意后传 confirm=true。"),
                    "cols": ["项", "值"],
                    "rows": [{"项": "店", "值": shop_}, {"项": "上限", "值": limit},
                             {"项": "待处理约", "值": pending_n},
                             {"项": "AI key", "值": "已配置" if has_key else "未配置"}]}
        if _PRESCREEN_STATE.get("running"):
            return {"note": f"预筛进行中（{_PRESCREEN_STATE.get('shop')} "
                            f"{_PRESCREEN_STATE.get('done')}/{_PRESCREEN_STATE.get('total')}），请稍候"}
        from ai_client import load_config as _lcfg2
        cfg2 = _lcfg2()
        if str(cfg2.get("driver") or "real").lower() != "fake" and not (cfg2.get("api_key") or "").strip():
            return {"note": "未配置 AI 接口密钥：请先在工具页「🔑 AI接口」配置 DeepSeek key"}
        try:
            threading.Thread(target=_run_prescreen_bg, args=(shop_, limit), daemon=True).start()
        except Exception as e:
            return {"note": f"启动失败：{e}"}
        _act_ledger(_ALm.TYPE_ACTION, f"prescreen:{shop_}",
                    {"limit": limit}, "启动AI预筛", "聊天助手")
        return {"note": f"已启动 {shop_} AI 预筛（最多 {limit} 条，后台）。稍后用 status_only=true 查进度。",
                "cols": ["项", "值"],
                "rows": [{"项": "店", "值": shop_}, {"项": "上限", "值": limit}],
                "links": [{"label": "竞品库", "url": "/browse?shop=" + quote(shop_)}]}

    # ---- M7 阶段4：后台任务面 ----
    def t_query_task_status(a):
        try:
            import backup_daily as _bd
            b = _bd.status()
        except Exception as e:
            b = {"last_ok": False, "last_error": str(e)}
        bg = _bg_snapshot()
        meta = {}
        try:
            mf = os.path.join(OUTPUT_DIR, "own_sales_meta.json")
            if os.path.exists(mf):
                with open(mf, encoding="utf-8") as _mf:
                    meta = json.load(_mf) or {}
        except Exception:
            meta = {}
        age = b.get("age_hours")
        rows = [
            {"项": "最近备份", "值": b.get("last_at") or "（无）",
             "说明": f"{b.get('last_files') or 0} 文件 {b.get('last_mb') or 0}MB · "
                     f"{'成功' if b.get('last_ok') else '失败'}"
                     + (f" · {age}小时前" if age is not None else "")},
            {"项": "备份份数", "值": f"{b.get('count') or 0} 份 / {b.get('total_mb') or 0}MB",
             "说明": b.get("last_error") or ""},
            {"项": "后台任务", "值": "进行中" if bg.get("running") else "空闲",
             "说明": (f"{bg.get('kind')} {bg.get('target')} · {bg.get('msg')}"
                      if bg.get("running") or bg.get("kind") else "")},
            {"项": "预筛任务", "值": "进行中" if (bg.get("prescreen") or {}).get("running") else "空闲",
             "说明": f"{(bg.get('prescreen') or {}).get('shop') or ''} "
                     f"{(bg.get('prescreen') or {}).get('done') or 0}/"
                     f"{(bg.get('prescreen') or {}).get('total') or 0}"},
            {"项": "销量口径", "值": (meta.get("note") or "未回填")[:40],
             "说明": f"快照 {meta.get('snapshot') or '-'} · {meta.get('bind_time') or '-'}"},
        ]
        note = (f"备份 {'正常' if b.get('last_ok') else '异常'}"
                f"（{age or '?'}h 前）；后台任务 "
                f"{'进行中: ' + str(bg.get('kind')) if bg.get('running') else '空闲'}")
        return {"note": note, "cols": ["项", "值", "说明"], "rows": rows,
                "links": [{"label": "工具页·体检", "url": "/tools#health"}]}

    def t_run_backup(a):
        confirm = a.get("confirm") is True
        bg = _bg_snapshot()
        if not confirm:
            import backup_daily as _bd
            try:
                b = _bd.status()
            except Exception:
                b = {}
            return {"step": "preview",
                    "note": (f"待执行：立即手动备份（zip+自检，后台约几十秒）。"
                             f"上次：{b.get('last_at') or '无'} "
                             f"{'成功' if b.get('last_ok') else '失败'}，"
                             f"现有 {b.get('count') or 0} 份。用户明确同意后传 confirm=true。"),
                    "cols": ["项", "值"],
                    "rows": [{"项": "上次备份", "值": b.get("last_at") or "无"},
                             {"项": "上次结果", "值": "成功" if b.get("last_ok") else "失败"},
                             {"项": "现有份数", "值": b.get("count") or 0}]}
        if bg.get("running"):
            return {"note": f"已有任务在跑：{bg.get('kind')}，请稍候"}
        ok, err = _bg_start("备份", "手动")
        if not ok:
            return {"note": err}
        threading.Thread(target=_run_backup_bg, daemon=True).start()
        _act_ledger(_ALm.TYPE_ACTION, "backup_manual", {}, "手动备份", "聊天助手")
        return {"note": "已启动手动备份（后台）。稍后用 query_task_status 查结果。",
                "links": [{"label": "工具页", "url": "/tools"}]}

    def t_rerun_align(a):
        shop_ = str(a.get("shop") or shop or "").strip()
        status_only = a.get("status_only") is True
        confirm = a.get("confirm") is True
        bg = _bg_snapshot()
        if status_only:
            return {"note": (f"后台任务：{'进行中 '+str(bg.get('kind'))+' '+str(bg.get('target'))}"
                             if bg.get("running") else "空闲")
                             + f"；最近结果：{bg.get('msg') or '（无）'}",
                    "cols": ["项", "值"],
                    "rows": [{"项": "运行中", "值": bg.get("running")},
                             {"项": "类型", "值": bg.get("kind")},
                             {"项": "目标", "值": bg.get("target")},
                             {"项": "消息", "值": bg.get("msg")},
                             {"项": "完成于", "值": bg.get("finished")}]}
        if shop_ not in competitor_lookup:
            return {"note": f"没有这家店：{shop_}。可用：{ '、'.join(list(competitor_lookup)[:6]) }"}
        std = os.path.join(OUTPUT_DIR, "竞店标准库", f"竞品_{shop_}_标准化.xlsx")
        has_std = os.path.exists(std)
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：重跑 {shop_} 对齐（归一化+对齐+汇总库，后台约1-2分钟）。"
                             f"标准库{'存在' if has_std else '缺失'}。"
                             "人工确认结果不会被覆盖。用户明确同意后传 confirm=true。"),
                    "cols": ["项", "值"],
                    "rows": [{"项": "店", "值": shop_},
                             {"项": "标准库", "值": "存在" if has_std else "缺失"},
                             {"项": "预计耗时", "值": "1-2 分钟"}]}
        if bg.get("running"):
            return {"note": f"已有任务在跑：{bg.get('kind')} {bg.get('target')}"}
        if not has_std:
            return {"note": f"{shop_} 标准库文件不存在，无法重跑"}
        ok, err = _bg_start("对齐重跑", shop_)
        if not ok:
            return {"note": err}
        threading.Thread(target=_run_align_bg, args=(shop_,), daemon=True).start()
        _act_ledger(_ALm.TYPE_ACTION, f"align:{shop_}", {}, "对齐重跑", "聊天助手")
        return {"note": f"已启动 {shop_} 对齐重跑（后台）。完成后面板自动重载；可用 status_only 查进度。",
                "links": [{"label": "竞品库", "url": "/browse?shop=" + quote(shop_)}]}

    def t_run_backfill_sales(a):
        status_only = a.get("status_only") is True
        confirm = a.get("confirm") is True
        bg = _bg_snapshot()
        if status_only:
            return t_query_task_status({})
        # 预览：上次口径 + 最新快照
        try:
            mf = os.path.join(OUTPUT_DIR, "own_sales_meta.json")
            meta = json.load(open(mf, encoding="utf-8")) if os.path.exists(mf) else {}
        except Exception:
            meta = {}
        root = os.path.dirname(OUTPUT_DIR)
        snap_dir = os.path.join(root, "每日补货参考")
        snaps = []
        try:
            if os.path.isdir(snap_dir):
                snaps = sorted([f for f in os.listdir(snap_dir)
                                if f.endswith(".xlsx") and "补货" in f], reverse=True)
        except OSError:
            snaps = []
        latest = snaps[0] if snaps else ""
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：销量回填（子进程写回自家表，约1分钟）。"
                             f"上次口径：{meta.get('note') or '未回填'}；"
                             f"最新快照：{latest or '未找到（检查 每日补货参考/）'}。"
                             "用户明确同意后传 confirm=true。"),
                    "cols": ["项", "值"],
                    "rows": [{"项": "上次口径", "值": meta.get("note") or "未回填"},
                             {"项": "上次快照", "值": meta.get("snapshot") or "-"},
                             {"项": "最新快照", "值": latest or "未找到"},
                             {"项": "回填时间", "值": meta.get("bind_time") or "-"}]}
        if bg.get("running"):
            return {"note": f"已有任务在跑：{bg.get('kind')}，请稍候"}
        if not latest:
            return {"note": "未找到补货参考快照（每日补货参考/ 下无 xlsx），无法回填"}
        ok, err = _bg_start("销量回填", latest)
        if not ok:
            return {"note": err}
        snap_path = os.path.join(snap_dir, latest)
        threading.Thread(target=_run_backfill_bg, args=(snap_path,), daemon=True).start()
        _act_ledger(_ALm.TYPE_ACTION, "backfill_sales",
                    {"snap": latest}, "销量回填", "聊天助手")
        return {"note": f"已启动销量回填（快照 {latest}，后台）。稍后用 query_task_status 查结果。",
                "links": [{"label": "工具页", "url": "/tools"}]}

    # ---- M7 阶段5：代理延展 ----
    def t_query_ai_model(a):
        try:
            from ai_client import load_config, log_stats
            cfg = load_config()
            st = log_stats()
        except Exception as e:
            return {"note": f"AI 配置读取失败：{e}", "rows": [], "links": []}
        has_key = bool((cfg.get("api_key") or "").strip()) or \
            str(cfg.get("driver") or "").lower() == "fake"
        rows = [
            {"项": "驱动", "值": cfg.get("driver") or "real",
             "说明": "fake=演示不花钱" if str(cfg.get("driver")) == "fake" else "真实接口"},
            {"项": "模型", "值": cfg.get("model") or "-", "说明": ""},
            {"项": "接口地址", "值": cfg.get("base_url") or "-", "说明": "OpenAI 兼容可换"},
            {"项": "Key", "值": "已配置" if (cfg.get("api_key") or "").strip() else "未配置",
             "说明": "改配置请到工具页「🔑 AI接口」（聊天禁止改配置）"},
            {"项": "累计调用", "值": st.get("calls", 0),
             "说明": f"成功 {st.get('ok', 0)} / 失败 {st.get('fail', 0)}"},
            {"项": "估算花费", "值": f"¥{st.get('cost_est', 0)}", "说明": "按牌价量级估算，非计费"},
        ]
        note = (f"当前 {cfg.get('driver')}/{cfg.get('model')} @ {cfg.get('base_url')}；"
                f"key {'已配' if (cfg.get('api_key') or '').strip() else '未配'}；"
                f"累计 {st.get('calls', 0)} 次 ≈¥{st.get('cost_est', 0)}")
        return {"note": note, "cols": ["项", "值", "说明"], "rows": rows,
                "links": [{"label": "工具页·AI接口", "url": "/tools"}]}

    def t_batch_propose(a):
        typ = str(a.get("type") or "全部").strip()
        try:
            n = max(1, min(int(a.get("n") or 20), 40))
        except (TypeError, ValueError):
            n = 20
        props = []
        try:
            from learn import load_syn_pending
            from learn import BRAND_FILL_FILE, BRAND_COLS, _read_df as _rd
            if typ in ("全部", "同义词"):
                for p in load_syn_pending():
                    if p.get("status") != "待启用":
                        continue
                    props.append({"id": f"syn:{p.get('variant')}",
                                  "类型": "同义词",
                                  "标题": f"{p.get('variant')} → {p.get('standard')}",
                                  "variant": p.get("variant"), "standard": p.get("standard")})
            if typ in ("全部", "品牌"):
                bdf = _rd(BRAND_FILL_FILE, BRAND_COLS)
                if len(bdf):
                    for _, r in bdf[bdf["状态"].astype(str) == "待启用"].iterrows():
                        sku = _safe(r.get("竞品Sku_id"))
                        props.append({"id": f"brand:{sku}",
                                      "类型": "品牌补全",
                                      "标题": f"{sku} → {_safe(r.get('推断品牌'))} "
                                              f"{_safe(r.get('推断型号')) or ''}".strip(),
                                      "sku": sku})
        except Exception as e:
            return {"note": f"提案读取失败：{e}", "rows": [], "links": []}
        if not props:
            return {"note": "没有待启用提案（可先跑学习提炼，或 type 选全部）",
                    "rows": [], "links": [{"label": "工具页", "url": "/tools"}]}
        rows = [{"id": p["id"], "类型": p["类型"], "标题": p["标题"][:40]} for p in props[:n]]
        return {"note": (f"批量提案 {len(props)} 条（展示前 {len(rows)}）。"
                         "选定后用 batch_apply，ids 传逗号分隔的 id。"),
                "cols": ["id", "类型", "标题"], "rows": rows,
                "full": {"提案": props[:80]},
                "links": [{"label": "工具页·学习", "url": "/tools"}]}

    def t_batch_apply(a):
        raw = a.get("ids") or ""
        if isinstance(raw, (list, tuple)):
            ids = [str(x).strip() for x in raw if str(x).strip()]
        else:
            ids = [s.strip() for s in str(raw).split(",") if s.strip()]
        confirm = a.get("confirm") is True
        if not ids:
            return {"note": "缺少 ids（来自 batch_propose）"}
        if len(ids) > 40:
            return {"note": "一次最多 40 条，请分批"}
        # 解析
        syns, brands = [], []
        for i in ids:
            if i.startswith("syn:"):
                syns.append(i[4:])
            elif i.startswith("brand:"):
                brands.append(i[6:])
            else:
                return {"note": f"无法识别的提案 id：{i}（应形如 syn:词 或 brand:sku）"}
        if not confirm:
            return {"step": "preview",
                    "note": (f"待执行：启用同义词 {len(syns)} 条 + 品牌补全 {len(brands)} 条。"
                             "用户明确同意后传 confirm=true。"),
                    "cols": ["id"],
                    "rows": [{"id": i} for i in ids[:30]]}
        ok_n, fail_n = 0, 0
        errs = []
        try:
            from learn import load_syn_pending, enable_synonym, enable_brand_fill
            pend = {p.get("variant"): p for p in load_syn_pending()}
            for v in syns:
                rec = pend.get(v) or {}
                try:
                    enable_synonym(v, str(rec.get("standard") or ""), True)
                    ok_n += 1
                except Exception as e:
                    fail_n += 1
                    errs.append(f"syn:{v}:{e}")
            for sku in brands:
                try:
                    enable_brand_fill(sku, True)
                    ok_n += 1
                except Exception as e:
                    fail_n += 1
                    errs.append(f"brand:{sku}:{e}")
        except Exception as e:
            return {"note": f"批量执行失败：{e}"}
        _act_ledger(_ALm.TYPE_ACTION, f"batch:{len(ids)}",
                    {"ok": ok_n, "fail": fail_n}, "批量启用提案", "聊天助手")
        return {"note": f"批量启用完成 ✓ 成功 {ok_n} / 失败 {fail_n}"
                        + (f"（{errs[0]}）" if errs else ""),
                "cols": ["项", "值"],
                "rows": [{"项": "成功", "值": ok_n}, {"项": "失败", "值": fail_n},
                         {"项": "同义词", "值": len(syns)}, {"项": "品牌", "值": len(brands)}],
                "links": [{"label": "工具页·学习", "url": "/tools"}]}

    def t_daily_brief(a):
        confirm = a.get("confirm") is True
        # 规则层数字（0 元）
        try:
            import insight_store as _IS2
            d_led = _IS2.list_items()
            items = [i for i in d_led.get("items") or [] if i.get("present", True)]
            open_n = len(items)
            top = sorted(items, key=lambda x: -float(x.get("severity") or 0))[:5]
        except Exception:
            open_n, top, d_led = 0, [], {}
        try:
            import backup_daily as _bd
            b = _bd.status()
        except Exception:
            b = {}
        bg = _bg_snapshot()
        try:
            import adopt_log as _AL2
            adopt = _AL2.stats() if hasattr(_AL2, "stats") else {}
        except Exception:
            adopt = {}
        rows = [
            {"项": "台账未处理", "值": open_n},
            {"项": "最近备份", "值": f"{b.get('last_at') or '无'} · "
                                    f"{'成功' if b.get('last_ok') else '失败'}"},
            {"项": "后台任务", "值": "进行中" if bg.get("running") else "空闲"},
            {"项": "数据日期", "值": (d_led.get("data_date") or "")[:10]},
        ]
        for t in top[:5]:
            rows.append({"项": f"Top｜{str(t.get('title') or '')[:20]}",
                         "值": f"严重度{t.get('severity')} · {t.get('scope') or ''}"})
        note = f"每日速报预览（规则层 0 元）：台账 {open_n} 条未处理。同意后 confirm=true 导出 markdown。"
        if not confirm:
            return {"step": "preview", "note": note, "cols": ["项", "值"], "rows": rows,
                    "links": [{"label": "台账", "url": "/analysis"}]}
        md_lines = [
            f"# 每日速报 {time.strftime('%Y-%m-%d %H:%M')}",
            "",
            f"- 台账未处理：{open_n}",
            f"- 最近备份：{b.get('last_at') or '无'}（{'成功' if b.get('last_ok') else '失败'}）",
            f"- 后台任务：{'进行中 '+str(bg.get('kind')) if bg.get('running') else '空闲'}",
            f"- 数据日期：{(d_led.get('data_date') or '')[:10] or '（未取到）'}",
            "",
            "## 严重度 Top",
        ]
        for t in top:
            md_lines.append(f"- [{t.get('severity')}] {t.get('title')}（{t.get('scope') or ''}）")
        md_lines += ["", "## 口径", str(d_led.get("caliber") or "")[:200], ""]
        fp = ""
        try:
            fp = os.path.join(OUTPUT_DIR, f"速报_{time.strftime('%Y%m%d')}.md")
            with open(fp, "w", encoding="utf-8") as _f:
                _f.write("\n".join(md_lines))
        except OSError:
            fp = ""
        _act_ledger(_ALm.TYPE_ACTION, "daily_brief", {"open": open_n}, "每日速报", "聊天助手")
        return {"note": f"每日速报已导出 ✓（{os.path.basename(fp) if fp else '仅内存'}）",
                "cols": ["项", "值"], "rows": rows,
                "links": [{"label": "台账", "url": "/analysis"}]}

    return {"shop": shop, "shops": sorted(competitor_lookup.keys()),
            "caliber": "", "tools": {
                "query_rank": t_query_rank, "query_product": t_query_product,
                "query_changes": t_query_changes, "query_insights": t_query_insights,
                "query_own": t_query_own, "query_caliber": t_query_caliber,
                "global_search": t_global_search,
                "analyze_buy": t_analyze_buy, "analyze_shop": t_analyze_shop,
                "analyze_business": t_analyze_business,
                "analyze_top50": t_analyze_top50, "brief_boss": t_brief_boss,
                "mark_insight": t_mark_insight, "confirm_match": t_confirm_match,
                "undo_match": t_undo_match, "rescan": t_rescan,
                "deep_report": t_deep_report,
                "query_slow": t_query_slow, "query_pool": t_query_pool,
                "query_onlist": t_query_onlist, "query_history": t_query_history,
                "query_health": t_query_health,
                "query_learn": t_query_learn, "query_catmap": t_query_catmap,
                "enable_synonym": t_enable_synonym, "enable_brand_fill": t_enable_brand_fill,
                "confirm_catmap": t_confirm_catmap, "start_prescreen": t_start_prescreen,
                "query_task_status": t_query_task_status, "run_backup": t_run_backup,
                "rerun_align": t_rerun_align, "run_backfill_sales": t_run_backfill_sales,
                "query_ai_model": t_query_ai_model, "batch_propose": t_batch_propose,
                "batch_apply": t_batch_apply, "daily_brief": t_daily_brief}}


@app.route("/api/assistant/chat", methods=["POST"])
def api_assistant_chat():
    """聊天助手入口。body={messages, shop, mock, sid}。
    只读 + 0 写库 → 免锁白名单；mock 模式真跑工具（真数字）不调模型（0 元）。
    sid 缺省时自动建会话；每轮回答后 把 user+assistant 两条记录落盘（含当时的
    范围与数据日期——回溯时能知道"这条答案是哪天的数据说的"）。"""
    body = request.get_json(force=True, silent=True) or {}
    msgs = body.get("messages") or []
    mock = bool(body.get("mock"))
    shop = (body.get("shop") or "").strip() or DEFAULT_SHOP
    if shop not in competitor_lookup:
        return jsonify({"ok": False, "msg": f"店铺不存在: {shop}"}), 400
    import chat_memory as CM
    sid = CM._safe_sid(body.get("sid"))
    if not sid:
        sid = CM.create_session(shop)["sid"]
    ctx = _chat_ctx(shop)
    import insight_store as IS
    try:
        d_led = IS.list_items()
        ctx["caliber"] = (d_led.get("caliber") or "")
        data_date = d_led.get("data_date", "")
    except Exception:
        ctx["caliber"], data_date = "", ""
    last_user = next((str(m.get("content") or "") for m in reversed(msgs)
                      if isinstance(m, dict) and m.get("role") == "user"), "")
    from assistant_chat import chat as _chatfn
    try:
        res = _chatfn(msgs, ctx, mock=mock)
    except RuntimeError as e:
        res = {"ok": False, "msg": str(e), "shop": shop}
    except Exception as e:
        return jsonify({"ok": False, "msg": f"助手异常：{e}", "shop": shop}), 500
    # 对话记忆落盘（失败不影响回答返回）
    try:
        CM.append(sid, {"role": "user", "text": last_user, "shop": shop,
                        "data_date": data_date})
        aidx = CM.append(sid, {"role": "assistant",
                               "text": res.get("answer") or res.get("msg") or "",
                               "trace": res.get("trace"), "links": res.get("links"),
                               "focus": res.get("focus"),
                               "shop": shop, "data_date": data_date})
        if aidx is not None:
            res["aidx"] = aidx
    except Exception:
        pass
    res["sid"] = sid
    if not res.get("ok"):
        return jsonify(res)
    return jsonify(res)


@app.route("/api/assistant/sessions", methods=["GET", "POST"])
def api_assistant_sessions():
    """会话管理：GET 列表；POST {action: create|rename|delete, sid, title, shop}。"""
    import chat_memory as CM
    if request.method == "GET":
        return jsonify({"ok": True, "sessions": CM.list_sessions()})
    body = request.get_json(force=True, silent=True) or {}
    act = str(body.get("action") or "")
    if act == "create":
        s = CM.create_session(_safe(body.get("shop")))
        return jsonify({"ok": True, **s})
    sid = CM._safe_sid(body.get("sid"))
    if not sid:
        return jsonify({"ok": False, "msg": "缺少 sid"}), 400
    if act == "rename":
        return jsonify({"ok": CM.rename_session(sid, body.get("title"))})
    if act == "delete":
        return jsonify({"ok": CM.delete_session(sid)})
    return jsonify({"ok": False, "msg": f"未知 action: {act}"}), 400


@app.route("/api/assistant/chat/history")
def api_assistant_history():
    """回溯：按 sid 取会话消息（append-only jsonl，含当时的 shop/data_date/trace/links）。"""
    import chat_memory as CM
    sid = CM._safe_sid(request.args.get("sid"))
    try:
        limit = max(1, min(int(request.args.get("limit") or 100), 200))
    except (TypeError, ValueError):
        limit = 100
    return jsonify({"ok": True, "sid": sid, "messages": CM.load_messages(sid, limit)})


@app.route("/api/assistant/export")
def api_assistant_export():
    """会话导出 Excel：sheet「对话记录」（时间/角色/内容/范围/数据日期）
    + 每个出现过的工具一张数据表（该会话所有工具返回的行，带 时间/范围/参数 列）。
    可选 idx：只导某一条（回答全文 + 该条引用的完整版工具表，供"⬇ 完整版"按钮）。"""
    import chat_memory as CM
    sid = CM._safe_sid(request.args.get("sid"))
    msgs = CM.load_messages(sid, 200)
    if not msgs:
        return jsonify({"ok": False, "msg": "会话不存在或还没有消息"}), 404
    try:
        idx = int(request.args.get("idx") or "")
    except (TypeError, ValueError):
        idx = None
    if idx is not None:
        if idx < 0 or idx >= len(msgs) or msgs[idx].get("role") != "assistant":
            return jsonify({"ok": False, "msg": "idx 不指向一条回答"}), 400
        keep = [msgs[idx - 1] if idx > 0 and msgs[idx - 1].get("role") == "user" else None,
                msgs[idx]]
        msgs = [m for m in keep if m]
    import io as _io
    import pandas as pd
    conv = [{"时间": m.get("t"), "角色": m.get("role"),
             "内容": (m.get("text") or "")[:1000], "范围": m.get("shop"),
             "数据日期": m.get("data_date")} for m in msgs]
    per_tool = {}
    sheet_seen = set()

    def _sheet_name(tool, i):
        base = str(tool or "tool")[:24]
        nm = base if i == 0 and base not in sheet_seen else f"{base}_{i}"
        sheet_seen.add(nm)
        return nm[:28]

    n_sheet = 0
    for m in msgs:
        if m.get("role") != "assistant":
            continue
        for t in (m.get("trace") or []):
            args_txt = ";".join(f"{k}={v}" for k, v in (t.get("args") or {}).items())[:80]
            full = t.get("full")
            groups = full.items() if isinstance(full, dict) and full \
                else [(None, t.get("data") or [])]
            for gname, data in groups:
                if not data:
                    continue
                rows = []
                for r in data[:200]:
                    if not isinstance(r, dict):
                        continue
                    row = {"时间": m.get("t"), "范围": m.get("shop"), "参数": args_txt}
                    for k, v in r.items():
                        row[str(k)[:30]] = v
                    rows.append(row)
                if rows:
                    per_tool[_sheet_name(t.get("tool"), n_sheet)] = rows
                    n_sheet += 1
    buf = _io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        pd.DataFrame(conv).to_excel(xw, sheet_name="对话记录", index=False)
        for name, rows in per_tool.items():
            df = pd.DataFrame(rows)
            df = df[[c for c in dict.fromkeys(
                ["时间", "范围", "参数"] + [c for c in df.columns
                                            if c not in ("时间", "范围", "参数")])
                if c in df.columns]]
            df.to_excel(xw, sheet_name=name, index=False)
    buf.seek(0)
    from flask import send_file
    tag = ("单条" if idx is not None else "会话")
    fname = f"助手{tag}导出_{sid[:12]}_{time.strftime('%Y%m%d_%H%M')}.xlsx"
    return send_file(buf, as_attachment=True, download_name=fname,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ---------------------------------------------------------------------------
# 每日体检（P4.4）：红黄绿灯 + 可照着做的处置
# health.run() 要扫 14MB 观测历史，所以结果缓存 10 分钟，首页横幅与接口共用一份。
# ---------------------------------------------------------------------------
_health_cache = {"at": 0.0, "data": None}
_health_lock = threading.Lock()


def _panel_listener_count(port=None):
    """本端口上的面板监听进程数（>1 = 有重复实例，员工机上等于多份内存）。"""
    try:
        import launch_panel as _lp
        pids = _lp.port_owner_pids(port or PANEL_PORT)
        return len([p for p in pids if "web_panel" in _lp.pid_cmdline(p)])
    except Exception:
        return 0


def health_summary(force=False):
    """体检结果（带缓存）。force=True 时立即重跑。"""
    import health as _hd
    now = time.time()
    if not force and _health_cache["data"] and now - _health_cache["at"] < 600:
        return _health_cache["data"]
    with _health_lock:
        if not force and _health_cache["data"] and time.time() - _health_cache["at"] < 600:
            return _health_cache["data"]
        try:
            r = _hd.run(shops=list(competitor_lookup.keys()))
        except Exception as e:
            r = {"items": [_hd._level("体检引擎", "bad", "跑不起来：" + str(e),
                                     "看 scripts/health.py 是否被改坏")],
                 "bad": 1, "warn": 0, "total": 1,
                 "at": time.strftime("%Y-%m-%d %H:%M:%S"), "verdict": "red"}
        n = _panel_listener_count()
        r["items"].append({
            "item": "面板实例", "level": "ok" if n <= 1 else "warn",
            "msg": ("只有 1 个面板进程在服务" if n <= 1 else
                    "端口上有 %d 个面板进程（重复实例，每个都持一份总库，员工机内存翻倍）" % n),
            "fix": "" if n <= 1 else "重跑 python scripts/launch_panel.py（会自动清残留并保留服务者）"})
        r["total"] = len(r["items"])
        _health_cache.update({"at": time.time(), "data": r})
    return r


@app.route("/api/health")
def api_health():
    force = request.args.get("refresh") == "1"
    r = health_summary(force=force)
    out = dict(r)
    out["ok"] = True
    return jsonify(out)


def _health_bar_for_page():
    """导航横幅用：只读缓存，绝不在渲染路径跑体检（要扫 14MB 观测历史）。
    缓存还没好时丢给后台线程跑一次，几秒后刷新页面就能看到灯标。"""
    d = _health_cache.get("data")
    if not d:
        if not _health_cache.get("pending"):
            _health_cache["pending"] = True
            threading.Thread(target=lambda: health_summary(), daemon=True).start()
        return None
    bad = [i for i in d["items"] if i["level"] == "bad"]
    warn = [i for i in d["items"] if i["level"] == "warn"]
    return {"verdict": d.get("verdict", "green"), "bad": len(bad), "warn": len(warn),
            "titles": [i["item"] for i in (bad or warn)][:4], "at": d.get("at", "")}


@app.context_processor
def _inject_health():
    return {"health_bar": _health_bar_for_page()}


def open_browser():
    time.sleep(1.2)
    webbrowser.open(f"http://127.0.0.1:{PANEL_PORT}")


def _warmup_readonly():
    """启动预热 = 只读盘、不现算（在开始服务之前同步调用）。

    把磁盘上已算好的派生结论读进内存（依赖指纹一致才认，约 0.1s/份）；没有有效
    快照的店留给首次访问懒加载。原因：本机 5 店 × 品卡 35s 的启动满算正是员工机上
    「点开就等 / 卡死 / 内存打满」的来源，而绝大多数时候盘上就有当期结论可用。
    导入新店后的现算由导入流程显式调 _warmup_one_shop，不混进启动路径。
    """
    fp = _fp_derived()
    n = 0
    for s in list(competitor_lookup.keys()):
        for kind in ("top500", "product", "detail"):
            if (s, kind) in _compare_cache:
                continue
            # 必须用 key_of()：与 remember 同一套指纹算法。自己拼 deps_fingerprint
            # 曾经导致预热命中 0 份、每次重启偷偷全量重算（表面看不出来）。
            hit = _snap.load("derived_%s_%s" % (kind, s),
                             _snap.key_of(_derived_deps(s, kind), fp))
            if hit is not None:
                _compare_cache[(s, kind)] = hit
                n += 1
    return n


def _warmup_one_shop(shop):
    """后台预热单店：(shop, "product") 与 (shop, "top500") 与 (shop, "detail") 缓存。
    导入新店后调用，避免首页首次访问同步重算。
    守卫：同店进行中不重复跑；构建失败 10 分钟退避——否则失败店的每次页面请求
    都会 spawn 新线程重跑重计算，线程堆积 + GIL 竞争把热请求从 0.2s 拖到 5s+。"""
    import time as _tt
    with _warmup_lock:
        st, ts = _warmup_state.get(shop, ("", 0))
        now = _tt.time()
        if st == "running":
            return
        if st == "failed" and now - ts < 600:
            return
        _warmup_state[shop] = ("running", now)
    failed = False
    from build_product_analysis import build_product_analysis
    try:
        key = (shop, "product")
        if key not in _compare_cache:
            _compare_cache[key] = _derived("product", shop,
                                           lambda: build_product_analysis(shop), code=_fp_derived())
    except Exception as e:
        failed = True
        print("[预热] 产品分析失败:", shop, e)
    try:
        key = (shop, "detail")
        if key not in _compare_cache:
            from build_comparison import build_comparison_data
            _compare_cache[key] = _derived("detail", shop,
                                           lambda: build_comparison_data(shop), code=_fp_derived())
    except Exception as e:
        failed = True
        print("[预热] 对齐明细失败:", shop, e)
    try:
        key = (shop, "top500")
        if key not in _compare_cache:
            from top50 import build_top50
            _compare_cache[key] = _derived("top500", shop,
                                           lambda: build_top50(shop, 500), code=_fp_derived())
    except Exception as e:
        failed = True
        print("[预热] top500 失败:", shop, e)
    # 桥接映射也预热（依赖上面 detail/product 缓存；失败不标记 failed——
    # 桥接可由请求路径惰性重建，只影响该店首个请求 ~0.9s，不阻塞页面）
    try:
        if (shop, "detail") in _compare_cache and shop not in _bridge_cache:
            _bridge_cache[shop] = _build_bridge(shop)
    except Exception as e:
        print("[预热] 桥接映射失败:", shop, e)
    with _warmup_lock:
        _warmup_state[shop] = ("failed" if failed else "ok", _tt.time())


def _auto_apply_keywords():
    """导入新店后自动跑关键词反哺（机器学习闭环），仅在用户未手动跑过 --apply 时执行。
    通过 mtime 检测，若脚本修改过则跳过（用户应主动跑）。"""
    try:
        import os, time as _t, subprocess as _sp
        script = os.path.join(SCRIPTS_DIR, "extract_search_keywords.py")
        kw_file = os.path.join(SCRIPTS_DIR, "progress.json")  # 仅作为锁文件信号
        # 仅执行一次：上次执行时间写入 progress.json（_kw_apply_ts）
        try:
            if os.path.exists(PROGRESS_FILE):
                with open(PROGRESS_FILE, encoding="utf-8") as _f:
                    _prog = json.load(_f)
        except Exception:
            _prog = {}
        last_ts = _prog.get("_kw_apply_ts", 0)
        if _t.time() - last_ts < 3600:  # 1h 内不重复跑
            return
        r = _sp.run([sys.executable, script, "--apply"],
                    cwd=BASE_DIR, capture_output=True, timeout=120,
                    encoding="utf-8", errors="replace")
        _prog["_kw_apply_ts"] = _t.time()
        _prog["_kw_apply_last"] = (r.stdout or "")[-200:]
        try:
            save_progress()
        except Exception:
            pass
    except Exception as e:
        print("[关键词反哺] 失败（无害）:", e)


# ---------------------------------------------------------------------------
# ③ 候选池：引进候选从展示变交互（入池/绑定组/跨店绑定推荐/LLM核实）
# ---------------------------------------------------------------------------
import candidate_pool as _cp


@app.context_processor
def _inject_pool_total():
    """导航栏候选池角标（池条目数；文件小，每次渲染读一次开销可忽略）。"""
    try:
        return {"pool_total": len(_cp.load_pool()["entries"]),
                "show_assistant": True}
    except Exception:
        return {"pool_total": 0, "show_assistant": True}


@app.route("/pool")
def pool_page():
    """候选池页：总库 → 店/区分组视图 + 绑定组管理 + 跨店同品绑定。"""
    data = _cp.pool_overview()
    # 补算折扣字段（候选池条目已含原价，在此统一算折扣徽章）
    for e in data["entries"]:
        e["折扣"] = _discount_badge(e.get("价格"), e.get("原价"))
    for g in data["groups"]:
        for m in g["members"]:
            m["折扣"] = _discount_badge(m.get("价格"), m.get("原价"))
    for p in data.get("pending") or []:
        for side in ("a", "b"):
            s = p.get(side, {})
            s["折扣"] = _discount_badge(s.get("价格"), s.get("原价"))
    shops = sorted({e.get("shop", "") for e in data["entries"] if e.get("shop")})
    return render_template("pool.html", active="pool", data=data, shops=shops)


def _pool_prod(shop, comp_id):
    """从 top500 缓存取品级 dict（无缓存则现场构建）。"""
    key5 = (shop, "top500")
    if key5 not in _compare_cache:
        from top50 import build_top50
        _compare_cache[key5] = _derived("top500", shop, lambda: build_top50(shop, 500),
                                        code=_fp_derived())
    for p in (_compare_cache[key5] or []):
        if str(p.get("comp_id")) == str(comp_id):
            return p
    return None


@app.route("/api/pool/add", methods=["POST"])
def api_pool_add():
    data = request.get_json(force=True)
    shop = (data.get("shop") or "").strip()
    comp_id = (data.get("comp_id") or "").strip()
    source = (data.get("source") or "").strip() or "手动入池"
    if not shop or not comp_id:
        return jsonify({"ok": False, "msg": "缺少 店铺/品ID"}), 400
    if shop not in competitor_lookup:
        return jsonify({"ok": False, "msg": f"店铺不存在: {shop}"}), 400
    prod = _pool_prod(shop, comp_id)
    if prod is None:
        return jsonify({"ok": False, "msg": "竞品品不存在（可能未入榜单前500，请手动添加）"}), 404
    entry, created = _cp.add_entry(shop, prod, source=source)
    sug_n = sum(1 for s in (entry.get("推荐") or []) if float(s.get("score") or 0) > 0)
    return jsonify({"ok": True, "created": created, "entry": entry, "suggest_n": sug_n,
                    "msg": "已入池" if created else "已在池中（未重复添加）"})


@app.route("/api/pool/remove", methods=["POST"])
def api_pool_remove():
    data = request.get_json(force=True)
    eid = (data.get("eid") or "").strip()
    if not eid:
        return jsonify({"ok": False, "msg": "缺少 eid"}), 400
    ok = _cp.remove_entry(eid)
    return jsonify({"ok": ok, "msg": "已移除" if ok else "条目不存在"})


@app.route("/api/pool/group/create", methods=["POST"])
def api_pool_group_create():
    data = request.get_json(force=True)
    eids = data.get("eids") or []
    name = (data.get("name") or "").strip()
    if len(eids) < 2:
        return jsonify({"ok": False, "msg": "绑定组至少需要 2 个条目（跨店同品才有意义）"}), 400
    g = _cp.create_group(eids, name)
    if g is None:
        return jsonify({"ok": False, "msg": "条目不存在"}), 404
    return jsonify({"ok": True, "group": g})


@app.route("/api/pool/group/add", methods=["POST"])
def api_pool_group_add():
    data = request.get_json(force=True)
    g = _cp.group_add((data.get("gid") or "").strip(), (data.get("eid") or "").strip())
    if g is None:
        return jsonify({"ok": False, "msg": "组或条目不存在"}), 404
    return jsonify({"ok": True, "group": g})


@app.route("/api/pool/group/remove", methods=["POST"])
def api_pool_group_remove():
    data = request.get_json(force=True)
    ok = _cp.group_remove((data.get("gid") or "").strip(), (data.get("eid") or "").strip())
    return jsonify({"ok": ok})


@app.route("/api/pool/group/rename", methods=["POST"])
def api_pool_group_rename():
    data = request.get_json(force=True)
    g = _cp.group_rename((data.get("gid") or "").strip(), (data.get("name") or "").strip())
    if g is None:
        return jsonify({"ok": False, "msg": "组不存在"}), 404
    return jsonify({"ok": True, "group": g})


@app.route("/api/pool/group/delete", methods=["POST"])
def api_pool_group_delete():
    data = request.get_json(force=True)
    ok = _cp.group_delete((data.get("gid") or "").strip())
    return jsonify({"ok": ok, "msg": "已解散（条目回到未绑定）" if ok else "组不存在"})


@app.route("/api/pool/group/link_own", methods=["POST"])
def api_pool_group_link_own():
    """绑定组关联自家品组（辅助商品对应/全面分析）。own_gid 空=解除关联。"""
    data = request.get_json(force=True)
    gid = (data.get("gid") or "").strip()
    own_gid = (data.get("own_gid") or "").strip()
    own_name = (data.get("own_name") or "").strip()
    if own_gid and own_df is not None and len(own_df) and not own_name:
        sub = own_df[own_df["SPUID"].astype(str).str.strip() == own_gid]
        if len(sub):
            own_name = _safe(sub.iloc[0].get("商品名称"))
    g = _cp.group_link_own(gid, own_gid, own_name)
    if g is None:
        return jsonify({"ok": False, "msg": "组不存在"}), 404
    return jsonify({"ok": True, "group": g})


@app.route("/api/pool/suggest", methods=["POST"])
def api_pool_suggest():
    """跨店绑定推荐（L1条码→L2品牌名称→L3/L4名称→L5语义 本地打分）。
    refresh_suggest：重算并落盘（自愈历史/过期推荐）。"""
    data = request.get_json(force=True)
    eid = (data.get("eid") or "").strip()
    if not eid:
        return jsonify({"ok": False, "msg": "缺少 eid"}), 400
    return jsonify({"ok": True, "suggestions": _cp.refresh_suggest(eid)})


@app.route("/api/pool/suggest/dismiss", methods=["POST"])
def api_pool_suggest_dismiss():
    """忽略一条候选推荐（双向移除，不再出现在推荐弹窗/待确认清单）。"""
    data = request.get_json(force=True)
    eid, eid2 = (data.get("eid") or "").strip(), (data.get("eid2") or "").strip()
    if not eid or not eid2:
        return jsonify({"ok": False, "msg": "缺少 eid"}), 400
    ok = _cp.suggest_dismiss(eid, eid2)
    return jsonify({"ok": ok, "msg": "已忽略" if ok else "无此候选（可能已处理）"})


@app.route("/api/pool/pending/confirm", methods=["POST"])
def api_pool_pending_confirm():
    """批量确认高置信候选对：顺序建组；任一方已入组/同店/不存在则跳过该对。"""
    data = request.get_json(force=True)
    pairs = data.get("pairs") or []
    created, skipped = 0, 0
    for pr in pairs:
        if not isinstance(pr, (list, tuple)) or len(pr) != 2:
            skipped += 1
            continue
        pool = _cp.load_pool()
        eidx = {e["eid"]: e for e in pool["entries"]}
        a, b = eidx.get(str(pr[0])), eidx.get(str(pr[1]))
        if a is None or b is None or a.get("组") or b.get("组") \
                or a.get("shop") == b.get("shop"):
            skipped += 1
            continue
        name = f"{a.get('品牌') or b.get('品牌') or ''} {a.get('品名', '')[:14]}".strip()
        if _cp.create_group([a["eid"], b["eid"]], name):
            created += 1
        else:
            skipped += 1
    return jsonify({"ok": True, "created": created, "skipped": skipped})


@app.route("/api/pool/llm_verify", methods=["POST"])
def api_pool_llm_verify():
    """L6：大模型核实两条目是否同品（结果缓存，同对不重复花钱）。"""
    data = request.get_json(force=True)
    eid, eid2 = (data.get("eid") or "").strip(), (data.get("eid2") or "").strip()
    if not eid or not eid2 or eid == eid2:
        return jsonify({"ok": False, "msg": "参数错误"}), 400
    try:
        out = _cp.llm_verify(eid, eid2)
        return jsonify({"ok": True, **out})
    except RuntimeError as e:
        return jsonify({"ok": False, "msg": str(e)}), 500


@app.route("/api/pool/export")
def api_pool_export():
    """导出候选池为调价分析模板格式（对齐 对比市场价格需调整明细 模板）。
    列结构：市场竞品价格区 → 门店自有价格区 → 调整决策区 → 商品信息区 → 系统扩展区。
    系统有数据的列自动填入，无数据的列留空供手工补充。
    导出时回溯归一化原表，补填入池快照中丢失的字段（category1/2/3、discount_price 等）。"""
    from io import BytesIO

    data = _cp.pool_overview()

    # ── 回溯归一化原表：构建 (shop, sku_id) → row 查找 ──
    # competitor_lookup 已在面板启动时加载，直接用
    def _norm(shop, sku_id):
        """从归一化表查 SKU 原始行，返回 dict（无则空 dict）。"""
        if not sku_id:
            return {}
        row = competitor_lookup.get(shop, {}).get(str(sku_id))
        if row is None:
            return {}
        # pandas Series → dict, 处理 NaN
        return {k: ("" if pd.isna(v) else str(v)) for k, v in row.items()}

    rows = []
    for e in data["entries"]:
        shop = e.get("shop", "")
        skus = e.get("SKU列表") or []
        if not skus:
            skus = [{"规格": "", "价格": e.get("价格", ""), "月售": e.get("月售", 0),
                     "条码": "", "sku_id": "", "图片": ""}]
        for s in skus:
            sku_id = str(s.get("sku_id", ""))
            sku_price = s.get("价格", "") or e.get("价格", "")
            sku_orig = e.get("原价", "")
            # 回溯归一化行
            nr = _norm(shop, sku_id)

            # 折扣计算
            disc = ""
            p = _extract_num(sku_price)
            o = _extract_num(sku_orig)
            if p is not None and o is not None and o > 0:
                r = p / o
                if r < 0.95 and r > 0.10:
                    disc = f"{round(r * 10, 1):g}折"

            # 从归一化行提取补填字段
            cat1 = nr.get("category1", "")
            cat2 = nr.get("category2", "")
            cat3 = nr.get("category3", "") or nr.get("category3_source", "")
            disc_price = nr.get("discount_price", "")
            god_price = nr.get("god_price", "")
            stock = nr.get("stock", "")
            brand_ext = nr.get("品牌（抽取）", "") or nr.get("brand", "")

            # 竞品零售价/活动价 区分：price=零售价，discount_price=活动价
            comp_retail = sku_price
            comp_promo = disc_price if disc_price and disc_price != sku_price else ""

            row = {}
            # ═══ A. 市场竞品价格区（模板 col 1-9）═══
            row["分配"] = ""
            row["市场是否有同款"] = ""
            row["竞品零售价"] = comp_retail
            row["竞品活动价"] = comp_promo
            row["竞品销量"] = s.get("月售", e.get("月售", 0))
            row["其他家零售价"] = ""
            row["其他家活动价"] = ""
            row["其他家销量"] = ""
            row["备注"] = ""

            # ═══ B. 门店自有价格区（模板 col 10-13）═══
            row["门店零售价"] = ""
            row["门店采购价"] = ""
            row["门店月售"] = ""
            row["门店活动价"] = ""

            # ═══ C. 调整决策区（模板 col 14-20，手工填写）═══
            row["调整后零售价"] = ""
            row["调整后活动价"] = ""
            row["调整原因"] = ""
            row["零售价-活动价"] = ""
            row["采购成本占比"] = ""
            row["门店零售价波动"] = ""
            row["门店活动价波动"] = ""

            # ═══ D. 商品信息区（模板 col 21-40）═══
            row["SPU编码"] = e.get("comp_id", "")
            row["商品名称"] = e.get("品名", "")
            row["商品条码"] = s.get("条码", "") or nr.get("barcode", "")
            row["店内分类"] = e.get("类目", "") or cat3
            row["店内分类末级分类"] = cat3
            row["规格名称"] = s.get("规格", "")
            row["京东渠道一级类目"] = cat1
            row["京东渠道二级类目"] = cat2
            row["京东渠道三级类目"] = cat3
            row["京东到家渠道外键编码"] = ""
            row["SKU编码"] = sku_id
            row["商品类型"] = "单品"
            row["外部编码"] = ""
            row["总部零售价"] = ""
            row["饿了么渠道售价"] = ""
            row["美团闪购渠道售价"] = ""
            row["京东秒送渠道售价"] = ""
            row["商品创建时间"] = ""
            row["货品单位1"] = ""
            row["转换比例1"] = ""

            # ═══ E. 调价策略说明（模板最后一列）═══
            row["调价策略说明"] = ""

            # ═══ F. 系统扩展（候选池独有数据，辅助分析）═══
            row["竞品原价"] = sku_orig
            row["竞品折后价"] = disc_price
            row["竞品神价"] = god_price
            row["折扣"] = disc
            row["竞品库存"] = stock
            row["竞品店铺"] = shop
            row["品牌"] = e.get("品牌", "") or brand_ext
            row["规格数"] = e.get("规格数", 0)
            row["绑定状态"] = "已绑定" if e.get("组") else "未绑定"
            row["绑定组ID"] = e.get("组", "")
            row["绑定组名"] = ""
            row["关联自家品组"] = ""
            row["关联自家品名"] = ""
            row["入池时间"] = e.get("入池时间", "")
            row["来源"] = e.get("来源", "")
            row["候选池ID"] = e.get("eid", "")

            rows.append(row)

    # 补填绑定组名称与关联自家品（通过 gid 反查）
    grp_info = {g["gid"]: g for g in data["groups"]}
    for row in rows:
        gid = row["绑定组ID"]
        if gid and gid in grp_info:
            row["绑定组名"] = grp_info[gid].get("名称", "")
            row["关联自家品组"] = grp_info[gid].get("own_gid", "")
            row["关联自家品名"] = grp_info[gid].get("own_name", "")

    # 排序：绑定组在前（按 gid），组内按店铺；未绑定在后
    def _sort_key(r):
        gid = r.get("绑定组ID") or "zzz"
        return (gid, r.get("竞品店铺", ""))
    rows.sort(key=_sort_key)

    df = pd.DataFrame(rows)

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="对比市场价格明细", index=False)
        # 自动列宽
        ws = w.sheets["对比市场价格明细"]
        for col_idx, col_name in enumerate(df.columns, 1):
            max_len = max(len(str(col_name)), df[col_name].astype(str).str.len().max() or 0)
            ws.column_dimensions[ws.cell(1, col_idx).column_letter].width = min(max_len + 4, 35)
    buf.seek(0)
    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=f"候选池_调价明细_{ts}.xlsx")


# ---------------------------------------------------------------------------
# ② 店铺分区 + 区热销榜（区=视图配置；同品合并=候选池绑定组）
# ---------------------------------------------------------------------------
import zone_lib as _zl


def _zone_rank_cached(zone, inc):
    """区榜合并结果缓存（实测每次请求重算：冷 22.6s / 稳态 2.77s）。

    键必须含「决定这份合并结果的全部输入」：区成员店清单 + 勾选子集 + 区配置与
    候选池绑定组的文件时间戳。因为 api_zone_save/delete 并不调用 _invalidate_stats，
    只用 (区名, 子集) 做键的话，改区之后会拿到旧的合并结果。
    返回深拷贝——路由会就地往 items 里补 对应参考/结论/状态/自家价格带 等展示字段。
    """
    import copy

    def _mt(p):
        try:
            return int(os.path.getmtime(p))
        except OSError:
            return 0
    key = (_safe(zone.get("name")), tuple(sorted(inc or [])),
           tuple(sorted(zone.get("comp_shops") or [])),
           _mt(os.path.join(OUTPUT_DIR, "店铺分区.json")),
           _mt(os.path.join(OUTPUT_DIR, "候选池.json")))
    got = _zone_rank_cache.get(key)
    if got is None:
        if len(_zone_rank_cache) > 12:      # 只留最近看过的区/勾选组合
            _zone_rank_cache.clear()
        got = _zl.build_zone_rank(zone, _zone_rank_getter, inc or None)
        _zone_rank_cache[key] = got
    return copy.deepcopy(got)


def _zone_rank_getter(shop):
    """区热销榜数据源：各店品级 top500（缓存复用，无缓存现场构建）。"""
    key5 = (shop, "top500")
    if key5 not in _compare_cache:
        from top50 import build_top50
        _compare_cache[key5] = _derived("top500", shop, lambda: build_top50(shop, 500),
                                        code=_fp_derived())
    return _compare_cache[key5]


def _shop_snapshot_times():
    """各店最新数据快照时间（价格观测历史 per 店 max t；无观测 → ''）。"""
    snaps = {}
    for (s, _sku), (t, _p) in _load_price_hist().items():
        if s not in snaps or t > snaps[s]:
            snaps[s] = t
    return snaps


@app.route("/zone")
def zone_page():
    """分区已整合进首页（区热销榜）与工具页（区管理）；旧链接重定向兼容。"""
    name = (request.args.get("zone") or "").strip()
    if name:
        return redirect(f"/?zone={quote(name)}")
    return redirect("/tools#zone-mgmt")


@app.route("/api/zone/save", methods=["POST"])
def api_zone_save():
    data = request.get_json(force=True)
    z = _zl.save_zone(data.get("name") or "", data.get("own_stores") or [],
                      data.get("comp_shops") or [])
    if z is None:
        return jsonify({"ok": False, "msg": "区名与竞店不能为空"}), 400
    _invalidate_stats()      # 区成员变了 → 区榜合并/依赖区配置的缓存必须作废（不靠 mtime 兜底）
    return jsonify({"ok": True, "zone": z})


@app.route("/api/zone/delete", methods=["POST"])
def api_zone_delete():
    data = request.get_json(force=True)
    ok = _zl.delete_zone((data.get("name") or "").strip())
    _invalidate_stats()      # 同上：删区即释放店铺，视图口径随之变化
    return jsonify({"ok": ok, "msg": "已删除" if ok else "区不存在"})


# ---------------------------------------------------------------------------
# ⑤ 滞销榜：自家门店占压资金排序 + 竞品交叉处置建议
# ---------------------------------------------------------------------------
import slow_lib as _sl


@app.route("/slow")
def slow_page():
    """滞销分析：选自家门店 → 零动销/慢动销榜（占压资金降序）+ 竞品交叉三档处置。"""
    import own_store_lib
    own_stores = own_store_lib.list_stores()
    own_store = (request.args.get("own_store") or "").strip()
    tier_f = (request.args.get("tier") or "all").strip()
    spu_set = None
    if own_store and any(s["name"] == own_store for s in own_stores):
        on_sale = own_store_lib.get_on_sale(own_store)
        if on_sale:
            spu_set = on_sale.get("spu_set")
        else:
            own_store = ""
    lib_sku_df = None
    try:
        _, lib_sku_df = _load_library_cached()
    except Exception:
        lib_sku_df = None
    # 滞销榜全量重算实测 0.7s/次 → 按 (门店, 在售集合) 缓存；id(spu_set) 进键，
    # 保证自家门店重新导入（在售集合被重建）后一定重算，不会拿旧榜。
    _slow_key = (own_store, id(spu_set))
    if _slow_key not in _slow_rank_cache:
        if len(_slow_rank_cache) > 8:          # 只留最近看过的几家店，防内存无限涨
            _slow_rank_cache.clear()
        _slow_rank_cache[_slow_key] = _sl.build_slow_rank(
            own_df, spu_set=spu_set, comp_lookup=competitor_lookup, lib_sku_df=lib_sku_df)
    res = _slow_rank_cache[_slow_key]
    items = res["items"]
    if tier_f == "zero":
        items = [i for i in items if i["层级"] == "零动销"]
    elif tier_f == "slow":
        items = [i for i in items if i["层级"] == "慢动销"]
    scope_bar = build_scope_bar("/slow", own_active=own_store,
                                keep=("tier",), own_stores=own_stores, show_pick=False)
    return render_template("slow.html", active="slow", items=items,
                           stats=res["stats"], own_stores=own_stores,
                           current_store=own_store, tier=tier_f,
                           scope_bar=scope_bar)


if __name__ == "__main__":
    _dev = os.environ.get("PANEL_DEV") == "1"      # 显式开发模式（Flask dev server）
    _no_browser = os.environ.get("PANEL_NO_BROWSER") == "1"  # 服务化运行（NSSM）不开浏览器
    print("正在加载数据...")
    load_data()
    load_progress()
    print(f"候选池 {len(pool_df)} 条（{pool_df['竞店名称'].nunique() if len(pool_df) else 0} 店） | 竞品表按店 {len(competitor_lookup)} 个 | 自家表 {len(own_df)} 行")
    print(f"已确认 {len(progress.get('confirmed', {}))} | 无对应 {len(progress.get('no_match', {}))}")
    _logger.info("面板启动 port=%d auth=%s server=%s", PANEL_PORT, _AUTH_ON,
                 "dev" if _dev else "waitress")
    print(f"启动 Web 面板: http://127.0.0.1:{PANEL_PORT}  鉴权={'开' if _AUTH_ON else '关'}  服务器={'dev' if _dev else 'waitress'}")
    if _AUTH_ON:
        print("  (Basic Auth 已启用：PANEL_USER/PANEL_PASS)")

    # 读盘预热放在开始服务之前：只做指纹有效的快照读取（约 1-2s），
    # 先热好再监听，用户第一次点击就不会撞上"缓存没热 → 现算"（实测首页首次 10s）。
    _t0 = time.time()
    n = _warmup_readonly()
    print(f"读盘预热 {time.time() - _t0:.1f}s，命中 {n} 份结论")
    if not _no_browser:
        threading.Thread(target=open_browser, daemon=True).start()

    if _dev:
        app.run(host="127.0.0.1", port=PANEL_PORT, debug=False, use_reloader=False, threaded=True)
    else:
        # 生产：waitress（官方 WSGI，Windows/Linux 通用；1进程8线程——
        # 内存缓存/线程安全按单进程多线程设计，勿加多进程）。
        # 安全兜底：未设密码时只绑 127.0.0.1（仅本机），设密码才放开局域网；
        # PANEL_HOST 可显式覆盖（如 0.0.0.0）。
        _host = os.environ.get("PANEL_HOST", "").strip() or ("0.0.0.0" if _AUTH_ON else "127.0.0.1")
        from waitress import serve
        serve(app, host=_host, port=PANEL_PORT, threads=8,
              channel_timeout=300, cleanup_interval=30)

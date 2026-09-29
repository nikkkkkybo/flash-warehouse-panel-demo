# -*- coding: utf-8 -*-
"""
聊天式 AI 小助手（assistant_chat）：一个入口、能动手的对话助手。

定位（用户拍板的方向）：不做一排"场景按钮"，改成一个**聊天窗口**——
用户用自然语言提问，助手调**本地工具**（0 成本、当期数据）拿真实数字后作答，
答案里带可点的跳转链接，落到系统里去操作。核心红线：

  R1 数字只能来自工具返回，模型绝不自己造数（提示词硬约束 + 前端只渲染服务端链接）。
  R2 工具全是本地只读查询（榜单/品/变动/台账/自家库/口径），读的就是面板当期缓存
     ——天然满足"不拿旧结论冒充新结论"。
  R3 成本可控：每问一次 = 1~3 次 LLM 轮次（工具轮次本地执行 0 元），max_tokens 900。
  R4 mock/演示模式：真跑工具（给真数字），不调模型（0 元），链路可验证。

分工：本模块只管 提示词/工具 schema/对话循环；工具实现由 web_panel 注入
（ctx["tools"]，闭包引用面板缓存，避免循环 import 与重复读表）。
"""
import json

MAX_ROUNDS = 3          # 工具轮次上限（防循环烧钱）
MAX_TOOL_ROWS = 20      # 单工具返回给模型的行数上限（控 token）
MAX_FOCUS = 24          # 单轮回答附带的关注品上限（防前端撑爆）

# ---------------------------------------------------------------------------
# 工具 schema 单一事实源 = cap_registry（M7 阶段1）：以上内联清单已迁入注册表，
# 新能力在 cap_registry.py 里 register() 一次，两端（聊天/页面）共用。
CAP_REGISTRY = __import__("cap_registry")
TOOLS = CAP_REGISTRY.tools()
TOOL_NAMES = {t["function"]["name"] for t in TOOLS}


# ---------------------------------------------------------------------------
# focus 契约（M7 阶段2）：工具 rows → 结构化关注品集合 → 前端产品卡片
# ---------------------------------------------------------------------------
# 卡片展示字段（从工具行挑非空的，最多 4 个键值，前端按序渲染）
_FOCUS_META_KEYS = (
    ("竞品价", "价格", "到手价", "售价", "竞品价格带", "价格带"),
    ("月售", "有效销量", "有效月售", "竞品有效月售", "月销", "月销合计", "占压资金", "总月售"),
    ("结论", "层级", "状态", "处置建议", "建议", "规则建议"),
    ("自家品名", "自家月销", "自家对应", "品牌", "类目", "竞品价", "竞店"),
)


def _focus_meta(row):
    """从工具行摘展示用键值（原列名→短值），供前端产品卡片。"""
    if not isinstance(row, dict):
        return {}
    meta = {}
    used = set()
    for group in _FOCUS_META_KEYS:
        for k in group:
            if k in used:
                continue
            v = row.get(k)
            if v is None or str(v).strip() == "":
                continue
            meta[k] = str(v)[:24]
            used.add(k)
            break
    return meta


def _focus_url(kind, shop, sku, own_gid, eid):
    """由服务端根据 rows 生成详情 URL；前端只渲染，不拼业务地址。"""
    from urllib.parse import quote
    if kind == "comp" and sku and shop:
        return "/comp_detail?shop=" + quote(str(shop)) + "&sku=" + quote(str(sku))
    if kind == "comp" and sku:
        return "/comp_detail?sku=" + quote(str(sku))
    if kind == "own" and own_gid:
        return "/product_detail?spu=" + quote(str(own_gid))
    if kind == "item" and eid:
        return "/pool"
    return ""


def extract_focus(rows, shop="", limit=MAX_FOCUS):
    """从工具完整 rows（含 `_` 内部键）抽取关注品。

    识别键：_sku/sku_id → 竞品；SPUID/_own_gid → 自家；eid → 候选池。
    每条带 meta（价格/月售/结论等展示字段）与详情 url。
    """
    if not rows:
        return []
    out, seen = [], set()
    shop_d = str(shop or "").strip()

    def _s(v):
        return str(v).strip() if v is not None else ""

    for r in rows:
        if not isinstance(r, dict) or len(out) >= limit:
            break
        sku = _s(r.get("_sku") or r.get("sku_id") or r.get("竞品Sku_id"))
        cid = _s(r.get("_cid") or r.get("comp_id") or r.get("竞品品ID"))
        rshop = _s(r.get("_mshop") or r.get("_shop") or r.get("竞店名称") or shop_d)
        own_gid = _s(r.get("_own_gid") or r.get("SPUID") or r.get("自家SPUID"))
        own_sku = _s(r.get("_own_sku") or r.get("SKUID") or r.get("自家SKUID"))
        eid = _s(r.get("eid") or r.get("候选池ID"))
        name = _s(r.get("品名") or r.get("自家品名") or r.get("商品名称")
                  or r.get("名称") or r.get("结论") or "")

        kind = ""
        if sku or cid:
            kind = "comp"
        elif own_gid or (own_sku and not eid):
            kind = "own"
        elif eid:
            kind = "item"
        if not kind:
            continue

        key = "|".join([kind, rshop if kind == "comp" else "",
                        sku or cid or own_gid or eid, name[:24]])
        if key in seen:
            continue
        seen.add(key)

        if kind == "comp":
            url = _focus_url("comp", rshop, sku or cid, "", "")
            label = name or (sku or cid)
        elif kind == "own":
            url = "/product_detail?spu=" + own_gid if own_gid else ""
            if url and shop_d:
                from urllib.parse import quote as _q
                url += "&shop=" + _q(shop_d)
            label = name or own_gid
        else:
            url = _focus_url("item", "", "", "", eid)
            label = name or eid

        out.append({"kind": kind, "shop": rshop if kind == "comp" else shop_d,
                    "sku": sku if kind == "comp" else "",
                    "comp_id": cid if kind == "comp" else "",
                    "own_gid": own_gid if kind == "own" else "",
                    "name": label[:40], "url": url,
                    "meta": _focus_meta(r)})
    return out


def merge_focus(*lists, limit=MAX_FOCUS):
    """合并多轮/多工具 focus，按 url+name 去重，保序。"""
    out, seen = [], set()
    for lst in lists:
        for f in (lst or []):
            if not isinstance(f, dict):
                continue
            k = (f.get("url") or "") + "|" + (f.get("name") or "")[:24]
            if k in seen:
                continue
            seen.add(k)
            out.append(f)
            if len(out) >= limit:
                return out
    return out

# ---------------------------------------------------------------------------
# 系统提示词（口径由 _bind_lambda 注入，与 sales_text 同源，防漂移）
# ---------------------------------------------------------------------------
SYSTEM_TMPL = """你是「闪电仓选品助手」，嵌在选品面板里的对话助手。用户是前置仓运营，用中文。

可用的工具（本地真实数据，0 成本，读的就是面板当期缓存）：
- query_rank 竞店品级销量榜（shop 传空=全部店合并榜；hot/opp/deflead/all）
- query_product 按关键词查竞店的竞品（含对应状态与详情页链接）
- query_changes 竞店最近两次观测的变动与 undercut 警报（shop 传空=全部店汇总）
- query_insights 洞察台账（规则层发现，带数字证据）
- query_own 自家总库查品（品组月销/售价区间）
- query_caliber 数据口径与新鲜度
- global_search 全局搜品（跨店聚合：在售店数/总有效月售/跨店价格带/对应状态）
- analyze_buy 购入分析底板（品卡规则结论，按有效月售排好）
- analyze_shop 店铺全景底板（统计+机会Top+类目缺口+变动）
- analyze_business 经营分析底板（热销/缺口/类目/价格带/比价五板块）
- analyze_top50 前50逐品底板（按状态给逐品建议）
- brief_boss 老板摘要底板（全局+细节选条，转成给老板的一段话）
- query_slow 滞销榜（自家门店占压资金，零动销/慢动销+处置建议）
- query_pool 候选池（引进候选清单）
- query_onlist 门店上架状态（某品在哪些自家店在售）
- query_history 价格/销量观测历史轨迹
- query_health 系统体检红黄绿灯
- query_learn 学习闭环待办（同义词/品牌补全）
- query_catmap 类目映射表与待确认项
- query_task_status 后台任务/备份/销量口径状态
- query_ai_model 当前 AI 模型/接口/花费（改配置请到工具页，聊天禁止改）
- batch_propose 批量提案预览（同义词/品牌待启用）
- daily_brief 每日速报预览/导出（0 元规则层）

动作类工具（都要两段式：先预览、用户明确同意后才 confirm=true 执行；每步留痕可审计）：
- mark_insight 给台账条目打状态（采纳/已做/忽略）
- confirm_match 确认 竞品品↔自家品组 对应（写入"人工确认·等同级"，可撤销）
- undo_match 撤销对应（整品退回人工池）
- rescan 重新扫描台账（0 元）
- deep_report 每周深报（预览 0 元 → 同意后生成导出）
- enable_synonym 启用/忽略同义词（归一化生效）
- enable_brand_fill 启用/忽略品牌补全（重跑对齐生效）
- confirm_catmap 确认类目映射（竞品→自家）
- start_prescreen 启动 AI 批量预筛（后台，需 key；status_only 查进度）
- run_backup 立即手动备份（后台 zip+自检）
- rerun_align 重跑某店对齐（后台；status_only 查进度）
- run_backfill_sales 销量回填（后台子进程写自家表）
- batch_apply 按提案 id 批量启用同义词/品牌（两段式）

规则：
1. 涉及数字的问题，必须先调工具拿数，绝不凭记忆或常识报数。
2. 只引用工具返回里的数字与品名；工具没给的信息就答"数据里没有"，不猜。
3. 分析类工具（analyze_buy/analyze_shop）返回的数字全部来自本地规则层（产品分析品卡/
   统计），已可信——你的工作是排序、取舍与解释后果（"为什么先做这个"），不得新增数字。
4. 口径：竞品销量列混有收藏（"N人想买"），工具给的"有效销量"=成交+LAMBDA×收藏，
   收藏是弱一档的需求信号；自家月销=补货参考快照实测。两者性质不同，只比量级趋势，不得相减。
5. 回答 ≤120 字：先结论，再 1-2 个关键数字；说清"所以该做什么"。
   **不要逐条复述商品列表**——工具返回的品会以「相关商品」卡片呈现在界面上（带价格/月售/结论，可点开）。
   你只写取舍与动作建议（谁优先、为什么、要不要引进/调价）。纯文本，不要用 markdown 星号/井号。
6. 用户说"这家店/本店/当前店"时指【当前范围】——工具参数 shop 传它，除非用户点名别家店
   或明说"全局/全部店"（此时 shop 传空）；不传 shop 的查询=全店视角，回答里要注明。
7. 需要用户去操作时，引用工具结果里的 link（前端会渲染成按钮）；不要自己编网址。
8. 与选品/数据无关的问题：一句话说明你只负责选品数据，并给一个你能答的例子。
9. 动作类工具（mark_insight/confirm_match/undo_match/rescan）一律两段式：
   第一次调用不传 confirm（得到预览与凭据），把预览讲给用户听；
   只有用户明确说"确认执行/就这么办"后，才第二次调用并传 confirm=true。
   用户没同意就绝不执行；confirm_match 只在证据清晰（双方名称/类目/数字都对得上）时提请确认。

当前范围：{shop}
当前口径（真实注入）：{caliber}"""


def _bind_lambda(text):
    try:
        import sales_text as _STx
        return text.replace("LAMBDA", "λ=%g" % _STx.INTENT_WEIGHT)
    except Exception:
        return text.replace("LAMBDA", "λ")


def _trim_history(messages, limit=10):
    """只留最近 limit 条、且只带 role/content（trace/links 不回传给模型）。"""
    out = []
    for m in (messages or [])[-limit:]:
        if not isinstance(m, dict):
            continue
        r = str(m.get("role") or "")
        c = str(m.get("content") or "").strip()
        if r in ("user", "assistant") and c:
            out.append({"role": r, "content": c[:1200]})
    return out


def _run_tools(model_msg, ctx, trace):
    """执行一轮工具调用；返回 (messages, links, focus)。"""
    import time as _t
    added, links, focus_all = [], [], []
    for tc in (model_msg.get("tool_calls") or [])[:4]:
        fn = (tc.get("function") or {})
        name = str(fn.get("name") or "")
        try:
            args = json.loads(fn.get("arguments") or "{}")
        except (ValueError, TypeError):
            args = {}
        if not isinstance(args, dict):
            args = {}
        impl = (ctx.get("tools") or {}).get(name)
        t0 = _t.time()
        if impl is None:
            result = {"note": f"未知工具 {name}", "rows": [], "links": []}
        else:
            try:
                result = impl(args) or {}
            except Exception as e:
                result = {"note": f"工具执行失败：{e}", "rows": [], "links": []}
        rows = result.get("rows") or []
        # focus 用完整 rows（含 `_` 键）；模型只吃 slim（无 `_` 键）
        fshop = str(args.get("shop") or ctx.get("shop") or "").strip()
        fitems = result.get("focus")
        if not isinstance(fitems, list):
            fitems = extract_focus(rows, shop=fshop)
        focus_all.extend(fitems)
        slim_rows = [{k: v for k, v in r.items() if not str(k).startswith("_")}
                     for r in rows[:MAX_TOOL_ROWS]]
        full = result.get("full")
        if isinstance(full, dict):
            full = {str(k)[:28]: list(v)[:200] for k, v in list(full.items())[:6]
                    if isinstance(v, list)}
        trace.append({"tool": name, "args": {k: str(v)[:40] for k, v in args.items()},
                      "rows": len(rows), "ms": int((_t.time() - t0) * 1000),
                      "note": str(result.get("note") or "")[:120],
                      "data": slim_rows,          # 喂给模型的切片
                      "full": full,               # 完整版（导出用，不喂模型）
                      "focus_n": len(fitems)})
        links.extend(result.get("links") or [])
        slim = {"note": result.get("note", ""),
                "cols": result.get("cols", []),
                "rows": slim_rows}
        added.append({"role": "tool", "tool_call_id": str(tc.get("id") or name),
                      "content": json.dumps(slim, ensure_ascii=False)})
    return added, links, merge_focus(focus_all)


def chat(messages_in, ctx, mock=False, config=None):
    """对话入口。ctx={"shop","shops","caliber","tools":{name:fn}}。
    返回 {"ok", "answer", "trace", "links", "focus", "mock", "shop"}。
    focus = 本轮工具点名的关注品（结构化，前端可点可筛）。"""
    shop = str(ctx.get("shop") or "")
    sys_prompt = _bind_lambda(SYSTEM_TMPL)
    sys_prompt = sys_prompt.replace("{shop}", shop or "（未选店，全店视角）")
    sys_prompt = sys_prompt.replace("{caliber}", str(ctx.get("caliber") or "（未取到）"))
    msgs = [{"role": "system", "content": sys_prompt}] + _trim_history(messages_in)
    trace, links, focus = [], [], []
    if not msgs[-1:] or msgs[-1].get("role") != "user":
        return {"ok": False, "msg": "缺少提问内容", "trace": trace, "links": [],
                "focus": focus, "mock": mock, "shop": shop}

    if mock:
        # 演示模式：不调模型；按关键词挑一个工具真跑（给真数字），拼确定性回答。
        q = msgs[-1]["content"]
        impls = ctx.get("tools") or {}
        pick, args = "query_caliber", {}
        for kw, (t, a) in (("台账", ("query_insights", {})),
                           ("处理", ("query_insights", {})),
                           ("变动", ("query_changes", {})),
                           ("降价", ("query_changes", {})),
                           ("购入", ("analyze_buy", {})),
                           ("经营", ("analyze_business", {})),
                           ("老板", ("brief_boss", {})),
                           ("摘要", ("brief_boss", {})),
                           ("前50", ("analyze_top50", {})),
                           ("怎么样", ("analyze_shop", {})),
                           ("全景", ("analyze_shop", {})),
                           ("机会", ("query_rank", {"filter": "opp"})),
                           ("榜", ("query_rank", {})),
                           ("全局", ("global_search", {})),
                           ("哪些店", ("global_search", {})),
                            ("我家", ("query_own", {})),
                            ("滞销", ("query_slow", {})),
                            ("占压", ("query_slow", {})),
                            ("候选池", ("query_pool", {})),
                            ("引进候选", ("query_pool", {})),
                            ("上架", ("query_onlist", {})),
                            ("在售", ("query_onlist", {})),
                            ("历史", ("query_history", {})),
                            ("体检", ("query_health", {})),
                            ("健康", ("query_health", {})),
                            ("学习", ("query_learn", {})),
                            ("同义词", ("query_learn", {})),
                            ("品牌补全", ("query_learn", {})),
                            ("类目映射", ("query_catmap", {})),
                            ("预筛", ("start_prescreen", {"status_only": True})),
                            ("备份", ("query_task_status", {})),
                            ("后台任务", ("query_task_status", {})),
                            ("回填", ("run_backfill_sales", {"status_only": True})),
                            ("对齐重跑", ("rerun_align", {"status_only": True})),
                            ("模型", ("query_ai_model", {})),
                            ("花费", ("query_ai_model", {})),
                            ("提案", ("batch_propose", {})),
                            ("速报", ("daily_brief", {})),
                            ("深报", ("deep_report", {})),
                            ("周报", ("deep_report", {})),
                            ("价", ("query_product", {})),
                           ("查", ("query_product", {}))):
            if kw in q and t in impls:
                pick, args = t, a
                break
        # 演示模式守卫（M7·注册层红线）：risk>=2 的能力一律不执行（演示只能看）
        if CAP_REGISTRY.risk_of(pick) >= CAP_REGISTRY.RISK_CONFIRM:
            return {"ok": True, "mock": True, "shop": shop,
                    "answer": "（演示模式）该能力属于写/导出操作，演示不执行。配置密钥后可用真实模式。",
                    "trace": [{"tool": pick, "args": {}, "rows": 0, "ms": 0,
                               "note": "演示模式拒绝执行", "data": [], "full": None}],
                    "links": [], "focus": []}
        if pick in ("query_product", "global_search"):
            # 演示模式没有模型帮忙提词：粗清洗问题再当关键词（多词 AND）
            kw2 = q
            for junk in ("哪些店在卖", "哪些店", "在卖", "帮我看", "查一下", "查查",
                         "哪些", "什么价", "多少钱", "怎么样", "有卖", "各自卖",
                         "怎么卖", "吗", "？", "?"):
                kw2 = kw2.replace(junk, "")
            args["keyword"] = kw2.strip() or q
        if pick in ("query_rank", "query_slow", "query_pool", "query_onlist",
                    "query_history", "analyze_buy", "analyze_top50"):
            args.setdefault("shop", shop)
        res = (impls.get(pick) or (lambda a: {"note": "工具未就绪"}))(args) or {}
        fshop = str(args.get("shop") or shop or "").strip()
        fitems = res.get("focus")
        if not isinstance(fitems, list):
            fitems = extract_focus(res.get("rows") or [], shop=fshop)
        focus = merge_focus(fitems)
        trace.append({"tool": pick, "args": {k: str(v)[:40] for k, v in args.items()},
                      "rows": len(res.get("rows") or []), "ms": 0,
                      "note": str(res.get("note") or "")[:120],
                      "data": res.get("rows") or [],
                      "full": res.get("full") if isinstance(res.get("full"), dict) else None,
                      "focus_n": len(focus)})
        links.extend(res.get("links") or [])
        ans = ("（演示模式，未调用 AI）我查了本地数据——\n"
               + str(res.get("note") or ""))
        if res.get("rows"):
            r0 = {k: v for k, v in res["rows"][0].items()
                  if not str(k).startswith("_") and k not in ("link", "url")}
            head = "；".join(f"{k}={v}" for k, v in list(r0.items())[:4] if str(v).strip())
            ans += f"\n第1行：{head}"
        if focus:
            ans += f"\n已点名 {len(focus)} 个相关品（见下方关注品，可点开操作）。"
        ans += "\n配置密钥后即可获得完整解读。"
        return {"ok": True, "answer": ans, "trace": trace, "links": links[:6],
                "focus": focus, "mock": True, "shop": shop}

    from ai_client import call_llm_msg, load_config as _lc
    cfg = config or _lc()
    for _ in range(MAX_ROUNDS):
        msg = call_llm_msg("assistant_chat", msgs, tools=TOOLS, config=cfg,
                           timeout=120, temperature=0.4, max_tokens=900)
        if msg.get("tool_calls"):
            msgs.append({"role": "assistant",
                         "content": str(msg.get("content") or ""),
                         "tool_calls": msg["tool_calls"]})
            added, ln, foc = _run_tools(msg, ctx, trace)
            links.extend(ln)
            focus = merge_focus(focus, foc)
            if not added:
                break
            msgs.extend(added)
            continue
        return {"ok": True, "answer": str(msg.get("content") or "").strip(),
                "trace": trace, "links": links[:6], "focus": focus,
                "mock": False, "shop": shop}
    # 轮次用尽：不带 tools 强制收口
    msgs.append({"role": "user", "content": "请基于以上工具结果直接作答，不要再调用工具。"})
    msg = call_llm_msg("assistant_chat", msgs, config=cfg,
                       timeout=120, temperature=0.4, max_tokens=900)
    return {"ok": True, "answer": str(msg.get("content") or "").strip(),
            "trace": trace, "links": links[:6], "focus": focus,
            "mock": False, "shop": shop}


if __name__ == "__main__":
    print("工具数:", len(TOOLS), list(TOOL_NAMES))
    print(_bind_lambda(SYSTEM_TMPL).splitlines()[0][:60])

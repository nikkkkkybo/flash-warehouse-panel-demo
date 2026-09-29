# -*- coding: utf-8 -*-
"""
M1 助手大脑（assistant_brain）：核心分析场景的增量洞察引擎。

定位：不是自由聊天。每个场景 = 本地已聚合的真实数据 → 强约束提示词 → DeepSeek
返回结构化洞察 JSON。核心目标（用户原话）："从海量内容中分析出人眼逐行看不出、
但值得马上行动的发现"，严禁空谈复述。

场景（M1 四个核心）：
  top   榜单机会解读    —— 机会/防守/领先品卡里"最该先动、最易看走眼"的判断
  cat   类目结构诊断    —— 系统性落后类目（非单品问题）
  price 价格防线        —— 变动里的真 undercut / 无需反应的"伪降价"
  match 对应质量体检    —— 已确认对应里可疑项（跨品牌/销量背离）提醒复查

输出 schema（模型被强约束，normalize 兜底）：
  {"summary": str, "insights": [{"title","evidence","action","name"}]}
    title    ≤18字 一句话发现
    evidence 数字证据（必须来自输入数据）
    action   具体动作（含品名/价格/数值，≤40字）
    name     关联竞品名（前端跳转链接用，无则空）

mock=True 返回本地确定性样例（回归/演示用，不调 API 不花钱）。
"""
import json
import os

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE, "config", "ai_config.yaml")

# ---------------------------------------------------------------------------
# 反空谈硬规则（公共，拼进每个场景 system prompt 尾部）
# ---------------------------------------------------------------------------
ANTI_FLUFF_RULES = """你唯一的职责：从数据里找出「人眼逐行看会漏掉、但值得马上行动」的增量判断。
为此必须遵守：

R1 禁止空话复述：不得出现"建议关注""注意风险""整体向好/偏弱""建议优化"等
   不带数字、不带具体对象的句子。每一条洞察都必须回答"所以我要做什么"。

R2 每条洞察四要素齐全：
   现象（必须引用数据中的具体数字）→ 为什么值得现在做（解读）→
   具体动作（给到品名/价格/数值）→ 关联品名（供跳转）。

R3 只准使用下方数据里的真实数字。数据里没有的信息（成本、毛利、平台政策等）
   明确标注【待核】，绝不编造、绝不脑补。

R4 销量口径（重要）：竞品销量列里混着两种东西——真实成交（"月售35"）与收藏（"35人想买"）。
   系统给你的数据若含"有效销量"，它已是 **成交 + LAMBDA×收藏** 折算后的值（收藏算得轻，但不是不算）。
   因此：① 不得把收藏数当成交笔数；② 不得说收藏无意义，它是弱一档的需求信号；
   ③ 不得再自己乘一次权重；④ 竞品销量与自家月销性质不同，只比量级趋势，不得相减宣称"差距XX"。

R5 宁缺毋滥：这批数据若确实没有值得立即行动的规律，summary 明说"本批未见
   显著可行动规律"，insights 最多给 1 条最轻观察。绝不硬凑、绝不制造焦虑。

R6 排序：最值得马上动手的洞察放最前。数量 2~4 条，不写论文。

R7 严格 JSON 输出（不要 markdown 代码块、不要任何多余文字）：
{"summary":"≤60字总体判断","insights":[{"title":"一句话发现(≤18字)","evidence":"数字证据",
"action":"具体动作(≤40字)","name":"关联竞品名，无则空字符串"}]}"""

def _bind_lambda(text):
    """把提示词里的 LAMBDA 占位符换成 sales_text 的实际权重。
    写死"0.2"迟早会和代码漂移——调权重时提示词必须跟着变。"""
    import sales_text as _STx
    return text.replace("LAMBDA", "λ=%g" % _STx.INTENT_WEIGHT)


ANTI_FLUFF_RULES = _bind_lambda(ANTI_FLUFF_RULES)

# 每场景的【数据来源口径】说明 + 专属任务，拼接在数据前
SCENE_HEAD = {
    "top": (
        "输入=产品分析品卡前 N 行（本地结论引擎已把每个品归为 机会品/防守品/领先品/观察品）。\n"
        "列含义：品名(竞品侧)｜竞品总月售(成交+收藏的原始值)｜竞品有效月售(结论判定用的值，"
        "收藏按 LAMBDA 折过)｜竞品收藏量(被折的那部分)｜自家月销(快照实测)｜在卖品牌数｜结论｜建议。\n"
        "你要做的：①指出最该先动手的 3 个机会品（不能只看月售大小——还要看自家是否真没有、"
        "动作成本是否低）；②揪出「看起来是机会、实际是坑」的（如自家已在卖且卖得不错却被标机会、"
        "或竞品是独有品牌无从对标）；③若机会品扎堆同一品牌/类目，点出这个结构性信号。\n"
        "结论判定口径：机会品=竞品卖得好而自家弱或没有，值得引进/补规格。"),
    "cat": (
        "输入=按类目聚合的产品分析统计（本地已算好：各细类目品数/机会防守领先分布/"
        "竞品总月售合计/自家月销合计/均价）。\n"
        "你要做的：指出「系统性落后」的类目——不是单品的锅，而是整个类目在竞对侧被压制"
        "（例如机会品全类目扎堆、自家合计月销远低于竞品合计）。给出：类目级结论 + 该先进哪个"
        "细类目 + 该从什么价格带切入。也要指出哪些类目看起来热但实际已是自家强项（不需动作）。"),
    "price": (
        "输入=该店最近两次观测的变动明细（本地 price_diff 已算：降价/涨价/新品/断货/回补/"
        "undercut警报）。列含义：变动｜品名｜前值→现值｜幅度｜月销量(平台观测)。\n"
        "你要做的：①挑出真正要应对的 undercut/降价（月售体量大 or 我方同品在卖 or 降幅深），"
        "说明跟/不跟及心理价；②剔除「伪警报」——月售个位数、或竞品降了但仍高于我方在售价、"
        "或只是清库存单品，不需要反应；③警惕：变动里连续多次降价同款 = 竞对在清退/清仓的信号。\n"
        "undercut警报行：现值=竞品现价，前值=自家同品售价。"),
    "match": (
        "输入=已确认对应关系抽样明细（本地对齐引擎输出）。列含义：竞品品名｜自家品名｜竞品品牌｜"
        "自家品牌｜跨品牌｜竞品月售｜自家月销｜结论。\n"
        "你要做的：充当质检员，找出「对应可疑、需要人工复查」的行：①跨品牌但品名看起来不是同一物"
        "（张冠李戴）；②竞品月售大而自家月销=0 或异常低（可能对应错了对象，不是真没卖）；"
        "③品名语义明显不符（规格/材质/功能对不上）。每挑一条都要写清「为什么可疑+建议怎么复核」。\n"
        "判定严格点：你只是怀疑，最终由人决定。"),
    "plan": (
        "输入=**本机洞察台账里未处理的条目**（已由规则层跑全量筛过，每条带数字证据、"
        "算式与连续未处理期数）。你做的是**取舍**，不是复述：假设操作者这周只有一小时，"
        "从里面挑 2-3 件「收益确定 × 动作成本低 × 现在不做会继续扩大」的事并按顺序给。\n"
        "每条 insight：title=要做的动作（动词开头、含品名或类目名）、evidence=只能引用该条目"
        "已有的数字、action=第一步做什么；同时要说明放弃了什么、为什么放弃。\n"
        "如果这些条目都不值得占用工时，直说「本期没有值得马上动手的动作」，不要硬凑。"),
}

SCENE_HEAD = {k: _bind_lambda(v) for k, v in SCENE_HEAD.items()}

MOCK_INSIGHTS = {
    "top": {
        "summary": "mock：榜单数据已就绪（未调真实AI）。见样例洞察格式。",
        "insights": [
            {"title": "样例：最该先动的机会品",
             "evidence": "样例证据，正式调用后为真实数字",
             "action": "样例动作：进入对应页确认该品后处理",
             "name": ""}]},
    "cat": {
        "summary": "mock：类目聚合已就绪。",
        "insights": [
            {"title": "样例：结构性落后类目",
             "evidence": "样例证据",
             "action": "样例动作",
             "name": ""}]},
    "price": {
        "summary": "mock：变动明细已就绪。",
        "insights": [
            {"title": "样例：真 undercut",
             "evidence": "样例证据",
             "action": "样例动作",
             "name": ""}]},
    "plan": {
        "summary": "mock：台账已就绪（未调真实AI）。",
        "insights": [
            {"title": "样例动作：先核这条对应",
             "evidence": "样例证据",
             "action": "样例动作：进台账点「看来源」核对凭据行",
             "name": ""}]},
    "match": {
        "summary": "mock：对应抽样已就绪。",
        "insights": [
            {"title": "样例：可疑对应",
             "evidence": "样例证据",
             "action": "样例动作",
             "name": ""}]},
}


def _truncate(v, n=60):
    s = str(v or "")
    return s if len(s) <= n else s[: n - 1] + "…"


def _as_table(rows, cols):
    """行列表 → 紧凑表格文本（宽度受限，控制 token）。"""
    if not rows:
        return "（无数据行）"
    head = " | ".join(cols)
    lines = [head, "-" * max(len(head), 12)]
    for r in rows:
        cells = [_truncate(r.get(c), 46) for c in cols]
        lines.append(" | ".join(cells))
    return "\n".join(lines)


def _num_field(v):
    try:
        f = float(v)
        return "" if f != f else v  # NaN
    except (TypeError, ValueError):
        return v if v not in ("", None, "nan", "None") else ""


def _clean_rows(rows, cols):
    """只保留给 AI 的列并压缩为 list[dict]（去掉 NaN/None）。"""
    out = []
    for r in rows:
        d = {}
        for c in cols:
            if c not in r:
                continue
            v = r[c]
            if isinstance(v, float) and v != v:  # NaN
                continue
            s = "" if v is None else str(v)
            if not s.strip() or s.strip().lower() in ("nan", "none", "nat"):
                continue
            d[c] = s.strip()
        out.append(d)
    return out


def analyze(scene, data, config=None, mock=False):
    """核心入口。data = {rows: [...], cols: [...], scope: str, shop: str}
    返回 {"summary":..., "insights":[...]}（mock/异常时也保证此结构）。"""
    rows = _clean_rows(data.get("rows", []), data.get("cols", []))
    if not rows:
        return {"summary": "当前范围暂无可用数据，换范围或确认数据已导入。", "insights": []}
    scope_txt = data.get("scope") or data.get("shop") or "当前范围"
    head = SCENE_HEAD.get(scene, "")
    body = _as_table(rows, data.get("cols", []))
    user_txt = (
        f"【分析范围】{scope_txt}\n【数据源】本地已聚合，真实。\n\n{head}\n\n"
        f"数据如下（共 {len(rows)} 行）：\n{body}"
    )

    if mock:
        return MOCK_INSIGHTS.get(scene, MOCK_INSIGHTS["top"])

    try:
        from ai_client import call_llm, load_config as _lc
        cfg = config or _lc()
        if str(cfg.get("driver") or "real").lower() == "fake":
            return MOCK_INSIGHTS.get(scene, MOCK_INSIGHTS["top"])
        raw = call_llm(
            "assistant", 
            [{"role": "system",
              "content": "你是闪电仓/五金零售选品分析员。\n\n" + ANTI_FLUFF_RULES},
             {"role": "user", "content": user_txt}],
            config=cfg, timeout=150, temperature=0.4,
            max_tokens=1800, json_mode=True)
        return _normalize(raw)
    except Exception as e:
        return {"summary": f"分析调用失败：{e}", "insights": [],
                "error": str(e)}


def _normalize(raw):
    """解析模型 JSON；坏格式也尽力提取，保证前端不炸。"""
    text = str(raw or "").strip()
    # 去掉可能的 markdown 代码围栏
    if text.startswith("```"):
        text = text.strip("`")
        if text.startswith("json"):
            text = text[4:]
        text = text.strip()
    obj = None
    try:
        obj = json.loads(text)
    except Exception:
        m = re_search_brace(text)
        if m:
            try:
                obj = json.loads(m)
            except Exception:
                obj = None
    if not isinstance(obj, dict):
        return {"summary": (text[:200] if text else "模型返回无法解析"),
                "insights": [], "raw": text[:500]}
    insights = obj.get("insights")
    if not isinstance(insights, list):
        insights = []
    cleaned = []
    for it in insights:
        if not isinstance(it, dict):
            continue
        cleaned.append({
            "title": _truncate(it.get("title") or "发现", 18),
            "evidence": _truncate(it.get("evidence") or "", 160),
            "action": _truncate(it.get("action") or "", 80),
            "name": _truncate(it.get("name") or "", 46),
        })
    return {
        "summary": _truncate(obj.get("summary") or "", 120),
        "insights": cleaned,
    }


def re_search_brace(text):
    """提取最外层 {..}（json 兜底）。"""
    import re
    m = re.search(r"\{.*\}", text, re.S)
    return m.group(0) if m else None


# 前端场景元数据（/api/assistant/meta）
SCENES = [
    {"id": "plan", "icon": "🎯", "name": "这周只做三件（跨条目取舍）",
     "desc": "读规则层筛好的洞察台账，替你排「收益确定 × 成本低 × 不处理会继续扩大」的 2-3 件事，并说明放弃了什么",
     "est": "约 ¥0.01 · 15-40s"},
    {"id": "top", "icon": "🔥", "name": "机会榜深度解读",
     "desc": "从机会/防守/领先品中挑出最该先动的 3 个，并揪出看着像机会的坑",
     "est": "约 ¥0.02 · 15-40s"},
    {"id": "cat", "icon": "🗂️", "name": "类目结构性落后诊断",
     "desc": "找系统性落后的类目（不是单品问题），告诉你先进哪个细类目、什么价格带",
     "est": "约 ¥0.02 · 15-40s"},
    {"id": "price", "icon": "⚡", "name": "价格防线 · 真 undercut",
     "desc": "从近期变动里挑出真要被跟的降价，剔除月售个位数的伪警报",
     "est": "约 ¥0.02 · 15-40s"},
    {"id": "match", "icon": "🔍", "name": "对应关系体检",
     "desc": "给已确认对应做质检，挑出可能对应错、需要你复查的行",
     "est": "约 ¥0.02 · 15-40s"},
]
SCENE_IDS = {s["id"] for s in SCENES}


# ---------------------------------------------------------------------------
# 洞察解读：规则层已经发现并给出数字，模型只做"解释 + 动作具体化 + 排序"
# （分工理由见 insight_scan.py 头部：模型读的是截断切片，看不见全局，还会复述错数字）
# ---------------------------------------------------------------------------
INTERPRET_RULES = """你收到的是**程序按统计规则跑出来的候选洞察**，数字与凭据已经给定，不是让你找问题。
你只做三件事：
I1 解释「为什么值得现在做」——把后果说清楚（会丢哪类单/让多少利/白占多少资金）。
I2 把动作说具体：给到品名、目标价格数值、第一步做什么；不要写"建议关注/优化/评估"这类空话。
I3 排序：说清哪 1-3 条今天就该动手，哪些其实不值得动（直说理由）。

硬约束：
X1 **禁止引入任何新数字**，只能原样引用条目里已有的数字。数据里没有的信息（成本、毛利、
   周转、平台政策）一律不得推测；缺它就下不了结论时，写【待核：缺什么】。
X2 禁止把标题复述一遍当解释——解释必须比标题多给出一个"所以"。
X3 口径不可混用：竞品销量里的「N人想买」是收藏，系统已按 LAMBDA 折成有效销量；
   不得把它当成交笔数，也不得再乘一次权重。自家销量是补货参考快照实测值，只比量级趋势。
X4 每条都要落回它自己的 key，不得新增、合并或丢弃条目。
X5 严格输出 JSON（不要代码块、不要多余文字）：
{"overall":"≤60字：这批里今天最该动的一件事",
 "notes":[{"key":"原样照抄","explain":"≤70字为什么现在做","action":"≤60字具体动作"}]}"""



# ---------------------------------------------------------------------------
# 追问：就着台账里某一条洞察提问。上下文=该条的凭据行（程序取好的真实数据）
# 模型依然不许报任何未在上下文里出现的数字——它只解释与取舍。
# ---------------------------------------------------------------------------
ASK_RULES = """你在回答关于**一条具体洞察**的追问。上下文给了这条洞察的数字证据、算式，
以及它所引用的真实数据行。

A1 只能引用上下文里出现过的数字与品名；上下文没有的（成本、毛利、库存周转、平台规则）
   一律写【待核：缺什么】，不得推测、不得凭常识补。
A2 回答必须能落到操作：给到品名/规格/价格数值/先做哪一步。不写"建议关注/评估/优化"。
A3 如果按现有数据无法回答（问题超出这条洞察的范围），直说"现有数据答不了，缺 X"，
   不要绕。这比给一个模糊答案有用。
A4 口径：竞品销量里的「N人想买」是收藏，已按 LAMBDA 折成有效销量；不得当成交笔数、
   不得再乘一次权重；自家月销是补货参考快照实测值，两者只比量级趋势，不得相减宣称差距。
A5 直接给结论，≤180字，可分 2-3 小点，不要客套、不要复述问题、不要 markdown 标题。"""

INTERPRET_RULES = _bind_lambda(INTERPRET_RULES)
ASK_RULES = _bind_lambda(ASK_RULES)


def ask(item, rows, question, config=None, mock=False):
    """item=台账条目；rows=凭据行（dict 列表）；question=用户提问。"""
    q = str(question or "").strip()
    if not q:
        return {"ok": False, "msg": "请输入要问的问题"}
    met = item.get("metrics") or {}
    ev = "\n".join(f"  - {k} = {v}" for k, v in met.items())
    lines = []
    for i, r in enumerate((rows or [])[:12], 1):
        kv = "；".join(f"{k}={v}" for k, v in list(r.items())[:12]
                       if str(v).strip() not in ("", "nan", "None"))
        lines.append(f"{i}. {kv}")
    user_txt = (f"【这条洞察】{item.get('title', '')}\n"
                f"范围：{item.get('scope', '')}｜层级：{item.get('layer', '')}"
                f"｜探测器：{item.get('detector', '')}\n"
                f"【数字证据与算式】\n{ev}\n"
                f"【它引用的真实数据行】共 {len(rows or [])} 行，最多列 12 行：\n"
                + ("\n".join(lines) if lines else "（无凭据行）")
                + f"\n\n【我的问题】{q}")

    if mock:
        return {"ok": True, "mock": True, "rows_used": len(rows or []),
                "answer": "（mock）未调用大模型。演示格式：\n"
                          "1) 这条的依据是上下文里的数字，模型不得新增；\n"
                          "2) 动作要给到品名与价格数值；\n"
                          "3) 答不了就直说缺什么。"}
    try:
        from ai_client import call_llm, load_config as _lc
        cfg = config or _lc()
        if str(cfg.get("driver") or "real").lower() == "fake":
            return {"ok": True, "rows_used": len(rows or []),
                    "answer": "（fake 驱动）未调用真实模型"}
        ans = call_llm("insight_ask",
                       [{"role": "system", "content": ASK_RULES},
                        {"role": "user", "content": user_txt}],
                       config=cfg, timeout=120, temperature=0.3, max_tokens=700)
        txt = str(ans or "").strip()
        if not txt:
            return {"ok": False, "msg": "模型没有返回内容"}
        return {"ok": True, "answer": txt, "rows_used": len(rows or [])}
    except Exception as e:
        return {"ok": False, "msg": f"追问失败：{e}"}


def _strip_fence(text):
    t = str(text or "").strip()
    if t.startswith("```"):
        t = t.strip("`")
        if t.startswith("json"):
            t = t[4:]
    return t.strip()


def interpret(findings, config=None, mock=False):
    """→ {"ok", "notes": {key: {explain, action}}, "overall"}；notes 只认合法 key。"""
    items = list(findings or [])[:14]
    if not items:
        return {"ok": False, "msg": "没有可解读的条目"}
    if mock:
        return {"ok": True, "overall": "mock：未调用大模型",
                "notes": {it["key"]: {"explain": "（mock）未调用大模型",
                                      "action": _truncate(it.get("hint", ""), 60)}
                          for it in items}}
    lines = []
    for i, it in enumerate(items, 1):
        met = it.get("metrics") or {}
        ev = "；".join(f"{k}={v}" for k, v in met.items() if k != "算式")
        lines.append(f"{i}. key={it['key']}\n   层级={it.get('layer')}｜范围={it.get('scope')}"
                     f"｜对象={it.get('subject')}\n   发现：{it.get('title')}\n"
                     f"   已有数字：{ev}\n   算法：{met.get('算式', '')}\n"
                     f"   规则层建议：{it.get('hint', '')}")
    user_txt = ("【待解读的候选洞察】共 %d 条。数字与凭据由程序算出，你不得更改：\n\n%s"
                % (len(items), "\n\n".join(lines)))
    try:
        from ai_client import call_llm, load_config as _lc
        cfg = config or _lc()
        if str(cfg.get("driver") or "real").lower() == "fake":
            return {"ok": True, "overall": "（fake 驱动）未调用真实模型",
                    "notes": {it["key"]: {"explain": "", "action": ""} for it in items}}
        raw = call_llm("insight_interpret",
                       [{"role": "system",
                         "content": "你是闪电仓/五金零售选品分析员，负责把统计发现翻译成动作。\n\n"
                                    + INTERPRET_RULES},
                        {"role": "user", "content": user_txt}],
                       config=cfg, timeout=150, temperature=0.3,
                       max_tokens=2000, json_mode=True)
        try:
            obj = json.loads(_strip_fence(raw))
        except Exception:
            m = re_search_brace(_strip_fence(raw))
            try:
                obj = json.loads(m) if m else None
            except Exception:
                obj = None
        if not isinstance(obj, dict):
            return {"ok": False, "msg": "模型返回无法解析"}
        valid = {it["key"] for it in items}
        notes = {}
        for n in obj.get("notes") or []:
            if isinstance(n, dict) and str(n.get("key")) in valid:
                notes[str(n["key"])] = {"explain": _truncate(n.get("explain"), 200),
                                        "action": _truncate(n.get("action"), 160)}
        return {"ok": bool(notes), "notes": notes,
                "overall": _truncate(obj.get("overall"), 120)}
    except Exception as e:
        return {"ok": False, "msg": f"解读调用失败：{e}"}


if __name__ == "__main__":
    # 自测：mock 全场景
    for sc in SCENE_IDS:
        res = analyze(sc, {"rows": [{"品名": "x", "竞品总月售": 100, "自家月销": 0}],
                           "cols": ["品名", "竞品总月售", "自家月销"],
                           "scope": "示例竞店A"}, mock=True)
        print(sc, "->", res["summary"], "| insights:", len(res["insights"]))

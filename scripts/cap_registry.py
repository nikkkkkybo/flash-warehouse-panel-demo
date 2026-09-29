# -*- coding: utf-8 -*-
"""能力注册表（cap_registry）：聊天助手与页面路由共用的"能力插拔层"。

每个能力 = {name, desc, parameters(OpenAI schema), risk, impl}
  risk 分级（写死为红线，见 docs/总任务规划 v2 阶段 5 / 本文件头注释）：
    0 只读        聊天可直接调，0 风险
    1 低风险写    台账类标记，一键 + 自动留痕
    2 两段式确认  预览 → 用户明确同意 → confirm=true 才执行
    3 后台任务    两段式 + 后台跑 + 会话查进度
    4 禁止        聊天只给链接（删店/清数据/改配置），永不注册为可执行工具

约定：
  - schema 单一事实源在这里；assistant_chat.TOOLS 由本表生成；
  - 实现（impl）由 web_panel._chat_ctx 按名字注入（闭包引用面板当期缓存）；
  - mock/演示模式：risk>0 的能力一律拒绝执行（演示只能看）；
  - L1-L3 的留痕由实现自己负责（多数写采纳台账），注册表负责 schema/风险/守卫。
"""
import json

RISK_READONLY = 0
RISK_LOW = 1
RISK_CONFIRM = 2
RISK_TASK = 3
RISK_FORBIDDEN = 4

RISK_TXT = {0: "只读", 1: "低风险写", 2: "两段式", 3: "后台任务", 4: "禁止"}

_REG = []


def register(name, desc, parameters, risk=RISK_READONLY):
    """注册一个能力。risk>=2 的能力 schema 里必须带 confirm 布尔参数（两段式红线）。"""
    if risk >= RISK_CONFIRM:
        props = (parameters or {}).get("properties") or {}
        if "confirm" not in props:
            raise ValueError(f"能力 {name} risk={risk} 必须声明 confirm 参数（两段式）")
    _REG.append({"name": name, "desc": desc, "parameters": parameters or {},
                 "risk": int(risk)})


def tools():
    """生成 OpenAI function-calling 的 tools 列表（assistant_chat.TOOLS 用）。"""
    return [{"type": "function", "function": {"name": e["name"],
                                              "description": e["desc"],
                                              "parameters": e["parameters"]}}
            for e in _REG]


def entries():
    return list(_REG)


def risk_of(name):
    for e in _REG:
        if e["name"] == name:
            return e["risk"]
    return RISK_READONLY


def names():
    return {e["name"] for e in _REG}


def audit_line(name, args, result_note, ms=0):
    """统一动作审计行（L1-L3 调用方自行落 采纳台账 时复用格式）。"""
    return json.dumps({"tool": name, "args": {k: str(v)[:40] for k, v in (args or {}).items()},
                       "note": str(result_note or "")[:120]}, ensure_ascii=False)

# ---------------------------------------------------------------------------
# 默认能力清单（M7 阶段1 自 assistant_chat 迁入；新增能力在此追加注册）
# ---------------------------------------------------------------------------

register("query_rank",
        "查某竞店的品级销量榜（当期缓存）。可按 hot=热销/opp=机会与引进/deflead=防守与领先/all 全量筛选，按月售降序返回。",
        {"type": "object", "properties": {"shop": {"type": "string", "description": "竞店名，缺省=当前店"}, "filter": {"type": "string", "enum": ["hot", "opp", "deflead", "all"]}, "n": {"type": "integer", "description": "返回行数，默认15，最大30"}}, "required": []},
        risk=RISK_READONLY)
register("query_product",
        "按关键词查某竞店里的竞品（品名/品牌/类目多词 AND），返回价格、月售、对应状态、对应的自家品与详情页链接。",
        {"type": "object", "properties": {"shop": {"type": "string"}, "keyword": {"type": "string", "description": "关键词，空格分隔多词 AND"}}, "required": ["keyword"]},
        risk=RISK_READONLY)
register("query_changes",
        "查某竞店最近两次价格观测的变动（降价/涨价/新品/断货/回补/undercut 警报），只报变动清单与警报行。",
        {"type": "object", "properties": {"shop": {"type": "string"}}, "required": []},
        risk=RISK_READONLY)
register("query_insights",
        "查洞察台账（规则层跑全量筛出的发现，每条带数字证据与算式）。可按 层级(全局/细节)、探测器名、店铺 过滤，按严重度降序。",
        {"type": "object", "properties": {"layer": {"type": "string", "enum": ["全局", "细节"]}, "detector": {"type": "string"}, "shop": {"type": "string"}, "n": {"type": "integer"}}, "required": []},
        risk=RISK_READONLY)
register("query_own",
        "在自家总库里按关键词查品（多词 AND），返回品组（SPU 级）月销、售价区间、链接，用于回答\"我家有没有/卖得怎样\"。",
        {"type": "object", "properties": {"keyword": {"type": "string"}}, "required": ["keyword"]},
        risk=RISK_READONLY)
register("query_caliber",
        "查数据口径与新鲜度：竞品数据截至、自家销量快照口径、各店观测期数、λ 折算权重、AI 累计花费。回答\"数据到什么时候/可信吗\"用这个。",
        {"type": "object", "properties": {}, "required": []},
        risk=RISK_READONLY)
register("global_search",
        "全局搜品：跨全部店按关键词聚合（同品跨店按品名归一并），返回在售店数、总有效月售、跨店价格带、自家对应状态；shop 传店名则只在单店内搜。",
        {"type": "object", "properties": {"keyword": {"type": "string"}, "shop": {"type": "string", "description": "留空=跨全部店"}}, "required": ["keyword"]},
        risk=RISK_READONLY)
register("analyze_buy",
        "购入分析底板：产品分析品卡里 结论=机会品/引进候选/上架候选（或指定口径）的候选清单，按竞品有效月售降序，附价格带/自家月销/规格数/规则建议。数字是规则层算好的，你只做取舍排序与解读。",
        {"type": "object", "properties": {"shop": {"type": "string"}, "mode": {"type": "string", "enum": ["机会类", "防守类", "领先类", "全部"]}, "category": {"type": "string", "description": "按类目名过滤（可选）"}, "n": {"type": "integer"}}, "required": []},
        risk=RISK_READONLY)
register("analyze_shop",
        "店铺全景底板：统计（竞品数/待确认/已对应）+ 机会Top + 类目缺口 + 最近变动，规则层拼装好，你做综合解读与下一步建议。",
        {"type": "object", "properties": {"shop": {"type": "string"}}, "required": []},
        risk=RISK_READONLY)
register("analyze_business",
        "经营分析底板：热销与引进/缺口品/类目缺口/价格带/同款比价 五板块结构化数据（本地规则层算好）。shop 传空=前5家店。你做解读；完整版在导出。",
        {"type": "object", "properties": {"shop": {"type": "string"}, "directions": {"type": "string", "description": "逗号分隔，可省略"}}, "required": []},
        risk=RISK_READONLY)
register("analyze_top50",
        "前50逐品底板：销量前50品的匹配状态（已匹配/无对应/待匹配）与双方数字。你按状态给逐品建议：已匹配=优化，无对应=是否引进+参考价，待匹配=先去匹配。",
        {"type": "object", "properties": {"shop": {"type": "string"}}, "required": []},
        risk=RISK_READONLY)
register("brief_boss",
        "老板摘要底板：全局问题+细节可疑 各按严重度选条（标题/依据/动作/连续未处理期数），口径照抄台账。你转成给老板的一段话，不新增数字。",
        {"type": "object", "properties": {"n": {"type": "integer", "description": "总条数，默认12"}}, "required": []},
        risk=RISK_READONLY)
register("mark_insight",
        "动作：给台账条目打处理状态（采纳/已做/忽略）。两段式：先不带 confirm 预览，用户明确同意后才 confirm=true 执行。",
        {"type": "object", "properties": {"key": {"type": "string", "description": "台账条目 key（query_insights 返回）"}, "status": {"type": "string", "enum": ["采纳", "已做", "忽略"]}, "note": {"type": "string"}, "confirm": {"type": "boolean", "description": "仅用户明确同意后传 true"}}, "required": ["key", "status"]},
        risk=RISK_CONFIRM)
register("confirm_match",
        "动作：把竞品品（品级）与自家品组确认为对应（人工确认·等同级，可撤销）。两段式：先预览双方凭据，用户明确同意后才 confirm=true。",
        {"type": "object", "properties": {"shop": {"type": "string"}, "comp_id": {"type": "string", "description": "竞品品ID"}, "own_gid": {"type": "string", "description": "自家品组ID（SPUID，来自 query_own/global_search）"}, "confirm": {"type": "boolean"}}, "required": ["comp_id", "own_gid"]},
        risk=RISK_CONFIRM)
register("undo_match",
        "动作：撤销该竞品品的人工对应（整品，把行退回人工池）。两段式：先预览当前对应凭据，用户明确同意后才 confirm=true。",
        {"type": "object", "properties": {"shop": {"type": "string"}, "comp_id": {"type": "string", "description": "竞品品ID（整品撤销）"}, "sku": {"type": "string", "description": "或单规格 Sku_id"}, "confirm": {"type": "boolean"}}, "required": []},
        risk=RISK_CONFIRM)
register("rescan",
         "动作：重新跑洞察台账规则扫描（0 元，后台约半分钟/店）。两段式：先预览上次扫描信息，用户明确同意后才 confirm=true。",
         {"type": "object", "properties": {
             "confirm": {"type": "boolean"}}, "required": []},
         risk=RISK_TASK)
register("deep_report",
         "动作：每周深报。先给预览（当周新增确认/采纳情况/待回看/未处理Top，规则层 0 元），用户明确同意后才 confirm=true 生成并导出 markdown（可加 ai=true 加 AI 解读，约 ¥0.03）。",
         {"type": "object", "properties": {
             "weeks": {"type": "integer", "description": "统计近几周，默认1，最大4"},
             "ai": {"type": "boolean", "description": "是否加 AI 解读段（需已配置 key）"},
             "confirm": {"type": "boolean", "description": "仅用户明确同意后传 true"}},
             "required": []},
         risk=RISK_READONLY)   # 预览只读可演示；导出仅写非核心报告文件，两段式由实现内 confirm 保证

# ---- M7 阶段2：查询面补全（全 risk=0，focus 契约） ----
register("query_slow",
         "滞销榜：自家门店占压资金排序（零动销/慢动销）+ 竞品交叉处置建议。store 传自家门店名，缺省=总库；tier=zero/slow/all。",
         {"type": "object", "properties": {
             "store": {"type": "string", "description": "自家门店名，缺省=总库"},
             "tier": {"type": "string", "enum": ["zero", "slow", "all"]},
             "n": {"type": "integer", "description": "返回行数，默认15，最大30"}},
          "required": []},
         risk=RISK_READONLY)
register("query_pool",
         "候选池：引进候选清单（已入池的竞品，含绑定组/自家对应/折扣）。可按店、类目过滤。",
         {"type": "object", "properties": {
             "shop": {"type": "string", "description": "竞店名，缺省=全部店"},
             "category": {"type": "string"},
             "n": {"type": "integer", "description": "默认15，最大30"}},
          "required": []},
         risk=RISK_READONLY)
register("query_onlist",
         "门店上架：查某自家门店在售品，或某关键词/SPUID 在各门店的上架状态。回答「这个品哪些店有/没上」。",
         {"type": "object", "properties": {
             "store": {"type": "string", "description": "自家门店名"},
             "keyword": {"type": "string", "description": "按品名关键词筛（可选）"},
             "spuid": {"type": "string", "description": "按自家SPUID查各店上架（可选）"}},
          "required": []},
         risk=RISK_READONLY)
register("query_history",
         "价格/销量观测历史：某竞店（或某品）历次导入快照的到手价与月售轨迹。数据不足两期时会明确说明。",
         {"type": "object", "properties": {
             "shop": {"type": "string", "description": "竞店名"},
             "sku": {"type": "string", "description": "竞品Sku_id（可选，不传则看全店期数概览）"},
             "keyword": {"type": "string", "description": "按品名关键词（可选）"}},
          "required": []},
         risk=RISK_READONLY)
register("query_health",
         "系统体检：每日健康检查（备份/进程/数据新鲜度/模型在岗等红黄绿灯）。only_bad=true 只看异常项。",
         {"type": "object", "properties": {
             "only_bad": {"type": "boolean", "description": "true=只回异常项，默认 true"}},
          "required": []},
         risk=RISK_READONLY)

# ---- M7 阶段3：动作面（学习/类目映射/AI预筛） ----
register("query_learn",
         "学习闭环状态：待启用同义词、待启用品牌补全、样本/词库规模。回答「学习有什么待办」用这个。",
         {"type": "object", "properties": {}, "required": []},
         risk=RISK_READONLY)
register("query_catmap",
         "类目映射：竞品三级类目 → 自家类目 的映射表与待确认项。可按关键词筛竞品类目。",
         {"type": "object", "properties": {
             "keyword": {"type": "string", "description": "按竞品类目名过滤（可选）"},
             "n": {"type": "integer", "description": "默认15，最大30"}},
          "required": []},
         risk=RISK_READONLY)
register("enable_synonym",
         "动作：启用/忽略一条待启用同义词（写入 synonyms.yaml，归一化生效）。两段式：先预览词条，用户同意后 confirm=true。",
         {"type": "object", "properties": {
             "variant": {"type": "string", "description": "变体词（待启用列表里的）"},
             "standard": {"type": "string", "description": "标准词；缺省用待启用记录里的"},
             "enable": {"type": "boolean", "description": "true=启用（默认），false=忽略"},
             "confirm": {"type": "boolean"}},
          "required": ["variant"]},
         risk=RISK_CONFIRM)
register("enable_brand_fill",
         "动作：启用/忽略一条品牌补全（重跑对齐生效）。两段式：先预览 sku/品牌，用户同意后 confirm=true。",
         {"type": "object", "properties": {
             "sku": {"type": "string", "description": "竞品Sku_id"},
             "enable": {"type": "boolean", "description": "true=启用，false=忽略"},
             "confirm": {"type": "boolean"}},
          "required": ["sku"]},
         risk=RISK_CONFIRM)
register("confirm_catmap",
         "动作：确认一条类目映射（竞品三级类目 → 自家类目）。两段式：先预览，同意后 confirm=true。写映射表+内存缓存，收敛层级即时生效。",
         {"type": "object", "properties": {
             "comp_cat": {"type": "string", "description": "竞品三级类目"},
             "own_cat": {"type": "string", "description": "自家类目（三级全名）"},
             "confirm": {"type": "boolean"}},
          "required": ["comp_cat", "own_cat"]},
         risk=RISK_CONFIRM)
register("start_prescreen",
         "动作：启动该店 AI 批量预筛（后台跑人工池待处理条目，需已配置 DeepSeek key）。两段式：先预览状态与条数，同意后 confirm=true。可查 start_prescreen 带 status_only 看进度。",
         {"type": "object", "properties": {
             "shop": {"type": "string", "description": "竞店名，缺省=当前店"},
             "limit": {"type": "integer", "description": "最多预筛条数，默认100"},
             "status_only": {"type": "boolean", "description": "true=只查进度不启动"},
             "confirm": {"type": "boolean"}},
          "required": []},
         risk=RISK_TASK)

# ---- M7 阶段4：后台任务面 ----
register("query_task_status",
         "后台任务与备份状态：最近备份健康度、当前后台任务（备份/对齐/回填/预筛）、销量口径元数据。回答「备份跑了吗/任务做完没」用这个。",
         {"type": "object", "properties": {}, "required": []},
         risk=RISK_READONLY)
register("run_backup",
         "动作：立即手动备份一次数据（zip 不可再生源+自检，后台跑）。两段式：先预览上次备份状态，同意后 confirm=true。",
         {"type": "object", "properties": {
             "confirm": {"type": "boolean"}},
          "required": []},
         risk=RISK_TASK)
register("rerun_align",
         "动作：重跑某竞店对齐（归一化+对齐+汇总库，后台约1-2分钟/店）。两段式：先预览店与标准库是否存在，同意后 confirm=true。完成后面板自动重载。",
         {"type": "object", "properties": {
             "shop": {"type": "string", "description": "竞店名，缺省=当前店"},
             "status_only": {"type": "boolean", "description": "true=只查后台任务进度"},
             "confirm": {"type": "boolean"}},
          "required": []},
         risk=RISK_TASK)
register("run_backfill_sales",
         "动作：销量回填（用最新补货参考快照写回自家表，后台子进程）。两段式：先预览上次口径与快照，同意后 confirm=true。",
         {"type": "object", "properties": {
             "status_only": {"type": "boolean", "description": "true=只查进度"},
             "confirm": {"type": "boolean"}},
          "required": []},
         risk=RISK_TASK)

# ---- M7 阶段5：代理延展 ----
register("query_ai_model",
         "AI 接口与花费：当前模型/base_url/driver/key 是否配置、累计调用与估算花费。回答「用的什么模型/花了多少钱」用这个。改配置请到工具页面板（聊天禁止改配置）。",
         {"type": "object", "properties": {}, "required": []},
         risk=RISK_READONLY)
register("batch_propose",
         "批量提案预览：汇总待启用同义词/品牌补全为可执行提案列表（带 id）。type=全部/同义词/品牌；n 默认20。",
         {"type": "object", "properties": {
             "type": {"type": "string", "enum": ["全部", "同义词", "品牌"]},
             "n": {"type": "integer", "description": "默认20，最大40"}},
          "required": []},
         risk=RISK_READONLY)
register("batch_apply",
         "动作：按提案 id 批量启用（同义词/品牌补全）。两段式：先不带 confirm 预览将执行的 id 列表，用户同意后 confirm=true。ids 逗号分隔或数组。",
         {"type": "object", "properties": {
             "ids": {"type": "string", "description": "提案 id，逗号分隔，来自 batch_propose"},
             "confirm": {"type": "boolean"}},
          "required": ["ids"]},
         risk=RISK_CONFIRM)
register("daily_brief",
         "动作：每日速报（规则层 0 元）：台账未处理/备份/口径/后台任务摘要，导出 output/速报_日期.md。两段式：先预览数字，同意后 confirm=true 导出。",
         {"type": "object", "properties": {
             "confirm": {"type": "boolean"}},
          "required": []},
         risk=RISK_READONLY)  # 预览只读；导出由实现内 confirm 保证（与 deep_report 同策略）


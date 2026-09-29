# -*- coding: utf-8 -*-
"""
AI 预筛（优化项A）：人工池待确认条目批量交大模型预判，
面板列表挂 AI 结论徽章，人工只处理 AI 拿不准/冲突的少数条目。

两种模式（config/ai_config.yaml prescreen.mode，默认 annotate）：
  annotate 标注模式：只写徽章参考，绝不写汇总库（Phase1，默认）
  triage   分流模式：Phase2，高置信自动落地——
      是+conf>=auto_confirm(80)  -> 汇总库 SPU匹配层（确认来源=AI预筛，状态=待抽检）
      否+conf>=auto_reject(85)   -> 无对应标记（进引进评估清单）
      其余                       -> 留人工池带徽章
    撤销过的条目进拉黑名单（ai_prescreen_blacklist.jsonl），不再自动分流。

产物：output/ai_prescreen.json
  {店: {竞品Sku_id: {verdict(是/否/拿不准), confidence(0-1), reason, suggest_own, time, model,
                     triage(自动确认/自动无对应/留人工)}}}
  output/对齐_预筛结果_<店>.xlsx  全量判定明细（审计+撤销依据）
  output/ai_prescreen_blacklist.jsonl  撤销拉黑名单

省钱机制：
  1. 已人工确认/无对应的条目跳过（progress.json 为准）
  2. 单条 AI 判断缓存（_ai_judge_cache.json）已有结论直接采纳
  3. 本预筛结果缓存：重复跑只补新条目
  4. 每 10 条落盘一次，中断不丢

用法：python ai_prescreen.py --shop 店名 [--limit 100]
面板：竞品库页「AI预筛」按钮（后台线程跑，/api/ai/prescreen/status 查进度）
"""
import argparse
import json
import os
import time
from datetime import datetime

import pandas as pd

from ai_client import call_llm, load_config, prescreen_config
from ai_match import judge as ai_judge, cache_get as judge_cache_get, load_examples
from summary_library import DEFAULT_SHOP, upsert_ai

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUTPUT = os.path.join(BASE, "output")
PRESCREEN_FILE = os.path.join(OUTPUT, "ai_prescreen.json")
BLACKLIST_FILE = os.path.join(OUTPUT, "ai_prescreen_blacklist.jsonl")
PROGRESS_FILE = os.path.join(BASE, "scripts", "progress.json")


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def load_results():
    """读取预筛结果缓存。返回 {店: {竞品Sku_id: {...}}}。"""
    if not os.path.exists(PRESCREEN_FILE):
        return {}
    try:
        with open(PRESCREEN_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_results(data):
    tmp = PRESCREEN_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, PRESCREEN_FILE)


def _pool_file(shop):
    """该店人工池文件路径（试跑目录优先，默认店回退单店池）。"""
    trial_f = os.path.join(OUTPUT, f"试跑_{shop}", "对齐_人工确认池.xlsx")
    if os.path.exists(trial_f):
        return trial_f
    if shop == DEFAULT_SHOP and os.path.exists(os.path.join(OUTPUT, "对齐_人工确认池.xlsx")):
        return os.path.join(OUTPUT, "对齐_人工确认池.xlsx")
    return None


def _done_skus(shop):
    """progress.json 中该店已人工处理的 Sku_id 集合（确认+无对应）。"""
    done = set()
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                prog = json.load(f)
            done |= set((prog.get("confirmed") or {}).get(shop, {}).keys())
            done |= set((prog.get("no_match") or {}).get(shop, {}).keys())
        except Exception:
            pass
    return {str(s) for s in done}


def _pending_rows(shop):
    """该店人工池中待处理的行（按 progress.json 过滤已确认/无对应）。"""
    pf = _pool_file(shop)
    if not pf:
        return []
    try:
        df = pd.read_excel(pf, dtype=str).fillna("")
    except Exception:
        return []
    if not len(df) or "竞品Sku_id" not in df.columns:
        return []
    done = _done_skus(shop)
    rows = []
    for _, r in df.iterrows():
        sku = _clean(r.get("竞品Sku_id"))
        if not sku or sku in done:
            continue
        # 池内确认结果列非空也视为已处理（Excel 手工填写的兜底）
        if _clean(r.get("确认结果")):
            continue
        rows.append(r)
    return rows


def _candidates(pool_row):
    """候选自家品名列表：品级候选 JSON top3，兜底单条自家商品名。"""
    out = []
    raw = _clean(pool_row.get("候选自家品"))
    if raw:
        try:
            for c in json.loads(raw)[:3]:
                nm = _clean(c.get("品名")) if isinstance(c, dict) else ""
                if nm:
                    out.append(nm)
        except (ValueError, TypeError):
            pass
    if not out:
        nm = _clean(pool_row.get("自家商品名"))
        if nm:
            out.append(nm)
    return out


def _pool_rows_all(shop):
    """该店人工池全量行（不按 progress 过滤）。返回 {竞品Sku_id: pool_row}。"""
    pf = _pool_file(shop)
    if not pf:
        return {}
    try:
        df = pd.read_excel(pf, dtype=str).fillna("")
    except Exception:
        return {}
    if not len(df) or "竞品Sku_id" not in df.columns:
        return {}
    return {_clean(r.get("竞品Sku_id")): r for _, r in df.iterrows() if _clean(r.get("竞品Sku_id"))}


# ---------------- 拉黑名单（撤销后不再自动分流） ----------------

def blacklist_set(shop):
    """该店已拉黑的竞品Sku_id 集合。"""
    out = set()
    if not os.path.exists(BLACKLIST_FILE):
        return out
    try:
        with open(BLACKLIST_FILE, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line:
                    continue
                try:
                    rec = json.loads(line)
                except (ValueError, TypeError):
                    continue
                if str(rec.get("店") or "") == shop:
                    out.add(str(rec.get("竞品Sku_id") or ""))
    except Exception:
        pass
    return out


def blacklist_add(shop, skus, reason="人工撤销AI分流"):
    """追加拉黑记录（jsonl）。"""
    skus = [str(s) for s in skus if str(s).strip()]
    if not skus:
        return
    with open(BLACKLIST_FILE, "a", encoding="utf-8") as f:
        for s in skus:
            f.write(json.dumps({"店": shop, "竞品Sku_id": s, "t": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                "原因": reason}, ensure_ascii=False) + "\n")


def prescreen_shop(shop, limit=100, progress_cb=None):
    """批量预筛该店人工池待处理条目。
    progress_cb(done, total, ok, fail) 每条回调（面板进度展示）。
    返回统计 dict。"""
    ps_cfg = prescreen_config()
    rows = _pending_rows(shop)
    results = load_results()
    shop_res = results.setdefault(shop, {})

    # 排除已有预筛结论的（缓存只补新）
    todo_all = [r for r in rows if _clean(r.get("竞品Sku_id")) not in shop_res]
    cached_skipped = len(rows) - len(todo_all)
    total = min(len(todo_all), limit) if limit else len(todo_all)
    todo = todo_all[:total]

    config = load_config()
    examples = load_examples()
    ok = fail = adopted = 0
    done = 0
    t0 = time.time()
    for r in todo:
        sku = _clean(r.get("竞品Sku_id"))
        if not sku:
            continue
        # 采纳单条判断缓存（已花过钱的不重复花）
        cached = judge_cache_get(shop, sku)
        if cached and cached.get("verdict") in ("是", "否", "拿不准"):
            rec = dict(cached)
            rec["time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            rec["source"] = "单条缓存"
            shop_res[sku] = rec
            adopted += 1
            ok += 1
        else:
            try:
                res = ai_judge(
                    _clean(r.get("竞品商品名")),
                    _clean(r.get("竞品品牌")),
                    _clean(r.get("竞品型号")),
                    _candidates(r),
                    examples=examples, config=config,
                    module="prescreen",
                )
                shop_res[sku] = {
                    "verdict": res.get("verdict", ""),
                    "confidence": res.get("confidence", ""),
                    "reason": res.get("reason", ""),
                    "suggest_own": res.get("suggest_own", ""),
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "model": config.get("model", ""),
                    "mode": ps_cfg.get("mode", "annotate"),
                }
                ok += 1
            except Exception:
                fail += 1
        done += 1
        if done % 10 == 0:
            save_results(results)          # 断点续跑
        if progress_cb:
            try:
                progress_cb(done, total, ok, fail)
            except Exception:
                pass

    save_results(results)
    return {
        "shop": shop, "pending_total": len(rows), "done": done,
        "ok": ok, "fail": fail, "adopted": adopted,
        "cached_skipped": cached_skipped,
        "elapsed_sec": round(time.time() - t0, 1),
        "remaining": len(todo_all) - done,
    }


def verdict_badge(rec):
    """预筛结论 -> 徽章文案（面板/CLI 共用口径；triage 已分流的追加标记）。"""
    if not rec:
        return ("", "")
    v = str(rec.get("verdict") or "")
    try:
        conf = int(float(rec.get("confidence") or 0) * 100)
    except (ValueError, TypeError):
        conf = 0
    tag = ""
    tri = str(rec.get("triage") or "")
    if tri == "自动确认":
        tag = "·已分流"
    elif tri == "自动无对应":
        tag = "·已判无对应"
    if v == "是":
        return (f"AI同品{conf}{tag}", "ai-yes")
    if v == "否":
        return (f"AI不同{conf}{tag}", "ai-no")
    if v == "拿不准":
        return (f"AI存疑{conf}", "ai-unsure")
    return ("", "")


# ---------------- Phase2：triage 分流落地 ----------------

def _pick_own_group(prow, suggest_name):
    """从候选自家品 JSON 选定自家品组（品组ID, 品名）：
    优先 AI 建议名精确命中 -> 首个候选 -> 池行自家列兜底。"""
    cands = []
    raw = _clean(prow.get("候选自家品"))
    if raw:
        try:
            cands = [c for c in json.loads(raw) if isinstance(c, dict)][:5]
        except (ValueError, TypeError):
            cands = []
    if suggest_name:
        for c in cands:
            if _clean(c.get("品名")) == suggest_name:
                return _clean(c.get("品组ID")), _clean(c.get("品名"))
    if cands:
        return _clean(cands[0].get("品组ID")), _clean(cands[0].get("品名"))
    return _clean(prow.get("自家SPUID")), _clean(prow.get("自家商品名"))


def _write_no_match_direct(shop, marks):
    """CLI 路径：无对应标记直接写 progress.json（原子替换）。
    marks: [(sku, note)]"""
    prog = {}
    if os.path.exists(PROGRESS_FILE):
        try:
            with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                prog = json.load(f)
        except Exception:
            prog = {}
    nm = prog.setdefault("no_match", {}).setdefault(shop, {})
    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for sku, note in marks:
        nm[sku] = {"time": now, "note": note, "source": "AI预筛"}
    tmp = PROGRESS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(prog, f, ensure_ascii=False, indent=2)
    os.replace(tmp, PROGRESS_FILE)


def apply_triage(shop, no_match_writer=None):
    """triage 分流落地（幂等：rec["triage"] 已标记的跳过）。

    规则（方案 §2.2 A2）：
      是+conf>=auto_confirm  -> 汇总库 SPU匹配层（确认来源=AI预筛，状态=待抽检）
      否+conf>=auto_reject   -> 无对应标记（进引进评估清单；suggest_own 非空=AI 自相矛盾，留人工）
      其余                   -> 留人工池带徽章

    安全：已人工处理(done)/拉黑名单 的条目永不自动分流。
    no_match_writer(sku, note)：面板进程传入（走面板内存 progress + save_progress，
    防止与面板态冲突）；缺省 CLI 直接写 progress.json。
    返回统计 dict。"""
    ps_cfg = prescreen_config()
    if ps_cfg.get("mode") != "triage":
        return {"skipped": f"prescreen.mode={ps_cfg.get('mode')}（分流未启用，仅标注）"}

    results = load_results()
    shop_res = results.get(shop) or {}
    if not shop_res:
        return {"skipped": "无预筛结论（先跑 prescreen_shop）"}

    done = _done_skus(shop)
    black = blacklist_set(shop)
    pool_rows = _pool_rows_all(shop)

    # 按置信度降序处理（同品多 SKU 写同一 SPU 行时，最高置信的配对胜出）
    def _conf_of(rec):
        try:
            return float(rec.get("confidence") or 0)
        except (ValueError, TypeError):
            return 0.0

    items = sorted(shop_res.items(), key=lambda kv: _conf_of(kv[1]), reverse=True)

    auto_confirm = auto_no_match = stay_manual = skipped_done = skipped_black = 0
    spu_df = sku_df = None
    nm_marks = []
    for sku, rec in items:
        if sku in done:
            skipped_done += 1
            continue
        if sku in black:
            skipped_black += 1
            continue
        if rec.get("triage"):
            continue  # 已分流过（幂等）
        v = str(rec.get("verdict") or "")
        conf = _conf_of(rec) * 100
        prow = pool_rows.get(sku)
        if v == "是" and conf >= ps_cfg["auto_confirm"] and prow is not None:
            gid, gname = _pick_own_group(prow, _clean(rec.get("suggest_own")))
            spu_row = {
                "竞品品ID": _clean(prow.get("竞品品ID")),
                "竞品商品名": _clean(prow.get("竞品商品名")),
                "竞品品牌": _clean(prow.get("竞品品牌")),
                "竞品型号": _clean(prow.get("竞品型号")),
                "竞品类目": _clean(prow.get("竞品类目")),
                "自家品组ID": gid,
                "自家品名": gname,
                "置信度": round(_conf_of(rec), 2),
            }
            if spu_df is None:
                from summary_library import load_library, save_library
                spu_df, sku_df = load_library()
            spu_df = upsert_ai(spu_df, "竞品品ID", shop, spu_row)
            rec["triage"] = "自动确认"
            auto_confirm += 1
        elif (v == "否" and conf >= ps_cfg["auto_reject"]
              and not _clean(rec.get("suggest_own"))):
            nm_marks.append((sku, f"AI预筛自动判定(否 conf={conf:.0f})"))
            rec["triage"] = "自动无对应"
            auto_no_match += 1
        else:
            rec["triage"] = "留人工"
            stay_manual += 1

    if spu_df is not None:
        from summary_library import save_library
        save_library(spu_df, sku_df)
    if nm_marks:
        if no_match_writer:
            for sku, note in nm_marks:
                no_match_writer(sku, note)
        else:
            _write_no_match_direct(shop, nm_marks)
    save_results(results)
    return {
        "shop": shop, "mode": "triage",
        "auto_confirm": auto_confirm, "auto_no_match": auto_no_match,
        "stay_manual": stay_manual,
        "skipped_done": skipped_done, "skipped_blacklist": skipped_black,
    }


def export_report(shop):
    """导出该店全量预筛明细 Excel（审计+撤销依据）。返回 (行数, 路径)。"""
    shop_res = load_results().get(shop) or {}
    pool_rows = _pool_rows_all(shop)
    v_map = {"是": "同品", "否": "不同品", "拿不准": "存疑"}
    rows = []
    for sku, rec in sorted(shop_res.items()):
        prow = pool_rows.get(sku)
        if prow is None:
            prow = {}
        rows.append({
            "竞品Sku_id": sku,
            "竞品商品名": _clean(prow.get("竞品商品名")),
            "竞品品牌": _clean(prow.get("竞品品牌")),
            "竞品型号": _clean(prow.get("竞品型号")),
            "AI结论": v_map.get(str(rec.get("verdict")), _clean(rec.get("verdict"))),
            "AI置信度": rec.get("confidence", ""),
            "AI理由": _clean(rec.get("reason")),
            "AI建议自家品": _clean(rec.get("suggest_own")),
            "分流": _clean(rec.get("triage")) or "未分流",
            "判定时间": rec.get("time", ""),
        })
    df = pd.DataFrame(rows, columns=["竞品Sku_id", "竞品商品名", "竞品品牌", "竞品型号",
                                     "AI结论", "AI置信度", "AI理由", "AI建议自家品",
                                     "分流", "判定时间"])
    out = os.path.join(OUTPUT, f"对齐_预筛结果_{shop}.xlsx")
    df.to_excel(out, index=False)
    return len(rows), out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--shop", required=True, help="竞店名称（=试跑目录名后缀）")
    ap.add_argument("--limit", type=int, default=100, help="本次最多处理条数（默认100）")
    args = ap.parse_args()

    cfg = load_config()
    if str(cfg.get("driver") or "real").lower() == "fake":
        print("【演示模式】driver=fake：不调真实接口，返回模板结论")
    ps = prescreen_config()
    if not ps.get("enabled"):
        print("提示：prescreen.enabled=false（标注模式可用；仅统计面板不展示计数）")
    print(f"开始预筛：{args.shop}（本次最多 {args.limit} 条）")
    st = prescreen_shop(args.shop, limit=args.limit)
    print(json.dumps(st, ensure_ascii=False, indent=2))
    # 结论分布
    res = load_results().get(args.shop, {})
    dist = {"是": 0, "否": 0, "拿不准": 0}
    for rec in res.values():
        v = str(rec.get("verdict") or "")
        if v in dist:
            dist[v] += 1
    print(f"该店累计预筛 {len(res)} 条 -> 同品 {dist['是']} | 不同 {dist['否']} | 存疑 {dist['拿不准']}")
    # Phase2：triage 分流落地（mode=triage 时生效；annotate 返回 skipped 说明）
    tri = apply_triage(args.shop)
    print("分流落地:", json.dumps(tri, ensure_ascii=False))
    # 全量明细 Excel（审计+撤销依据）
    n, path = export_report(args.shop)
    print(f"预筛明细 -> {path}（{n} 条）")
    if ps.get("mode") == "triage":
        print("分流已落地：AI预筛行(待抽检)在汇总库SPU层；无对应标记在引进评估清单；撤销走面板或 confirm_undo")


if __name__ == "__main__":
    main()

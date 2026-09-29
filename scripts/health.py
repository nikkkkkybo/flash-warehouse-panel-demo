# -*- coding: utf-8 -*-
"""每日体检（P4.4）：把"故障当天就该看见"的检查项集中成一页红黄绿灯。

设计原则：每项检查都必须给出**能照着做的处置**，只报"异常"不报怎么办等于没报。
本模块不 import web_panel（避免循环依赖），一律直接读文件；需要面板内存状态的项
（如重复进程、当前在册店清单）由调用方通过 checks=... 注入。

级别：bad 红（今天就要处理）／warn 黄（本周内处理）／ok 绿／info 中性说明。
"""
import glob
import json
import os
import shutil
import time
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")
LIB = os.path.join(OUT, "竞店标准库")
HIST = os.path.join(LIB, "价格观测历史.jsonl")

_obs_cache = {"mtime": None, "shops": {}, "counts": {}}


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _level(name, level, msg, fix):
    return {"item": name, "level": level, "msg": msg, "fix": fix}


def _observations():
    """逐店：最新观测日期 + 观测天数（期数）。按 jsonl mtime 缓存。"""
    try:
        mt = os.path.getmtime(HIST)
    except OSError:
        return {}, {}
    if _obs_cache["mtime"] == mt:
        return _obs_cache["shops"], _obs_cache["counts"]
    shops, dates = {}, {}
    try:
        with open(HIST, "rb") as f:
            for line in f:
                t = line.strip()
                if not t:
                    continue
                try:
                    rec = json.loads(t.decode("utf-8", "replace"))
                except Exception:
                    continue
                s = str(rec.get("店") or "")
                d = str(rec.get("t") or "")[:10]
                if not s:
                    continue
                shops[s] = max(shops.get(s, ""), d)
                dates.setdefault(s, set()).add(d)
    except OSError:
        return {}, {}
    _obs_cache.update({"mtime": mt, "shops": shops,
                       "counts": {k: len(v) for k, v in dates.items()}})
    return _obs_cache["shops"], _obs_cache["counts"]


def _days_since(datestr):
    try:
        d = datetime.strptime(str(datestr)[:10], "%Y-%m-%d")
        return (datetime.now() - d).days
    except Exception:
        return None


def registered_shops():
    """命令行视角的"在册竞店"：标准库文件 ∪ 试跑目录（面板会注入权威清单）。"""
    s = set()
    for p in glob.glob(os.path.join(LIB, "竞品_*_标准化.xlsx")):
        n = os.path.basename(p)[len("竞品_"):-len("_标准化.xlsx")]
        if n:
            s.add(n)
    for d in glob.glob(os.path.join(OUT, "试跑_*")):
        s.add(os.path.basename(d)[len("试跑_"):])
    return s


def run(shops=None):
    """跑全部检查，返回 {"items": [...], "bad": n, "warn": n, "at": ...}。"""
    items = []
    shops = set(shops or registered_shops())
    obs, periods = _observations()

    # 1 备份是否在跑
    try:
        import backup_daily
        b = backup_daily.status()
    except Exception as e:
        b = {}
        items.append(_level("数据备份", "bad", "读不到备份状态：" + str(e), "手动跑 python scripts/backup_daily.py"))
    if b:
        age = b.get("age_hours")
        if not b.get("newest_zip"):
            items.append(_level("数据备份", "bad", "backup/ 里没有任何备份", "立即跑 backup_daily.py 并注册计划任务"))
        elif not b.get("last_ok"):
            items.append(_level("数据备份", "bad", "最近一次备份自检失败：" + (b.get("last_error") or ""),
                               "看 backup/最近备份.json，确认磁盘空间与关键文件存在"))
        elif age is not None and age > 26:
            items.append(_level("数据备份", "bad", "最新备份已是 %s 小时前（%s）" % (age, b.get("newest_zip")),
                                "计划任务可能没在跑：重跑「注册备份任务.bat」并手动备份一次"))
        else:
            items.append(_level("数据备份", "ok",
                                "%s｜%s 个文件 %.1fMB｜保留 %d 份 %.1fMB"
                                % (b.get("last_at"), b.get("last_files"), b.get("last_mb"),
                                   b.get("count"), b.get("total_mb")), ""))

    # 2 磁盘剩余
    try:
        free_gb = shutil.disk_usage(BASE).free / 1024.0 ** 3
        lv = "ok" if free_gb > 5 else ("warn" if free_gb > 2 else "bad")
        items.append(_level("磁盘剩余", lv, "本盘剩 %.1f GB" % free_gb,
                            "" if lv == "ok" else "清理 backup/ 旧备份或 output/img_cache（可再生）"))
    except OSError:
        pass

    # 3 竞品数据新鲜度
    stale = []
    for s in sorted(shops):
        d = obs.get(s)
        if not d:
            stale.append((s, "无观测"))
        else:
            days = _days_since(d)
            if days is not None and days > 14:
                stale.append((s, "%s（%d 天前）" % (d, days)))
    if not shops:
        items.append(_level("竞品数据", "warn", "一家竞店都没在册", "工具页导入竞品 Excel"))
    elif stale:
        items.append(_level("竞品数据", "warn",
                            "%d 家店数据超 14 天或未观测：%s" % (len(stale), "、".join(x[0][:10] for x in stale)),
                            "重新导出竞品在卖表并导入（同一店再导一次才会产生变动结论）"))
    else:
        items.append(_level("竞品数据", "ok", "%d 家店观测均在 14 天内" % len(shops), ""))

    # 3.5 导入健康：上传了却没落地的店 = 导入中途失败
    try:
        rec_path = os.path.join(BASE, "config", "import_uploads", ".import_records.json")
        with open(rec_path, encoding="utf-8") as f:
            recs = json.load(f)
        imported = {}
        for _h, v in (recs.items() if isinstance(recs, dict) else []):
            imp = (v or {}).get("imported") or {}
            name = str(imp.get("source_name") or "").strip()
            if name:
                imported[name] = str(imp.get("time") or "")
        dangling = []
        for name, at in imported.items():
            std = os.path.join(LIB, "竞品_%s_标准化.xlsx" % name)
            trial = os.path.join(OUT, "试跑_" + name)
            if not os.path.exists(std) and not os.path.isdir(trial):
                dangling.append("%s（%s 上传，未落地）" % (name[:14], at[:10]))
        if dangling:
            items.append(_level("导入健康", "warn",
                                "%d 家店有导入记录但产物缺失：%s" % (len(dangling), "、".join(dangling[:3])),
                                "重新导入这几家的竞品表；失败原因看 output/logs/panel.log"))
        else:
            items.append(_level("导入健康", "ok",
                                "%d 次导入产物齐全（标准库/试跑目录均在）" % len(imported), ""))
    except FileNotFoundError:
        items.append(_level("导入健康", "info", "还没有导入记录", "工具页导入竞品 Excel"))
    except Exception as e:
        items.append(_level("导入健康", "warn", "读导入记录失败：" + str(e), "检查 config/import_uploads/.import_records.json"))

    # 4 观测期数（变动/价格防线结论的前提）
    lt2 = [s for s in shops if int(periods.get(s) or 0) < 2]
    if lt2:
        items.append(_level("变动结论前提", "warn",
                            "%d 家店只有 1 期观测，价格防线/变化摘要/历史轨迹算不出东西" % len(lt2),
                            "对同一家店再导入一次竞品数据即可激活（不是功能坏了）"))
    else:
        items.append(_level("变动结论前提", "ok", "各店观测期数 ≥2", ""))

    # 5 幽灵店
    ghosts = sorted(set(obs) - shops)
    if ghosts:
        items.append(_level("幽灵店残留", "warn",
                            "观测历史里有 %d 家不在册：%s" % (len(ghosts), "、".join(g[:14] for g in ghosts)),
                            "删店时应已自动清理；手动清可跑 scripts/_rebuild_price_hist.py 前的核对流程或调 /api/delete_shop"))
    else:
        items.append(_level("幽灵店残留", "ok", "观测历史与在册店一致", ""))

    # 6 店↔仓映射
    try:
        import store_wh as SW
        m = SW.load()
        try:
            import own_store_lib as O
            stores = [x.get("name") for x in O.list_stores()]
        except Exception:
            stores = []
        unmapped = [s for s in stores if not SW.warehouse_for_store(s)]
        pending = [k for k, v in m.items() if v.get("source") == "pending"]
        if not stores:
            items.append(_level("店↔仓映射", "info", "还没导入自家门店", "工具页·自家门店 导入"))
        elif unmapped:
            items.append(_level("店↔仓映射", "bad",
                                "%d 家自家店未指认补货仓：%s" % (len(unmapped), "、".join(s[:12] for s in unmapped)),
                                "工具页·补货参考 指认（未指认的店页面已禁用店对店销量）"))
        elif pending:
            items.append(_level("店↔仓映射", "warn", "%d 条为系统推定待确认" % len(pending),
                                "工具页·补货参考 逐条点「指认」升级为已确认"))
        else:
            items.append(_level("店↔仓映射", "ok", "%d 家店↔%d 个仓 已人工确认" % (len(stores), len(m)), ""))
    except Exception as e:
        items.append(_level("店↔仓映射", "warn", "读取失败：" + str(e), "检查 config/store_warehouse_map.yaml"))

    # 7 销量回填时效
    try:
        with open(os.path.join(OUT, "own_sales_meta.json"), encoding="utf-8") as f:
            meta = json.load(f)
        days = _days_since(meta.get("bind_time") or meta.get("snapshot_date") or "")
        lv = "ok" if days is not None and days <= 7 else "warn"
        items.append(_level("自家销量回填", lv,
                            "快照 %s（%s）%s" % (meta.get("snapshot_date"), meta.get("note"),
                                            "距今 %d 天" % days if days is not None else ""),
                            "" if lv == "ok" else "把最新「前置仓补货参考」放进 每日补货参考/ 后在工具页执行回填"))
    except Exception:
        items.append(_level("自家销量回填", "warn", "还没回填过（无 own_sales_meta.json）",
                            "工具页·补货参考：选快照 → 指认仓 → 执行回填"))

    # 8 洞察台账与探测器
    try:
        with open(os.path.join(OUT, "AI洞察.json"), encoding="utf-8") as f:
            led = json.load(f)
        items.append(_level("洞察台账", "info",
                            "%d 条｜更新 %s｜口径 %s" % (len(led.get("items") or []), led.get("updated", ""),
                                                    (led.get("caliber") or "")[:40]),
                            "拿不准就先扫描（0 元），条目数字都带算式可复算"))
    except Exception:
        items.append(_level("洞察台账", "info", "台账为空", "分析页点「重新扫描」（不花钱）"))
    try:
        import insight_scan as ISc
        cov = ISc.coverage() or {}
        blocked = ["%s（%s）" % (k, (v.get("blocked") or [""])[0]) for k, v in cov.items()
                   if not v.get("fired") and (v.get("blocked") or [""])[0]]
        if blocked:
            items.append(_level("探测器启用情况", "info", "本期未启用：" + "、".join(blocked),
                                "多为源数据缺列（如「意向」）；下次导出补上即自动生效"))
    except Exception:
        pass

    # 9 快照缓存占用
    try:
        snap = os.path.join(OUT, "_快照")
        files = glob.glob(os.path.join(snap, "*.pkl")) if os.path.isdir(snap) else []
        mb = sum(os.path.getsize(f) for f in files) / 1024.0 ** 2
        lv = "info" if mb < 500 else "warn"
        items.append(_level("快照缓存", lv, "%d 份 / %.0f MB（纯缓存）" % (len(files), mb),
                           "" if lv == "info" else "可整体删除 output/_快照/，面板会按需重算；打包发人时必须排除"))
    except OSError:
        pass

    # 10 采纳追踪（M6·3.3）：建议 → 人做了什么 → 30 天回看
    try:
        import adopt_log as AL
        st = AL.stats()
        pend = AL.pending_review()
        msg = "流水 %d 条（采纳 %d / 已做 %d / 忽略 %d）" % (
            st["total"], st["adopted"], st["done"], st["ignored"])
        if pend:
            lv = "warn"
            msg += "｜⏰ %d 条采纳满 30 天待回看" % len(pend)
            fix = "分析页·采纳追踪 看清单；竞品补第二期观测后即可出效果比对"
        elif st["total"] == 0:
            lv = "info"
            fix = "台账/聊天里对建议做标记，即可开始追踪效果"
        else:
            lv = "info"
            fix = ""
        items.append(_level("采纳追踪", lv, msg, fix))
    except Exception as e:
        items.append(_level("采纳追踪", "info", "读取失败：" + str(e), "检查 output/采纳记录.jsonl"))

    # 11 本地小模型健康度（M6）：FINETUNED 标记 / 模型文件 / 向量自动行
    try:
        marker = os.path.join(BASE, "scripts", "models", "FINETUNED")
        mdir = ""
        if os.path.exists(marker):
            with open(marker, encoding="utf-8") as _mf:
                mdir = (json.load(_mf) or {}).get("ft_dir") or ""
        ft_q = os.path.join(BASE, "scripts", "models", mdir or "bge-small-zh-v1.5-ft3",
                            "model_quantized.onnx")
        if not os.path.exists(marker):
            items.append(_level("本地小模型", "warn",
                                "FINETUNED 标记不在（向量候选层关闭，漏网品不会进人工池）",
                                "与用户确认后再放标记：python scripts/mark_finetuned.py --ft3 --i-confirm"))
        elif not os.path.exists(ft_q):
            items.append(_level("本地小模型", "bad",
                                "标记指向的模型文件缺失：" + os.path.relpath(ft_q, BASE),
                                "重新导出 ONNX 或恢复 scripts/models/ 目录"))
        else:
            n_auto = None
            try:
                import summary_library as _SL
                _spu, _ = _SL.load_library()
                n_auto = int((_spu["确认来源"].astype(str) == "向量自动").sum())
            except Exception:
                n_auto = None
            msg = "ft3 量化模型在位（%s）" % (mdir or "")
            if n_auto is not None:
                msg += "｜汇总库向量自动 %d 行（参考级，可撤销）" % n_auto
            items.append(_level("本地小模型", "info", msg,
                                "汇总库筛「确认来源=向量自动」抽检，可疑点撤销"))
    except Exception as e:
        items.append(_level("本地小模型", "info", "检查失败：" + str(e), ""))

    # 12 AI 调用花费可见性（M6·3.5）：ai_call_log 聚合
    try:
        import ai_client as AC
        ls = AC.log_stats() or {}
        if ls.get("calls", 0):
            items.append(_level("AI 调用花费", "info",
                                "累计 %d 次（成功 %d）｜估算 ¥%.2f｜最近错误 %s"
                                % (ls.get("calls", 0), ls.get("ok", 0),
                                   ls.get("cost_est", 0.0),
                                   (ls.get("last_err") or "")[:40]), ""))
        else:
            items.append(_level("AI 调用花费", "info", "AI 功能未用过（配置 key 后可用）", ""))
    except Exception as e:
        items.append(_level("AI 调用花费", "info", "读取失败：" + str(e), ""))

    bad = sum(1 for i in items if i["level"] == "bad")
    warn = sum(1 for i in items if i["level"] == "warn")
    result = {"items": items, "bad": bad, "warn": warn, "total": len(items),
              "at": time.strftime("%Y-%m-%d %H:%M:%S"),
              "verdict": ("red" if bad else ("yellow" if warn else "green"))}
    # 速报自动留档（M6·3.2）：每日一条到 output/体检历史.jsonl（同日重跑覆盖当日行）
    try:
        fp = os.path.join(OUT, "体检历史.jsonl")
        today = time.strftime("%Y-%m-%d")
        rows = []
        try:
            with open(fp, encoding="utf-8") as f:
                rows = [json.loads(l) for l in f if l.strip()]
        except (FileNotFoundError, json.JSONDecodeError, OSError):
            rows = []
        rows = [r for r in rows if str(r.get("date") or "") != today]
        rows.append({"date": today, "at": result["at"], "verdict": result["verdict"],
                     "bad": bad, "warn": warn, "total": len(items),
                     "top": [f"{i['item']}:{i['level']}" for i in items if i["level"] in ("bad", "warn")][:6]})
        rows = rows[-90:]
        tmp = fp + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            for r in rows:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
        os.replace(tmp, fp)
    except Exception:
        pass
    return result


if __name__ == "__main__":
    import io
    import sys
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    r = run()
    print("体检结果：%s（红 %d / 黄 %d / 共 %d 项）%s"
          % (r["verdict"], r["bad"], r["warn"], r["total"], r["at"]))
    icon = {"bad": "🔴", "warn": "🟡", "ok": "🟢", "info": "ℹ️ "}
    for i in r["items"]:
        print("  %s %-12s %s%s" % (icon.get(i["level"], "?"), i["item"], i["msg"],
                                   ("\n      处置：" + i["fix"]) if i["fix"] else ""))
    sys.exit(1 if r["bad"] else 0)

# -*- coding: utf-8 -*-
"""每日数据备份：zip 打包不可再生数据 + 完整性自检 + 备份状态落盘（供每日体检读取）。

为什么重写（2026-09-04）：计划任务从未注册成功，backup/ 最新备份停在 8/18；
而当天一次误删价格观测历史证明了"唯一的数据保护其实没在跑"是什么后果。
旧版还有一个隐患：把可再生产物一起打包（img_cache 91MB、_快照 79MB、*.bak_*、
事故留存目录），既虚增体积，又把开发期残留发给别人。

三条规矩：
  1 只备不可再生的事实源；可再生缓存一律排除（清单见 EXCLUDE_*）。
  2 备份必须自检：zip 能打开 + 关键文件确实在包内，否则删掉这次产物并报错退出。
  3 结果写 backup/最近备份.json —— 体检页与人工都靠它判断"备份到底有没有在跑"。

用法：python scripts/backup_daily.py            立即备份一次
      python scripts/backup_daily.py --check    只体检不写（返回最近备份健康度）
注册每日 03:30 计划任务：见 注册备份任务.bat
"""
import json
import os
import sys
import zipfile
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")
BACKUP_DIR = os.path.join(BASE, "backup")
STATUS_FILE = os.path.join(BACKUP_DIR, "最近备份.json")
TASK_LOG = os.path.join(BACKUP_DIR, "备份任务.log")   # 计划任务侧留痕（Python 写，编码可控）
KEEP = 14                 # 保留份数
MAX_TOTAL_MB = 1200       # 总占用上限（超了继续丢最旧的），防盘满

# 可再生产物 / 开发期残留：一律不进备份
EXCLUDE_DIRS = {"logs", "img_cache", "__pycache__", "_快照", "_事故留存",
                "_verify_snap_tmp", "_verify_cache_tmp", "_batch_test", "_batch_test2",
                "import_uploads"}
EXCLUDE_FILES = {"ai_config.yaml", ".import_records.json"}
EXCLUDE_PREFIX = ("_verify_", "_dup_test", "_ab_diff", "_bench_", "_pool_llm")
EXCLUDE_EXT = (".lock", ".tmp", ".pkl", ".pyc")
EXCLUDE_MARK = (".bak_", ".bak-", "备份_删店前")   # 备份里不装备份
# 完整性自检：这些文件必须在包里（缺一个就判定备份失败）
MUST_HAVE = ("output/对齐汇总库.xlsx", "output/自家_归一化后.xlsx",
             "scripts/progress.json", "output/竞店标准库/价格观测历史.jsonl")


def _skip_dir(name):
    return name in EXCLUDE_DIRS or name.startswith(".")


def _skip_file(name):
    return (name in EXCLUDE_FILES or name.startswith(EXCLUDE_PREFIX)
            or name.endswith(EXCLUDE_EXT) or any(m in name for m in EXCLUDE_MARK))


def _iter_sources():
    """产出 (绝对路径, 包内相对路径)。"""
    for root, dirs, files in os.walk(OUT):
        dirs[:] = [d for d in dirs if not _skip_dir(d)]
        for f in files:
            if _skip_file(f):
                continue
            p = os.path.join(root, f)
            yield p, os.path.relpath(p, BASE).replace(os.sep, "/")
    extra = [(os.path.join(BASE, "scripts", "progress.json"), "scripts/progress.json"),
             (os.path.join(BASE, "scripts", "scoring_weights.yaml"), "scripts/scoring_weights.yaml"),
             (os.path.join(BASE, "config", "store_warehouse_map.yaml"),
              "config/store_warehouse_map.yaml"),
             (os.path.join(BASE, "config", "category_tree.yaml"), "config/category_tree.yaml"),
             (os.path.join(BASE, "config", "synonyms.yaml"), "config/synonyms.yaml")]
    for p, rel in extra:
        if os.path.exists(p):
            yield p, rel


def make_backup():
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = os.path.join(BACKUP_DIR, "数据备份_%s.zip" % stamp)
    n, rels = 0, set()
    tmp = dest + ".writing"
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
            for p, rel in _iter_sources():
                try:
                    z.write(p, rel)
                    rels.add(rel)
                    n += 1
                except OSError:
                    continue
        # 自检一：zip 结构可用
        with zipfile.ZipFile(tmp) as z:
            bad = z.testzip()
        missing = [m for m in MUST_HAVE if m not in rels]
        if bad or missing:
            os.remove(tmp)
            msg = "zip 内损坏文件 %s" % bad if bad else "关键文件缺失：%s" % "、".join(missing)
            _write_status(dest=None, n=0, mb=0.0, ok=False, error=msg)
            return None, 0, 0.0, msg
        os.replace(tmp, dest)
    except Exception as e:
        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except OSError:
            pass
        _write_status(dest=None, n=0, mb=0.0, ok=False, error=str(e))
        return None, 0, 0.0, str(e)
    mb = os.path.getsize(dest) / 1024.0 / 1024.0
    _prune()
    _write_status(dest=dest, n=n, mb=mb, ok=True, error="")
    return dest, n, mb, ""


def _prune():
    olds = sorted([f for f in os.listdir(BACKUP_DIR)
                   if f.startswith("数据备份_") and f.endswith(".zip")],
                  key=lambda f: os.path.getmtime(os.path.join(BACKUP_DIR, f)))
    while olds:
        total_mb = sum(os.path.getsize(os.path.join(BACKUP_DIR, f)) for f in olds) / 1024.0 / 1024.0
        if len(olds) <= KEEP and total_mb <= MAX_TOTAL_MB:
            break
        victim = olds.pop(0)
        try:
            os.remove(os.path.join(BACKUP_DIR, victim))
        except OSError:
            break


def _write_status(dest, n, mb, ok, error):
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        tmp = STATUS_FILE + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump({"at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                       "file": os.path.basename(dest) if dest else "",
                       "files": n, "size_mb": round(mb, 1),
                       "ok": bool(ok), "error": error},
                      f, ensure_ascii=False, indent=1)
        os.replace(tmp, STATUS_FILE)
    except OSError:
        pass
    # 任务侧留痕：计划任务里跑成功/失败都要在日志里看得见（曾静默停摆 17 天的教训）
    try:
        with open(TASK_LOG, "a", encoding="utf-8") as f:
            f.write("%s %s 文件数=%d 体积=%.1fMB 产物=%s%s\n" % (
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "OK " if ok else "FAIL", n, mb,
                os.path.basename(dest) if dest else "-",
                ("" if ok else " 错误=" + str(error)[:200])))
    except OSError:
        pass


def status():
    """最近一次备份的健康度（每日体检调用）。"""
    zips = sorted([f for f in os.listdir(BACKUP_DIR)
                   if f.startswith("数据备份_") and f.endswith(".zip")]) \
        if os.path.isdir(BACKUP_DIR) else []
    newest = zips[-1] if zips else ""
    newest_path = os.path.join(BACKUP_DIR, newest) if newest else ""
    try:
        with open(STATUS_FILE, encoding="utf-8") as f:
            rec = json.load(f)
    except Exception:
        rec = {}
    age_h = None
    if newest_path:
        age_h = (datetime.now() - datetime.fromtimestamp(os.path.getmtime(newest_path))).total_seconds() / 3600.0
    total_mb = round(sum(os.path.getsize(os.path.join(BACKUP_DIR, z)) for z in zips) / 1024.0 / 1024.0, 1) if zips else 0.0
    return {"last_at": rec.get("at", ""), "last_file": rec.get("file", ""),
            "last_ok": rec.get("ok", False), "last_error": rec.get("error", ""),
            "last_files": rec.get("files", 0), "last_mb": rec.get("size_mb", 0),
            "newest_zip": newest, "age_hours": round(age_h, 1) if age_h is not None else None,
            "count": len(zips), "total_mb": total_mb,
            "keep": KEEP, "status_file": STATUS_FILE}


def main():
    if "--check" in sys.argv:
        print(json.dumps(status(), ensure_ascii=False, indent=1))
        return 0
    dest, n, mb, err = make_backup()
    if not dest:
        print("备份失败:", err)
        return 1
    print("备份完成:", dest)
    print("  %d 个文件, %.1f MB（可再生产物已排除）" % (n, mb))
    st = status()
    print("  当前保留 %d 份 / %.1f MB，最旧备份距今 %s 小时"
          % (st["count"], st["total_mb"], st["age_hours"]))
    return 0


if __name__ == "__main__":
    sys.exit(main())

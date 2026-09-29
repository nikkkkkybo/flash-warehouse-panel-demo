# -*- coding: utf-8 -*-
"""派生结论磁盘快照层：把「每次现算」变成「算一次、复用有据」。

背景（实测）：面板启动读 自家_归一化后.xlsx 需 14.3s，单店品卡重算需 34.9s，
且 5 家店同时并行预热 —— 这是员工本机「点开就等 / 卡死」的根因。

三条铁律（违反就会让旧结论冒充新结论，本层因此失去意义）：
  1 快照只是缓存，不是事实源：只允许缓存「能从上游 100% 重算、且不含任何
    人工写入」的产物。人工确认、候选池、分区配置都是事实源，绝不进本层。
  2 快照必绑依赖指纹：上游文件的 (mtime_ns, size) + 生成器源码指纹。任一项
    变化 → 快照视同不存在 → 重算。因此不存在「读到过期结论」的可能。
  3 宁慢勿错：读写任何异常（文件损坏/pandas 升级/结构变化）一律退回重算，
    写盘用 临时文件 + 原子替换，不留半成品。

存储用 pickle（实测往返无损且比 Excel 快两个数量级；DataFrame 里的 list/tuple
单元格也能原样还原）。只读写本目录自己写下的文件，不加载外部输入。
"""
import hashlib
import inspect
import json
import os
import sys
import time

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SNAP_DIR = os.path.join(BASE, "output", "_快照")
MANIFEST = os.path.join(SNAP_DIR, "清单.json")
SCHEMA = "1"

_disabled = None  # 环境变量 SNAPCACHE=0 可整体关闭（排障/对拍用）


def use_dir(d):
    """重定向快照目录（测试用：不把 t_xxx 测试产物混进生产目录）。"""
    global SNAP_DIR, MANIFEST
    SNAP_DIR = d
    MANIFEST = os.path.join(d, "清单.json")


def enabled():
    global _disabled
    if _disabled is None:
        _disabled = str(os.environ.get("SNAPCACHE", "1")).strip().lower() in ("0", "off", "false")
    return not _disabled


def _safe_name(name):
    keep = [c if (c.isalnum() or c in "._-") else "_" for c in str(name)]
    s = "".join(keep).strip("_")
    return (s[:80] or "snap") + "_" + hashlib.sha1(str(name).encode("utf-8")).hexdigest()[:8]


def _file_tag(path):
    """单个依赖文件的指纹；不存在记 MISSING（缺文件也是依赖状态的一种，
    这样「后来补上的文件」一定会让旧快照失效）。"""
    try:
        st = os.stat(path)
        return "%s|%d|%d" % (os.path.basename(path), st.st_mtime_ns, st.st_size)
    except OSError:
        return "%s|MISSING" % os.path.basename(path)


def code_fingerprint(obj):
    """生成器源码指纹：改了算法，旧快照自动作废。"""
    try:
        src = inspect.getsource(inspect.getmodule(obj) or obj)
    except Exception:
        try:
            src = open(inspect.getfile(obj), "rb").read()
        except Exception:
            return "code?"
    if isinstance(src, str):
        src = src.encode("utf-8", "replace")
    return "code:" + hashlib.sha1(src).hexdigest()[:12]


def deps_fingerprint(deps, code=None):
    h = hashlib.sha1()
    h.update(("v" + SCHEMA + "|").encode("utf-8"))
    # 路径规范化：同一文件的不同写法（d:\ 与 D:\、斜杠方向、多余的 . \\）
    # 必须算作同一个依赖，否则会攒出重复快照、命中率减半。
    for d in sorted(os.path.normcase(os.path.normpath(str(x))) for x in (deps or [])):
        h.update(_file_tag(d).encode("utf-8"))
        h.update(b";")
    if code:
        h.update(("|" + str(code)).encode("utf-8"))
    return h.hexdigest()[:16]


def _code_part(code):
    """规范化 code 参数。

    传模块/函数 → 取模块源码指纹；传字符串 → 视为调用方已算好的指纹直接用作一部分。
    不做这个区分的话，调用方把拼好的指纹串传进来会被 inspect 当对象取源码，
    静默退化成常量 "code?"，算法指纹就等于从来没生效（改了生成器模块也不作废）。
    """
    if code is None:
        return None
    if isinstance(code, str):
        return code
    return code_fingerprint(code)


def key_of(deps, code_obj=None):
    """与 remember() 完全一致的指纹算法。

    任何"想直接读快照而不经过 remember"的调用方（启动预热、体检页）都必须用这个，
    不要自己拼 deps_fingerprint——两边算法一旦不同就永远读不到（实测这样导致预热
    命中 0 份、每次重启偷偷全量重算，而且表面看不出来）。
    """
    return deps_fingerprint(deps, _code_part(code_obj))


def _slot(name):
    s = _safe_name(name)
    return os.path.join(SNAP_DIR, s + ".pkl"), os.path.join(SNAP_DIR, s + ".json")


def peek(name):
    """读快照元数据（不加载数据本体）：{fingerprint, built_at, deps, hits} 或 None。"""
    try:
        with open(_slot(name)[1], "rb") as f:
            return json.loads(f.read().decode("utf-8"))
    except Exception:
        return None


def _manifest_update(name, meta):
    try:
        os.makedirs(SNAP_DIR, exist_ok=True)
        cur = {}
        if os.path.exists(MANIFEST):
            with open(MANIFEST, "rb") as f:
                cur = json.loads(f.read().decode("utf-8"))
        cur[name] = {"fingerprint": meta["fingerprint"], "built_at": meta["built_at"],
                     "deps": meta["deps"]}
        tmp = MANIFEST + ".tmp"
        with open(tmp, "wb") as f:
            f.write(json.dumps(cur, ensure_ascii=False, indent=1).encode("utf-8"))
        os.replace(tmp, MANIFEST)
    except Exception:
        pass  # 清单只服务于展示与体检，坏了不影响正确性


def load(name, fp):
    """命中且指纹一致才返回数据，否则 None。任何异常都当作未命中。"""
    if not enabled():
        return None
    pkl, meta_p = _slot(name)
    try:
        with open(meta_p, "rb") as f:
            meta = json.loads(f.read().decode("utf-8"))
        if meta.get("fingerprint") != fp or meta.get("schema") != SCHEMA:
            return None
        import pandas as pd
        obj = pd.read_pickle(pkl)
        return obj
    except Exception:
        return None


def store(name, fp, payload, deps=None, note=""):
    """原子写：临时文件 → os.replace。写失败只是下次重算，不影响本次返回。"""
    if not enabled():
        return False
    try:
        os.makedirs(SNAP_DIR, exist_ok=True)
        pkl, meta_p = _slot(name)
        tmp = pkl + ".tmp"
        with open(tmp, "wb") as f:
            if isinstance(payload, bytes):
                f.write(payload)
            else:
                _dump(payload, f)
        os.replace(tmp, pkl)
        meta = {"fingerprint": fp, "schema": SCHEMA,
                "built_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "deps": [os.path.basename(str(d)) for d in (deps or [])],
                "note": note}
        with open(meta_p + ".tmp", "wb") as f:
            f.write(json.dumps(meta, ensure_ascii=False, indent=1).encode("utf-8"))
        os.replace(meta_p + ".tmp", meta_p)
        _manifest_update(name, meta)
        return True
    except Exception as e:
        print("[快照] 写入失败（无害，下次重算）:", name, e)
        return False


def _dump(obj, fh):
    import pickle
    pickle.dump(obj, fh, protocol=pickle.HIGHEST_PROTOCOL)


def remember(name, deps, builder, code=None, note=""):
    """核心入口：有合法快照直接用，否则算一次并落盘。

    name    快照名（含店名等，自动清洗）
    deps    上游文件路径列表 —— 快照的可信边界
    builder 无参可调用，返回任意可 pickle 载荷（DataFrame / tuple / dict）
    code    生成器源码（传函数/模块即可），用于「改了算法即作废」
    """
    fp = deps_fingerprint(deps, _code_part(code))
    hit = load(name, fp)
    if hit is not None:
        return hit, True
    t0 = time.time()
    payload = builder()
    dt = time.time() - t0
    store(name, fp, payload, deps=deps, note=note)
    if dt > 0.5:
        print(f"[快照] 重算 {name} 用时 {dt:.1f}s（已落盘，下次秒取）")
    return payload, False


def drop(name):
    for p in _slot(name):
        try:
            os.remove(p)
        except OSError:
            pass
    try:  # 清单条目一起删，否则留下一条永远指向不存在文件的记录
        with open(MANIFEST, "rb") as f:
            names = json.loads(f.read().decode("utf-8"))
        if name in names:
            names.pop(name)
            tmp = MANIFEST + ".tmp"
            with open(tmp, "wb") as f:
                f.write(json.dumps(names, ensure_ascii=False, indent=1).encode("utf-8"))
            os.replace(tmp, MANIFEST)
    except Exception:
        pass


def clear_all():
    """清空全部快照（改算法后手工校验、或体检页的一键重建）。"""
    try:
        for f in os.listdir(SNAP_DIR):
            if f.endswith((".pkl", ".json")):
                os.remove(os.path.join(SNAP_DIR, f))
        return True
    except OSError:
        return False


def prune():
    """清理失效条目与孤儿文件（历史键遗留、手工删过文件、路径规范化前的重名快照）。

    判据：清单里登记的条目若其 .json/.pkl 已不在或指纹已变 → 删；目录里没被任何
    条目引用的 .pkl → 孤儿，删。快照层是纯缓存，删了只会多算一次，不会丢数据。
    """
    try:
        with open(MANIFEST, "rb") as f:
            names = json.loads(f.read().decode("utf-8"))
    except Exception:
        names = {}
    freed, removed = 0, 0
    referenced = set()
    for name, info in list(names.items()):
        pkl, meta_p = _slot(name)
        referenced.update((os.path.basename(pkl), os.path.basename(meta_p)))
        cur = peek(name)
        if cur is None or cur.get("fingerprint") != info.get("fingerprint"):
            for p in (pkl, meta_p):
                try:
                    if os.path.exists(p):
                        freed += os.path.getsize(p)
                        os.remove(p)
                        removed += 1
                except OSError:
                    pass
            names.pop(name, None)
    try:
        for f in os.listdir(SNAP_DIR):
            if f.endswith((".pkl", ".json")) and f not in referenced and f != "清单.json":
                p = os.path.join(SNAP_DIR, f)
                freed += os.path.getsize(p)
                os.remove(p)
                removed += 1
    except OSError:
        pass
    try:
        tmp = MANIFEST + ".tmp"
        with open(tmp, "wb") as f:
            f.write(json.dumps(names, ensure_ascii=False, indent=1).encode("utf-8"))
        os.replace(tmp, MANIFEST)
    except Exception:
        pass
    return {"removed": removed, "freed_mb": round(freed / 1e6, 1)}


def status():
    """快照清单（工具页/每日体检用）。"""
    out = []
    try:
        with open(MANIFEST, "rb") as f:
            names = json.loads(f.read().decode("utf-8"))
    except Exception:
        names = {}
    for name, info in sorted(names.items()):
        pkl, meta_p = _slot(name)
        cur = peek(name) or {}
        out.append({"name": name, "built_at": info.get("built_at"),
                    "deps": info.get("deps"),
                    "alive": bool(cur.get("fingerprint") == info.get("fingerprint")),
                    "size_kb": round(os.path.getsize(pkl) / 1024) if os.path.exists(pkl) else 0})
    return out


if __name__ == "__main__":
    import pandas as pd
    f = os.path.join(BASE, "output", "自家_归一化后.xlsx")
    for i in range(2):
        t0 = time.time()
        obj, hit = remember("own_table", [f],
                            lambda: pd.read_excel(f, dtype=str), code=__import__("snapcache"))
        print(f"第{i+1}次: 命中={hit} 用时={time.time()-t0:.2f}s 行数={len(obj)}")

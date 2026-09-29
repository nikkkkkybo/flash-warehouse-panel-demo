# -*- coding: utf-8 -*-
"""
小模型微调监控面板（train_panel）：独立零依赖 HTTP 服务（http.server，端口 5010）。
- 页面：scripts/templates/train_panel.html
- 接口：GET /api/status（读 output/model_train/train_status.json + 数据/模型状态）
        POST /api/train/start（用 .ftvenv 后台拉起 scripts/finetune_bge.py）
运行：python scripts/train_panel.py
"""
import json
import os
import subprocess
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
HTML_PATH = os.path.join(BASE, "scripts", "templates", "train_panel.html")
TRAIN_DIR = os.path.join(BASE, "output", "model_train")
STATUS_PATH = os.path.join(TRAIN_DIR, "train_status.json")
LOG_PATH = os.path.join(TRAIN_DIR, "train_log.txt")
FT_PY = os.path.join(BASE, ".ftvenv", "Scripts", "python.exe")
FT_SCRIPT = os.path.join(BASE, "scripts", "finetune_bge_v3.py")
DATA_DIR = os.path.join(BASE, "题词向量库")
FT_MODEL = os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft")

PORT = 5010
STALE_SEC = 90
_cache = {"data": None, "ts": 0}
_start_lock = False


def count_lines(path):
    try:
        with open(path, encoding="utf-8") as f:
            return sum(1 for line in f if line.strip())
    except Exception:
        return 0


def data_stats():
    now = os.path.getmtime(os.path.join(DATA_DIR, "mnr_train.jsonl")) \
        if os.path.exists(os.path.join(DATA_DIR, "mnr_train.jsonl")) else 0
    if _cache["data"] and _cache["ts"] == now:
        return _cache["data"]
    stats = {}
    for key, fn in (("mnr", "mnr_train.jsonl"), ("bi", "bi_encoder_train.jsonl"),
                    ("triplet", "triplet_train.jsonl"),
                    ("cross", "cross_encoder_train.jsonl")):
        p = os.path.join(DATA_DIR, fn)
        stats[key] = {"lines": count_lines(p), "mtime": int(os.path.getmtime(p)) if os.path.exists(p) else 0}
    for key, fn in (("pos", "pos_pairs.jsonl"), ("tri", "triplet_v3.jsonl"),
                    ("ev", "eval_pairs.jsonl"), ("hn", "hard_neg_v3.jsonl")):
        p = os.path.join(BASE, "output", "model_train", "v3", fn)
        stats[key] = {"lines": count_lines(p), "mtime": int(os.path.getmtime(p)) if os.path.exists(p) else 0}
    _cache["data"] = stats
    _cache["ts"] = now
    return stats


def model_stats():
    out = {}
    for key, d in (("src", os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-hf")),
                   ("onnx", os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5")),
                   ("ft", os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft3"))):
        bin_ok = os.path.exists(os.path.join(d, "pytorch_model.bin")) or \
            os.path.exists(os.path.join(d, "model.safetensors"))
        out[key] = {"exists": os.path.isdir(d), "ready": bin_ok}
    out["onnx_onnx"] = os.path.exists(
        os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft3", "model_quantized.onnx"))
    out["accepted"] = False
    try:
        with open(os.path.join(BASE, "scripts", "models", "bge-small-zh-v1.5-ft3",
                               "metrics.json"), encoding="utf-8") as f:
            mt = json.load(f)
        out["accepted"] = bool(mt.get("accepted"))
        out["metrics"] = mt.get("best_metrics")
    except Exception:
        pass
    out["finetuned_flag"] = os.path.exists(os.path.join(BASE, "scripts", "models", "FINETUNED"))
    return out


V4_DIR = os.path.join(TRAIN_DIR, "v4")
_v4_cache = {"sig": None, "data": None}


def v4_stats():
    """v4.1 流水线状态：数据计数 / 打标进度 / L4 裁决统计 / 多种子探针。"""
    paths = {k: os.path.join(V4_DIR, fn) for k, fn in (
        ("pos", "pos_pairs.jsonl"), ("ev", "eval_pairs.jsonl"),
        ("tri", "triplet_v4.jsonl"), ("hn", "hard_neg_v4.jsonl"),
        ("llm", "llm_verdicts.jsonl"), ("queue", "queue_compact.jsonl"),
        ("meta", "meta.json"))}
    probe = os.path.join(BASE, "output", "_train_v4", "result.json")
    mtimes = [int(os.path.getmtime(p)) if os.path.exists(p) else 0
              for p in list(paths.values()) + [probe]]
    tdir = os.path.join(BASE, "output", "_train_v4")
    for extra in ("probe_log.txt", "probe6_log.txt", "errprobe_log.txt", "rank_detail_seed43.jsonl"):
        p = os.path.join(tdir, extra)
        mtimes.append(int(os.path.getmtime(p)) if os.path.exists(p) else 0)
    gdir = os.path.join(BASE, "output", "golden_recheck")
    gfiles = [os.path.join(gdir, f) for f in (
        "pending_pairs.jsonl", "golden_v2.jsonl", "verify_result.json", "recheck_log.txt")]
    if os.path.isdir(os.path.join(gdir, "verdicts")):
        gfiles += [os.path.join(gdir, "verdicts", f)
                   for f in os.listdir(os.path.join(gdir, "verdicts")) if f.endswith(".json")]
    for p in gfiles:
        mtimes.append(int(os.path.getmtime(p)) if os.path.exists(p) else 0)
    vdir = os.path.join(V4_DIR, "agent_verdicts")
    if os.path.isdir(vdir):
        try:
            mtimes.append(max(int(os.path.getmtime(os.path.join(vdir, f)))
                              for f in os.listdir(vdir) if f.endswith(".json")))
        except Exception:
            pass
    sig = tuple(mtimes)
    if _v4_cache["sig"] == sig:
        return _v4_cache["data"]
    out = {}
    for k in ("pos", "ev", "tri", "hn", "llm", "queue"):
        out[k] = count_lines(paths[k])
    out["llm_pos"] = out["llm_neg"] = out["llm_null"] = 0
    try:
        with open(paths["llm"], encoding="utf-8") as f:
            for line in f:
                if not line.strip():
                    continue
                s = json.loads(line).get("same")
                if s is True:
                    out["llm_pos"] += 1
                elif s is False:
                    out["llm_neg"] += 1
                else:
                    out["llm_null"] += 1
    except Exception:
        pass
    anchors = set()
    vdir = os.path.join(V4_DIR, "agent_verdicts")
    try:
        for fn in os.listdir(vdir):
            if fn.endswith(".json"):
                with open(os.path.join(vdir, fn), encoding="utf-8") as f:
                    for b in json.load(f).get("b", []):
                        anchors.add(b["i"])
    except Exception:
        pass
    out["anchors_labeled"] = len(anchors)
    out["anchors_total"] = out["queue"]
    out["probe"] = None
    try:
        with open(probe, encoding="utf-8") as f:
            r = json.load(f)
        out["probe"] = {
            "started": r.get("started"), "finished": r.get("finished"),
            "error": r.get("error"),
            "recipe": r.get("recipe"), "gates": r.get("gates"),
            "summary": r.get("summary"),
            "seeds": [{"seed": s.get("seed"), "error": s.get("error"),
                       "best": s.get("best"), "best_epoch": s.get("best_epoch"),
                       "history": s.get("history", [])} for s in r.get("seeds", [])]}
    except Exception:
        pass
    out["probe_log"] = ""
    out["probe_log_file"] = ""
    try:
        logs = [os.path.join(BASE, "output", "_train_v4", f)
                for f in os.listdir(os.path.join(BASE, "output", "_train_v4"))
                if f.startswith("probe") and f.endswith(".txt") and "log" in f]
        if logs:
            newest = max(logs, key=os.path.getmtime)
            out["probe_log"] = tail_log_n(newest, 25)
            out["probe_log_file"] = os.path.basename(newest)
    except Exception:
        pass
    # 错例分析探针（errprobe）
    out["errprobe"] = {"log": "", "model_saved": os.path.isdir(
        os.path.join(BASE, "output", "_train_v4", "errprobe_model_seed43")), "detail_lines": 0}
    out["errprobe"]["log"] = tail_log_n(os.path.join(BASE, "output", "_train_v4", "errprobe_log.txt"), 15)
    try:
        out["errprobe"]["detail_lines"] = count_lines(
            os.path.join(BASE, "output", "_train_v4", "rank_detail_seed43.jsonl"))
    except Exception:
        pass
    # 黄金集 v2 重核（老标准 320+134 → 新标准复核）
    gdir = os.path.join(BASE, "output", "golden_recheck")
    g = {"total": 0, "pos": 0, "neg": 0, "labeled": 0, "confirmed": 0, "flipped": 0,
         "v2_pos": 0, "v2_neg": 0, "verify": None, "log": "", "done": False}
    try:
        pend = os.path.join(gdir, "pending_pairs.jsonl")
        pairs = []
        with open(pend, encoding="utf-8") as f:
            for line in f:
                if line.strip():
                    pairs.append(json.loads(line))
        g["total"] = len(pairs)
        g["pos"] = sum(1 for p in pairs if p.get("old") == "pos")
        g["neg"] = sum(1 for p in pairs if p.get("old") == "neg")
        vdir = os.path.join(gdir, "verdicts")
        seen = {}
        if os.path.isdir(vdir):
            for fn in sorted(os.listdir(vdir)):
                if fn.endswith(".json"):
                    with open(os.path.join(vdir, fn), encoding="utf-8") as f:
                        for b in json.load(f).get("b", []):
                            seen[b["i"]] = b
        g["labeled"] = len(seen)
        g["confirmed"] = sum(1 for i, b in seen.items()
                             if (pairs[i].get("old") == "pos") == (b.get("same") is True))
        g["flipped"] = sum(1 for i, b in seen.items()
                           if (pairs[i].get("old") == "pos") != (b.get("same") is True))
        gv2 = os.path.join(gdir, "golden_v2.jsonl")
        if os.path.exists(gv2):
            n = {"pos": 0, "neg": 0}
            with open(gv2, encoding="utf-8") as f:
                for line in f:
                    if line.strip():
                        n[json.loads(line).get("lbl") or "pos"] += 1
            g["v2_pos"], g["v2_neg"] = n["pos"], n["neg"]
            g["done"] = True
        vr = os.path.join(gdir, "verify_result.json")
        if os.path.exists(vr):
            with open(vr, encoding="utf-8") as f:
                g["verify"] = json.load(f)
    except Exception:
        pass
    g["log"] = tail_log_n(os.path.join(gdir, "recheck_log.txt"), 12)
    out["golden"] = g
    return out


def is_running(st):
    if not st or st.get("status") != "running":
        return False
    import time
    return (time.time() - st.get("updated_ts", 0)) < STALE_SEC


def api_status():
    st = {}
    try:
        with open(STATUS_PATH, encoding="utf-8") as f:
            st = json.load(f)
    except Exception:
        pass
    running = is_running(st)
    if st and st.get("status") == "running" and not running:
        st["status"] = "interrupted"
    st["running"] = running
    st["log_tail"] = tail_log()
    st["data"] = data_stats()
    st["models"] = model_stats()
    st["v4"] = v4_stats()
    st["port"] = PORT
    return st


def tail_log(n=120):
    try:
        with open(LOG_PATH, encoding="utf-8", errors="replace") as f:
            return "".join(f.readlines()[-n:])
    except Exception:
        return ""


def tail_log_n(path, n=25):
    try:
        with open(path, "rb") as f:
            data = f.read()
        lines = data.splitlines(keepends=True)[-n:]
        out = []
        for ln in lines:
            try:
                out.append(ln.decode("utf-8"))
            except UnicodeDecodeError:
                out.append(ln.decode("gbk", errors="replace"))
        return "".join(out)
    except Exception:
        return ""


def ft_process_alive(pid):
    try:
        if not pid:
            return False
        import ctypes
        k32 = ctypes.windll.kernel32
        SYNCHRONIZE = 0x00100000
        h = k32.OpenProcess(SYNCHRONIZE, False, int(pid))
        if h:
            k32.CloseHandle(h)
            return True
    except Exception:
        pass
    return False


def start_train():
    global _start_lock
    if _start_lock:
        return {"ok": False, "msg": "已有启动请求处理中，请勿连点"}
    _start_lock = True
    try:
        st = {}
        try:
            with open(STATUS_PATH, encoding="utf-8") as f:
                st = json.load(f)
        except Exception:
            pass
        if is_running(st) and ft_process_alive(st.get("pid")):
            return {"ok": False, "msg": "训练已在运行中（pid=%s）" % st.get("pid")}
        if not os.path.exists(FT_PY):
            return {"ok": False, "msg": "未找到 .ftvenv 训练环境（.ftvenv\\Scripts\\python.exe）"}
        import time
        time.sleep(1.0)
        try:
            with open(STATUS_PATH, encoding="utf-8") as f:
                st2 = json.load(f)
        except Exception:
            st2 = {}
        if is_running(st2) and st2.get("pid") != st.get("pid"):
            return {"ok": False, "msg": "检测到其他请求刚启动了训练（pid=%s）" % st2.get("pid")}
        os.makedirs(TRAIN_DIR, exist_ok=True)
        flags = 0
        if hasattr(subprocess, "CREATE_NO_WINDOW"):
            flags = subprocess.CREATE_NO_WINDOW
        lf = open(LOG_PATH, "a", encoding="utf-8")
        subprocess.Popen([FT_PY, FT_SCRIPT], cwd=BASE, stdout=lf, stderr=lf,
                         creationflags=flags, stdin=subprocess.DEVNULL)
        return {"ok": True, "msg": "训练已启动"}
    finally:
        threading.Timer(2.0, lambda: globals().__setitem__("_start_lock", False)).start()


class Handler(BaseHTTPRequestHandler):
    def _send(self, code, body, ctype):
        data = body.encode("utf-8") if isinstance(body, str) else body
        self.send_response(code)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", "no-store")
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self):
        path = self.path.split("?")[0]
        if path == "/api/status":
            self._send(200, json.dumps(api_status(), ensure_ascii=False),
                       "application/json; charset=utf-8")
        elif path in ("/", "/train"):
            try:
                with open(HTML_PATH, encoding="utf-8") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except Exception as e:
                self._send(500, str(e), "text/plain; charset=utf-8")
        elif path == "/report":
            rp = os.path.join(BASE, "docs", "模型训练报告_FT3_20260910.html")
            try:
                with open(rp, encoding="utf-8") as f:
                    self._send(200, f.read(), "text/html; charset=utf-8")
            except Exception as e:
                self._send(500, str(e), "text/plain; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")

    def do_POST(self):
        if self.path == "/api/train/start":
            self._send(200, json.dumps(start_train(), ensure_ascii=False),
                       "application/json; charset=utf-8")
        else:
            self._send(404, "not found", "text/plain")

    def log_message(self, fmt, *args):
        pass


if __name__ == "__main__":
    srv = ThreadingHTTPServer(("127.0.0.1", PORT), Handler)
    print(f"训练监控面板 http://127.0.0.1:{PORT}  (Ctrl+C 停止)")
    try:
        srv.serve_forever()
    except KeyboardInterrupt:
        pass

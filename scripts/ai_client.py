# -*- coding: utf-8 -*-
"""
AI 调用统一封装（优化项 D）：
  1. 所有 DeepSeek/OpenAI 兼容调用收口到 call_llm()（限速+重试+超时+决策日志）
  2. 决策日志 output/ai_call_log.jsonl（模块/模型/耗时/成败，可审计可算成本）
  3. fake 演示模式（config ai_config.yaml driver: fake）：不调真实接口，
     返回确定性模板，供演示/测试/CI 全链路跑通不花钱
  4. call_embedding()：null-object 模式——未配置 embeddings 段时返回 None，
     调用方据此退化为现状逻辑（如纯 TF-IDF），永不阻断

配置（config/ai_config.yaml，均可选）：
  api_key / base_url / model        既有三项不变
  driver: real|fake                 默认 real；fake=演示模式
  min_interval_sec: 0.3             真实调用最小间隔（限速，默认0.3秒）
  prescreen:                        AI 预筛（优化项A）配置
    enabled: false                  总开关（默认关）
    mode: annotate|triage           标注模式/分流模式（默认annotate）
    auto_confirm: 80                分流：同品+置信度≥此值 自动写入汇总库
    auto_reject: 85                 分流：非同品+置信度≥此值 自动标无对应
  embeddings:                       语义向量（优化项B/Phase5，未配置即不启用）
    provider / api_key / base_url / model
"""
import json
import os
import time
from datetime import datetime

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CONFIG_PATH = os.path.join(BASE, "config", "ai_config.yaml")
LOG_FILE = os.path.join(BASE, "output", "ai_call_log.jsonl")

DEFAULT_CONFIG = {
    "api_key": "",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-chat",
    "driver": "real",
    "min_interval_sec": 0.3,
}

DEFAULT_PRESCREEN = {
    "enabled": False,
    "mode": "annotate",       # annotate=只标注 | triage=自动分流
    "auto_confirm": 80,
    "auto_reject": 85,
}

# 1 元/百万输入 token 估算（DeepSeek 牌价量级，用于成本可见性，非计费）
PRICE_IN_PER_M = 1.0
PRICE_OUT_PER_M = 2.0

_last_call_ts = 0.0  # 上次真实调用时间戳（进程内限速）


def load_config(path=None):
    """读取 AI 配置（含 driver/prescreen/embeddings 新段）。
    环境变量 DEEPSEEK_API_KEY 优先于配置文件。"""
    path = path or CONFIG_PATH
    cfg = dict(DEFAULT_CONFIG)
    if os.path.exists(path):
        try:
            import yaml
            with open(path, "r", encoding="utf-8") as f:
                data = yaml.safe_load(f) or {}
            for k in cfg:
                if data.get(k) is not None and str(data.get(k)).strip():
                    cfg[k] = str(data[k]).strip()
            cfg["_raw"] = data
        except Exception:
            cfg["_raw"] = {}
    else:
        cfg["_raw"] = {}
    env_key = os.environ.get("DEEPSEEK_API_KEY", "").strip()
    if env_key:
        cfg["api_key"] = env_key
    return cfg


def prescreen_config(config=None):
    """读取预筛配置（带默认值兜底）。"""
    cfg = config or load_config()
    raw = cfg.get("_raw") or {}
    ps = dict(DEFAULT_PRESCREEN)
    if isinstance(raw.get("prescreen"), dict):
        for k in ps:
            v = raw["prescreen"].get(k)
            if v is not None:
                ps[k] = v
    # 数值类型规整
    for k in ("auto_confirm", "auto_reject"):
        try:
            ps[k] = float(ps[k])
        except (ValueError, TypeError):
            ps[k] = DEFAULT_PRESCREEN[k]
    if ps["mode"] not in ("annotate", "triage"):
        ps["mode"] = "annotate"
    return ps


def embeddings_config(config=None):
    """读取 embeddings 配置。未配置/不完整返回 None（null-object：调用方退化为现状）。"""
    cfg = config or load_config()
    raw = cfg.get("_raw") or {}
    eb = raw.get("embeddings")
    if not isinstance(eb, dict):
        return None
    api_key = str(eb.get("api_key") or "").strip()
    model = str(eb.get("model") or "").strip()
    if not api_key or not model:
        return None
    return {
        "api_key": api_key,
        "model": model,
        "base_url": (str(eb.get("base_url") or "").strip().rstrip("/")
                     or "https://open.bigmodel.cn/api/paas/v4"),
    }


def _log(module, model, in_len, out_len, latency_ms, ok, err="", driver="real",
         tokens_in=0, tokens_out=0):
    """追加一条决策日志（失败不阻断主流程）。"""
    try:
        os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)
        cost_est = round(tokens_in / 1e6 * PRICE_IN_PER_M
                         + tokens_out / 1e6 * PRICE_OUT_PER_M, 6)
        rec = {
            "t": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "module": module, "model": model, "driver": driver,
            "in_len": in_len, "out_len": out_len,
            "latency_ms": latency_ms, "ok": ok,
        }
        if err:
            rec["err"] = str(err)[:200]
        if tokens_in:
            rec["tokens_in"] = tokens_in
            rec["tokens_out"] = tokens_out
            rec["cost_est"] = cost_est
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except Exception:
        pass


def _throttle(config):
    """真实调用最小间隔限速（进程内）。"""
    global _last_call_ts
    try:
        interval = float(config.get("min_interval_sec") or 0)
    except (ValueError, TypeError):
        interval = 0.3
    now = time.time()
    wait = _last_call_ts + interval - now
    if wait > 0:
        time.sleep(wait)
    _last_call_ts = time.time()


_FAKE_RESPONSES = {
    # 模块 -> 确定性假响应（json 模块返回 JSON 字符串，其余返回文本）
    "match_judge": json.dumps({"verdict": "拿不准", "confidence": 0.5,
                               "reason": "演示模式：未调用真实接口",
                               "suggest_own": ""}, ensure_ascii=False),
    "prescreen": json.dumps({"verdict": "拿不准", "confidence": 0.5,
                             "reason": "演示模式：未调用真实接口",
                             "suggest_own": ""}, ensure_ascii=False),
    "pool_bind_llm": json.dumps({"verdict": "拿不准", "confidence": 0.5,
                                 "reason": "演示模式：未调用真实接口"},
                                ensure_ascii=False),
    "keyword_lib": "{}",   # 词库提取：fake 模式返回空结果（不产生词条）
}


def call_llm(module, messages, config=None, timeout=120, temperature=0.1,
             max_tokens=300, json_mode=False):
    """统一 LLM 调用入口。返回模型输出文本（str）。
    driver=fake 时返回确定性假响应（不调真实接口、不花钱）。
    失败重试1次后抛 RuntimeError（含 401 明确提示）。"""
    config = config or load_config()
    driver = (config.get("driver") or "real").lower()
    model = config.get("model") or DEFAULT_CONFIG["model"]
    in_len = sum(len(str(m.get("content") or "")) for m in messages)

    # ---- fake 演示模式 ----
    if driver == "fake":
        _log(module, model, in_len, 0, 0, True, driver="fake")
        if module in _FAKE_RESPONSES:
            return _FAKE_RESPONSES[module]
        if json_mode:
            return json.dumps({"demo": True}, ensure_ascii=False)
        return "【演示模式】未调用真实 AI 接口（config ai_config.yaml driver: fake）。"

    # ---- real 真实调用 ----
    api_key = (config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("未配置 AI 接口密钥，请先在「选店分析」页配置 DeepSeek key")
    base = (config.get("base_url") or "").strip().rstrip("/") or DEFAULT_CONFIG["base_url"]
    url = base + "/chat/completions"
    payload = {
        "model": model,
        "messages": messages,
        "temperature": temperature,
        "max_tokens": max_tokens,
        "stream": False,
    }
    if json_mode:
        payload["response_format"] = {"type": "json_object"}
    headers = {"Authorization": f"Bearer {api_key}",
               "Content-Type": "application/json"}

    import requests
    last_err = None
    for _ in range(2):
        _throttle(config)
        t0 = time.time()
        try:
            r = requests.post(url, json=payload, headers=headers, timeout=timeout)
            if r.status_code == 401:
                raise RuntimeError("API key 无效或已过期，请检查配置")
            r.raise_for_status()
            data = r.json()
            content = data["choices"][0]["message"]["content"]
            usage = data.get("usage") or {}
            _log(module, model, in_len, len(content or ""),
                 int((time.time() - t0) * 1000), True,
                 tokens_in=int(usage.get("prompt_tokens") or 0),
                 tokens_out=int(usage.get("completion_tokens") or 0))
            if not content or not str(content).strip():
                raise RuntimeError("大模型返回为空")
            return str(content).strip()
        except RuntimeError:
            raise
        except Exception as e:
            last_err = e
            _log(module, model, in_len, 0, int((time.time() - t0) * 1000),
                 False, err=last_err)
            time.sleep(2)
    raise RuntimeError(f"调用大模型失败: {last_err}")


def call_llm_msg(module, messages, tools=None, config=None, timeout=120,
                 temperature=0.1, max_tokens=900):
    """消息级调用（对话助手用）：支持 OpenAI 兼容 tools/function-calling，
    返回完整消息 dict：{"content": str, "tool_calls": list|None}。
    普通 call_llm() 语义不变；fake 驱动返回确定性文本（不产生工具调用，不花钱）。"""
    config = config or load_config()
    driver = (config.get("driver") or "real").lower()
    model = config.get("model") or DEFAULT_CONFIG["model"]
    if driver == "fake":
        _log(module, model, 0, 0, 0, True, driver="fake")
        return {"content": _FAKE_RESPONSES.get(module)
                or "【演示模式】未调用真实 AI（config ai_config.yaml driver: fake）。",
                "tool_calls": None}
    api_key = (config.get("api_key") or "").strip()
    if not api_key:
        raise RuntimeError("未配置 AI 接口密钥，请先在「工具·选店分析」配置 DeepSeek key")
    base = (config.get("base_url") or "").strip().rstrip("/") or DEFAULT_CONFIG["base_url"]
    payload = {"model": model, "messages": messages, "temperature": temperature,
               "max_tokens": max_tokens, "stream": False}
    if tools:
        payload["tools"] = tools
        payload["tool_choice"] = "auto"
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    import requests
    last_err = None
    for _ in range(2):
        _throttle(config)
        t0 = time.time()
        try:
            r = requests.post(base + "/chat/completions", json=payload,
                              headers=headers, timeout=timeout)
            if r.status_code == 401:
                raise RuntimeError("API key 无效或已过期，请检查配置")
            r.raise_for_status()
            data = r.json()
            msg = (data.get("choices") or [{}])[0].get("message") or {}
            usage = data.get("usage") or {}
            tcs = msg.get("tool_calls") or None
            _log(module, model, sum(len(str(m.get("content") or "")) for m in messages),
                 len(msg.get("content") or "") + len(tcs or []),
                 int((time.time() - t0) * 1000), True,
                 tokens_in=int(usage.get("prompt_tokens") or 0),
                 tokens_out=int(usage.get("completion_tokens") or 0))
            if not msg.get("content") and not tcs:
                raise RuntimeError("大模型返回为空")
            return {"content": str(msg.get("content") or ""),
                    "tool_calls": tcs}
        except RuntimeError:
            raise
        except Exception as e:
            last_err = e
            _log(module, model, 0, 0, int((time.time() - t0) * 1000), False, err=last_err)
            time.sleep(2)
    raise RuntimeError(f"调用大模型失败: {last_err}")


def call_embedding(texts, config=None, timeout=60):
    """批量文本向量。未配置 embeddings 段返回 None（调用方自行退化）；
    配置了但调用失败也返回 None（null-object：永不阻断主流程）。"""
    cfg = embeddings_config(config)
    if cfg is None or not texts:
        return None
    url = cfg["base_url"].rstrip("/") + "/embeddings"
    headers = {"Authorization": f"Bearer {cfg['api_key']}",
               "Content-Type": "application/json"}
    import requests
    t0 = time.time()
    try:
        _throttle(config or load_config())
        r = requests.post(url, json={"model": cfg["model"], "input": list(texts)},
                          headers=headers, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        vecs = [d["embedding"] for d in data["data"]]
        _log("embedding", cfg["model"], sum(len(t) for t in texts), 0,
             int((time.time() - t0) * 1000), True, tokens_in=len(texts))
        return vecs
    except Exception as e:
        _log("embedding", cfg["model"], sum(len(t) for t in texts), 0,
             int((time.time() - t0) * 1000), False, err=e)
        return None


def log_stats():
    """决策日志聚合统计（面板「AI调用统计」用）。"""
    if not os.path.exists(LOG_FILE):
        return {"calls": 0, "ok": 0, "fail": 0, "cost_est": 0.0,
                "by_module": {}, "last_err": ""}
    calls = ok = 0
    cost = 0.0
    by_module = {}
    last_err = ""
    try:
        with open(LOG_FILE, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    rec = json.loads(line)
                except (ValueError, TypeError):
                    continue
                calls += 1
                if rec.get("ok"):
                    ok += 1
                else:
                    last_err = rec.get("err") or ""
                cost += float(rec.get("cost_est") or 0)
                m = rec.get("module") or "?"
                by_module.setdefault(m, {"calls": 0, "ok": 0})
                by_module[m]["calls"] += 1
                if rec.get("ok"):
                    by_module[m]["ok"] += 1
    except Exception:
        pass
    return {"calls": calls, "ok": ok, "fail": calls - ok,
            "cost_est": round(cost, 4), "by_module": by_module,
            "last_err": last_err[:120]}


def save_config_fields(api_key=None, base_url=None, model=None, driver=None, path=None):
    """OpenAI 兼容模型可替换（M7 阶段5）：按字段更新 ai_config.yaml。
    仅改传入且非 None 的标量字段；完整保留 driver/prescreen/embeddings 等其他段。
    聊天侧禁止改配置（L4）；本函数供工具页「AI 接口」面板调用。"""
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
    for k, v in DEFAULT_CONFIG.items():
        data.setdefault(k, v)
    if api_key is not None:
        data["api_key"] = str(api_key).strip()
    if base_url is not None:
        data["base_url"] = str(base_url).strip().rstrip("/") or DEFAULT_CONFIG["base_url"]
    if model is not None:
        data["model"] = str(model).strip() or DEFAULT_CONFIG["model"]
    if driver is not None:
        d = str(driver).strip().lower()
        if d not in ("real", "fake"):
            raise ValueError("driver 只能是 real 或 fake")
        data["driver"] = d
    os.makedirs(os.path.dirname(path), exist_ok=True)
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)
    os.replace(tmp, path)
    return load_config(path)


# 常见 OpenAI 兼容供应商预设（工具页下拉；chat 不改配置）
PROVIDER_PRESETS = [
    {"name": "DeepSeek", "base_url": "https://api.deepseek.com", "model": "deepseek-chat"},
    {"name": "通义千问", "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
     "model": "qwen-plus"},
    {"name": "智谱 GLM", "base_url": "https://open.bigmodel.cn/api/paas/v4",
     "model": "glm-4-flash"},
    {"name": "Moonshot", "base_url": "https://api.moonshot.cn/v1", "model": "moonshot-v1-8k"},
    {"name": "OpenAI", "base_url": "https://api.openai.com/v1", "model": "gpt-4o-mini"},
    {"name": "自定义…", "base_url": "", "model": ""},
]


if __name__ == "__main__":
    st = log_stats()
    print(f"AI 调用统计: 共 {st['calls']} 次（成功 {st['ok']} / 失败 {st['fail']}），"
          f"估算成本 ¥{st['cost_est']}")
    for m, v in st["by_module"].items():
        print(f"  {m}: {v['calls']} 次（成功 {v['ok']}）")
    cfg = load_config()
    print(f"\n当前 driver: {cfg.get('driver')} | model: {cfg.get('model')}")
    print(f"预筛配置: {prescreen_config(cfg)}")
    print(f"embeddings: {'已配置' if embeddings_config(cfg) else '未配置（语义层退化）'}")

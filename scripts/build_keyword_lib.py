# -*- coding: utf-8 -*-
"""
种子词库提取：批量调大模型为自家高价值品组提取"点名关键词"。
输出 output/自家关键词库.xlsx（关键词/品组ID/品名/类目/来源/时间/状态）。
用法：python build_keyword_lib.py [--limit 2000] [--mock]
  --limit  提取的品组数（按月销合计降序，默认 2000）
  --mock   本地模拟（不调用真实 API，验证链路）
成本：批量合并调用（每批30品组），约 2000 组 ≈ 60-70 次调用，预计 10-25 元。
"""
import argparse
import json
import os
import sys
import time
from datetime import datetime

import pandas as pd

from ai_client import call_llm, load_config

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OWN_FILE = os.path.join(BASE, "output", "自家_归一化后.xlsx")
LIB_FILE = os.path.join(BASE, "output", "自家关键词库.xlsx")

LIB_COLS = ["关键词", "品组ID", "品名", "类目", "来源", "时间", "状态"]

SYSTEM_PROMPT = """你是闪电仓（五金建材零售）的商品专家。为每个"品组"提取能点名指向该品的"关键词"。

规则：
1. 只提取能强烈指向该品的词（如"五孔插座""螺口尖泡""燃气软管"这类点名词），一般 2-5 个；
2. 排除：颜色词（白色/金色）、材质特性词（加厚/防滑）、泛品类大词（插座/开关/灯泡/灯具这类单独出现无指向的词）、营销词（正品/包邮）、品牌词（公牛/佛山照明）；
3. 若该品没有点名词（太普通），给空数组。

只输出 JSON 对象（不要输出其他文字）：
{"品组ID1": ["词1", "词2"], "品组ID2": ["词1"], ...}"""


def _clean(v):
    if v is None:
        return ""
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def load_spu_groups(limit=2000):
    """自家_归一化后 按 SPUID 分组：品组名=组内月销最高SKU名称，月销=组内合计。"""
    df = pd.read_excel(OWN_FILE, dtype=str).fillna("")
    if "SPUID" not in df.columns or "SKUID" not in df.columns:
        raise RuntimeError("自家_归一化后.xlsx 缺 SPUID/SKUID 列")
    groups = []
    for spuid, g in df.groupby("SPUID"):
        spuid = _clean(spuid)
        if not spuid:
            continue
        name_col = "商品名称" if "商品名称" in g.columns else "清洗后商品名"
        sales_col = "月销量" if "月销量" in g.columns else "月售"
        cat_col = "类目三级_归一" if "类目三级_归一" in g.columns else "美团闪购类目三级"

        def _sales_num(v):
            try:
                return float(str(v).replace(",", "").strip())
            except (ValueError, TypeError):
                return 0.0

        g2 = g.copy()
        g2["__sales"] = g2[sales_col].apply(_sales_num) if sales_col in g2.columns else 0.0
        total_sales = float(g2["__sales"].sum())
        rep = g2.loc[g2["__sales"].idxmax()] if len(g2) else g2.iloc[0]
        groups.append({
            "品组ID": spuid,
            "品名": _clean(rep.get(name_col)),
            "类目": _clean(rep.get(cat_col)) if cat_col in g2.columns else "",
            "月销": total_sales,
            "SKU数": len(g2),
        })
    groups.sort(key=lambda x: -x["月销"])
    return groups[:limit]


def extract_batch(config, batch):
    """调大模型提取一批品组的关键词。返回 {品组ID: [词]}。（已收口 ai_client）"""
    msgs = [{"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": "品组清单：\n" + "\n".join(
                f"{b['品组ID']}|{b['品名']}|{b['类目']}" for b in batch)}]
    try:
        content = call_llm("keyword_lib", msgs, config=config, timeout=120,
                           temperature=0.2, max_tokens=800, json_mode=True)
    except RuntimeError as e:
        print(f"  批次调用失败：{e}")
        return {}
    s, e = content.find("{"), content.rfind("}")
    if 0 <= s < e:
        try:
            return json.loads(content[s:e + 1])
        except (ValueError, TypeError):
            return {}
    return {}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--limit", type=int, default=2000)
    ap.add_argument("--mock", action="store_true")
    args = ap.parse_args()

    print(f"加载自家品组（Top {args.limit} 按月销）...")
    groups = load_spu_groups(args.limit)
    print(f"  品组 {len(groups)} 个")

    rows = []
    if args.mock:
        for g in groups[:30]:
            name = g["品名"]
            segs = [s for s in __import__("re").findall(r"[\u4e00-\u9fff]{2,}", str(name)) if len(s) >= 2][:3]
            for w in segs:
                rows.append({"关键词": w, "品组ID": g["品组ID"], "品名": g["品名"],
                             "类目": g["类目"], "来源": "大模型提取(mock)", "时间": "", "状态": "启用"})
        print("（mock 模式：不调用真实 API）")
    else:
        config = load_config()
        driver = str(config.get("driver") or "real").lower()
        if driver != "fake" and not (config.get("api_key") or "").strip():
            print("未配置 DeepSeek key，退出（可用 --mock 验证链路，"
                  "或在 ai_config.yaml 设 driver: fake 演示模式）")
            sys.exit(1)
        BATCH = 30
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i in range(0, len(groups), BATCH):
            batch = groups[i:i + BATCH]
            got = extract_batch(config, batch)
            for g in batch:
                for w in got.get(g["品组ID"], []) or []:
                    w = _clean(w)
                    if w and len(w) >= 2:
                        rows.append({"关键词": w, "品组ID": g["品组ID"], "品名": g["品名"],
                                     "类目": g["类目"], "来源": "大模型提取",
                                     "时间": now, "状态": "启用"})
            print(f"  批 {i // BATCH + 1}/{(len(groups) + BATCH - 1) // BATCH}: "
                  f"{len(batch)} 组 -> 词 {sum(1 for r in rows if r['来源'] == '大模型提取')}")
            time.sleep(0.5)

    # 去重（同词同品组只保留一条）
    seen = set()
    uniq = []
    for r in rows:
        k = (r["关键词"], r["品组ID"])
        if k in seen:
            continue
        seen.add(k)
        uniq.append(r)

    df = pd.DataFrame(uniq, columns=LIB_COLS)
    df.to_excel(LIB_FILE, index=False)
    n_groups = df["品组ID"].nunique() if len(df) else 0
    print(f"\n词库已生成：{LIB_FILE}")
    print(f"  词条 {len(df)} 条 | 覆盖品组 {n_groups} 个 | 去重后平均每品组 {len(df) / max(n_groups, 1):.1f} 词")
    print("\n抽样 10 条（供抽查）：")
    print(df.head(10).to_string(index=False))
    print("\n提示：抽查发现不合适的词，可在 Excel 里把该行 状态 改为「忽略」；")
    print("确认竞品对应时会自动把真实词对补进词库（来源=确认反哺）。")


if __name__ == "__main__":
    main()

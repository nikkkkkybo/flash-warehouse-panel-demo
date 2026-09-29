# -*- coding: utf-8 -*-
"""
训练样本导出器（P2.4 第一步：黄金集 + 微调样本积累）。

正样本 = 「人工确认过」的对应关系（这些就是真值，也作向量层召回率回测的黄金集）：
  ① progress.json confirmed（人工 SKU 级确认，最高质量）
  ② 两层汇总库人工行（人工确认的品级）
  ③ 跨店绑定组（人工绑定的跨店同品，来自 候选池.json）
难负样本（长得像但不是——教模型分辨 M8/M10、沉头/蘑菇头）：
  ① progress no_match 竞品 与 同店同品类目下 TF-IDF 最近邻自家品（被人工否掉的最强干扰对）
  ② 探测器⑤撤销类（无对应备注含"错"的）——暂并入①
输出 output/训练样本库.xlsx（正样本 / 难负样本 两 sheet，带来源列）。
用法：python scripts/build_match_dataset.py
"""
import io
import json
import os
import sys

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT = os.path.join(BASE, "output")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def _clean(v):
    s = str(v or "").strip()
    return "" if s.lower() in ("nan", "none") else s


def collect_positives():
    """正样本：[(店, 竞品名, 自家名, 来源)]。自家名经 SKUID→商品名称 解析。"""
    import pandas as pd
    rows = []
    own_f = os.path.join(OUT, "自家_归一化后.xlsx")
    own_names = {}
    if os.path.exists(own_f):
        own = pd.read_excel(own_f, dtype=str)
        own = own.drop_duplicates(subset=["SKUID"])
        name_col = "商品名称" if "商品名称" in own.columns else "name"
        own_names = {_clean(r.get("SKUID")): _clean(r.get(name_col))
                     for _, r in own.iterrows() if _clean(r.get("SKUID"))}
    prog_path = os.path.join(BASE, "scripts", "progress.json")
    if os.path.exists(prog_path):
        prog = json.load(open(prog_path, encoding="utf-8"))
        confirmed = prog.get("confirmed", {})
        # 竞品名查找：各店归一表
        names = {}
        for d in sorted(os.listdir(OUT)):
            if not d.startswith("试跑_"):
                continue
            shop = d[len("试跑_"):]
            f = os.path.join(OUT, d, "竞品_归一化后.xlsx")
            if not os.path.exists(f):
                continue
            import pandas as pd
            df = pd.read_excel(f, dtype=str)
            sku_col = "sku_id" if "sku_id" in df.columns else "Sku_id"
            name_col = "name" if "name" in df.columns else "商品名字"
            for _, r in df.iterrows():
                sid = _clean(r.get(sku_col))
                if sid:
                    names[(shop, sid)] = _clean(r.get(name_col))
        seen = set()
        for shop, m in prog.get("confirmed", {}).items():
            for csku, v in (m or {}).items():
                osku = v.get("sku_id") if isinstance(v, dict) else v
                osku = str(osku or "").strip()
                cn = names.get((shop, str(csku)), "")
                on = own_names.get(osku, "")
                if cn and on and (shop, csku) not in seen:
                    seen.add((shop, csku))
                    rows.append((shop, cn, on, "人工确认(SPU级)", osku))
    # 两层汇总库人工行（品名对 + 真实自家SPUID）
    lib_path = os.path.join(OUT, "对齐汇总库.xlsx")
    if os.path.exists(lib_path):
        import pandas as pd
        try:
            spu = pd.read_excel(lib_path, sheet_name="SPU匹配层", dtype=str)
            for _, r in spu.iterrows():
                cn, on = _clean(r.get("竞品名称") or r.get("店")), _clean(r.get("自家品名"))
                cn2 = _clean(r.get("竞品品名"))
                ogid = _clean(r.get("自家品组ID"))
                if cn2 and on and str(r.get("来源") or "").strip() == "人工" and ogid:
                    rows.append((cn, cn2, on, "人工确认(SPU层)", ogid))
        except Exception as e:
            print("  汇总库读取失败（跳过）:", e)
    return rows


def collect_hard_negatives():
    """难负：无对应竞品 ↔ 同店同品类目 TF-IDF 最近邻自家品（最强干扰对）。
    这些"长得像但被人工否掉"的对，是微调里最值钱的负样本。"""
    import pandas as pd
    out = []
    own_f = os.path.join(OUT, "自家_归一化后.xlsx")
    if not os.path.exists(own_f):
        return out
    own = pd.read_excel(own_f, dtype=str)
    own = own.drop_duplicates(subset=["SKUID"])
    own = own[own["商品名称"].notna() & (own["商品名称"].astype(str).str.strip() != "")]
    prog_path = os.path.join(BASE, "scripts", "progress.json")
    if not os.path.exists(prog_path):
        return out
    prog = json.load(open(prog_path, encoding="utf-8"))
    from sklearn.feature_extraction.text import TfidfVectorizer
    from sklearn.neighbors import NearestNeighbors
    own_names = own["商品名称"].astype(str).tolist()
    own_cats = own["类目三级_归一"].astype(str).tolist() \
        if "类目三级_归一" in own.columns else [""] * len(own_names)
    try:
        vec = TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 3), min_df=1)
        M = vec.fit_transform(own_names)
        nn = NearestNeighbors(n_neighbors=1, metric="cosine").fit(M)
    except Exception as e:
        print("  TF-IDF 构建失败（跳过难负）:", e)
        return out
    for shop, m in (prog.get("no_match") or {}).items():
        comp_f = os.path.join(OUT, f"试跑_{shop}", "竞品_归一化后.xlsx")
        if not os.path.exists(comp_f) or not m:
            continue
        comp = pd.read_excel(comp_f, dtype=str)
        sku_col = "sku_id" if "sku_id" in comp.columns else "Sku_id"
        name_col = "name" if "name" in comp.columns else "商品名字"
        cat_col = "类目三级_归一" if "类目三级_归一" in comp.columns else "category3"
        rows = []
        for _, r in comp.iterrows():
            sid = _clean(r.get(sku_col))
            if sid and sid in (m or {}):
                rows.append((sid, _clean(r.get(name_col)), _clean(r.get(cat_col))))
        if not rows:
            continue
        q = vec.transform([c[1] for c in rows])
        _, idx = nn.kneighbors(q)
        for (sid, cn, ccat), j in zip(rows, idx.flatten()):
            on = own_names[j]
            if not cn or not on:
                continue
            if ccat and own_cats[j] and ccat != own_cats[j]:
                continue          # 跨类目的近邻是"易负"，不作难负
            out.append((shop, cn, on, "人工无对应", "同店同品类目最近邻仍被否"))
    return out


def main():
    pos = collect_positives()
    neg = collect_hard_negatives()
    print(f"正样本 {len(pos)} 条 | 难负样本 {len(neg)} 条")
    import pandas as pd
    dfp = pd.DataFrame(pos, columns=["店", "竞品名", "自家名", "来源", "自家ID"])
    dfn = pd.DataFrame(neg, columns=["店", "竞品名", "自家名", "来源", "说明"])
    path = os.path.join(OUT, "训练样本库.xlsx")
    with pd.ExcelWriter(path, engine="openpyxl") as xw:
        dfp.to_excel(xw, sheet_name="正样本", index=False)
        dfn.to_excel(xw, sheet_name="难负样本", index=False)
    print("->", path)


if __name__ == "__main__":
    main()

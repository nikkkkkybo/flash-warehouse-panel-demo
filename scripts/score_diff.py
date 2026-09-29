# -*- coding: utf-8 -*-
"""
闪电仓选品工具 - 阶段4：跨店差异打分模块

对已对齐的匹配组做 7 维度差异分析：
  4.1 价格差异      竞品到手价(神价>折后价>原价) vs 自家售价(美团闪购渠道售价)
  4.2 毛利测算      自家独有：当前毛利/毛利率 + 对标毛利/毛利率(若降到竞品价)
  4.3 图因素        仅判有无，缺图风险标记
  4.4 标题/规格/活动/评价
  4.5 汇总表导出 + 综合建议动作

输入:
  output/对齐_汇总.xlsx      30 行已确认匹配（含 竞品品ID/竞品Sku_id/自家SKUID ... ）
  output/竞品_归一化后.xlsx  竞品全表（Sku_id 唯一，商品ID 非唯一）
  output/自家_归一化后.xlsx  自家全表（SKUID 有重复，去重取首行）
输出:
  output/差异打分_汇总.xlsx

关联说明（注意：与任务描述的字段名差异）：
  对齐_汇总 实际列名为「竞品品ID」（非"竞品商品ID"），另有精确的「竞品Sku_id」。
  竞品表 商品ID 是 SPU 级、非唯一（一个商品ID平均 1.61 个 SKU），直接按 商品ID 关联会
  把同 SPU 下其它 SKU 的价格/月售错配到本匹配行。因此本脚本采用更精确的关联键：
      竞品侧：对齐_汇总.竞品Sku_id  -> 竞品_归一化后.Sku_id   （Sku_id 唯一）
      自家侧：对齐_汇总.自家SKUID    -> 自家_归一化后.SKUID    （按任务要求；先去重）
  竞品品ID 仍原样带出到输出列「竞品商品ID」，满足汇总表展示需要。
"""
import os
import sys
from pathlib import Path

import pandas as pd

# 复用阶段3 的月售解析器
sys.path.insert(0, str(Path(__file__).resolve().parent))
from hot_products import parse_month_sales  # noqa: E402


# ============================================================
# 路径常量
# ============================================================
ROOT = Path(__file__).resolve().parent.parent          # <项目根目录>
OUTPUT_DIR = ROOT / "output"
ALIGN_FILE = OUTPUT_DIR / "对齐_汇总.xlsx"
COMP_FILE = OUTPUT_DIR / "竞品_归一化后.xlsx"
OWN_FILE = OUTPUT_DIR / "自家_归一化后.xlsx"
OUT_FILE = OUTPUT_DIR / "差异打分_汇总.xlsx"

# 关键词维度（标题覆盖用）
TITLE_KW_FIELDS = ["品牌", "型号", "功率", "容量", "尺寸"]


# ============================================================
# 工具函数
# ============================================================
def _clean(v):
    """统一清洗：None/NaN/'nan'/'None' -> ''，其余 strip 后转 str。"""
    if v is None:
        return ""
    if isinstance(v, float):
        return "" if pd.isna(v) else str(v).strip()
    s = str(v).strip()
    return "" if s.lower() in ("nan", "none") else s


def _num(v):
    """转 float；空/无法解析返回 None。"""
    if v is None:
        return None
    try:
        if isinstance(v, str):
            s = v.strip()
            if s == "" or s.lower() in ("nan", "none"):
                return None
            return float(s)
        f = float(v)
        return None if pd.isna(f) else f
    except (ValueError, TypeError):
        return None


def _fmt_rate(x, nd=4):
    """比率格式化：None -> None，否则 round。"""
    if x is None:
        return None
    return round(float(x), nd)


# ============================================================
# 数据加载
# ============================================================
def load_data():
    """读取三张表，ID 列强制 str；自家表按 SKUID 去重取首行。"""
    if not ALIGN_FILE.exists():
        print(f"[错误] 对齐_汇总不存在: {ALIGN_FILE}", file=sys.stderr)
        sys.exit(1)
    if not COMP_FILE.exists():
        print(f"[错误] 竞品表不存在: {COMP_FILE}", file=sys.stderr)
        sys.exit(1)
    if not OWN_FILE.exists():
        print(f"[错误] 自家表不存在: {OWN_FILE}", file=sys.stderr)
        sys.exit(1)

    da = pd.read_excel(
        ALIGN_FILE,
        dtype={"竞品品ID": str, "竞品Sku_id": str,
               "自家SPUID": str, "自家SKUID": str},
    )
    dfc = pd.read_excel(
        COMP_FILE,
        dtype={"商品ID": str, "Sku_id": str, "条形码": str},
    )
    dfo = pd.read_excel(
        OWN_FILE,
        dtype={"SPUID": str, "SKUID": str, "商品条码": str},
    )

    # 空值统一
    for df in (da, dfc, dfo):
        for c in df.columns:
            if df[c].dtype == object:
                df[c] = df[c].fillna("")

    # 自家 SKUID 有重复（实测 384 行重复），去重取首行，保证 1:1 关联
    own_dup = int(dfo["SKUID"].duplicated().sum())
    dfo = dfo.drop_duplicates("SKUID", keep="first").reset_index(drop=True)

    # 竞品 Sku_id 唯一性校验
    comp_dup = int(dfc["Sku_id"].duplicated().sum())

    print(f"[加载] 对齐_汇总: {da.shape}  竞品: {dfc.shape}  自家(去重后): {dfo.shape}")
    print(f"      竞品 Sku_id 重复: {comp_dup}  自家 SKUID 去重前重复: {own_dup}")
    return da, dfc, dfo


# ============================================================
# 预处理：规格组合统计（同品牌+同型号 的 SKU 数）
# ============================================================
def build_spec_counts(dfc, dfo):
    """返回 (comp_counts, own_counts)：(品牌, 型号) -> SKU 数。

    型号为空时并入 (品牌, "") 桶；品牌为空则不计（无法定义同品）。
    """
    def _key(row):
        b = _clean(row.get("品牌（抽取）"))
        m = _clean(row.get("型号"))
        return (b, m)

    comp_counts = {}
    for _, r in dfc.iterrows():
        k = _key(r)
        if k[0] == "":
            continue
        comp_counts[k] = comp_counts.get(k, 0) + 1

    own_counts = {}
    for _, r in dfo.iterrows():
        k = _key(r)
        if k[0] == "":
            continue
        own_counts[k] = own_counts.get(k, 0) + 1
    return comp_counts, own_counts


# ============================================================
# Task 4.1 价格差异
# ============================================================
def score_price(crow, orow):
    """返回 dict: 到手价/售价/价差/价差率/竞争力 + 来源标记。"""
    shen = _num(crow.get("神价"))
    zhe = _num(crow.get("折后价"))
    yuan = _num(crow.get("原价"))

    if shen is not None:
        comp_price = shen
        comp_src = "神价"
    elif zhe is not None:
        comp_price = zhe
        comp_src = "折后价"
    else:
        comp_price = yuan
        comp_src = "原价"

    own_price = _num(orow.get("美团闪购渠道售价"))

    out = {
        "竞品到手价": comp_price,
        "竞品到手价来源": comp_src,
        "自家售价": own_price,
        "价格_价差": None,
        "价格_价差率": None,
        "价格_竞争力": "N/A",
    }
    if comp_price is None or own_price is None:
        return out

    diff = own_price - comp_price
    rate = diff / comp_price if comp_price != 0 else None
    if abs(diff) < 0.01:
        comp = "持平"
    elif diff < 0:
        comp = "优势"      # 自家更低
    else:
        comp = "劣势"      # 自家更高
    out.update({
        "价格_价差": round(diff, 4),
        "价格_价差率": _fmt_rate(rate),
        "价格_竞争力": comp,
    })
    return out


# ============================================================
# Task 4.2 毛利测算（自家独有，仅当采购价非空）
# ============================================================
def score_margin(crow, orow, comp_price, own_price):
    """返回 dict: 当前毛利/毛利率/对标毛利/对标毛利率/毛利变化/可算标记。"""
    purchase = _num(orow.get("采购价"))
    out = {
        "自家采购价": purchase,
        "毛利_可算": "否",
        "毛利_当前毛利": None,
        "毛利_当前毛利率": None,
        "毛利_对标毛利": None,
        "毛利_对标毛利率": None,
        "毛利_毛利变化": None,
    }
    if purchase is None or own_price is None or comp_price is None:
        return out
    out["毛利_可算"] = "是"

    cur_gp = own_price - purchase
    cur_rate = cur_gp / own_price if own_price != 0 else None
    bench_gp = comp_price - purchase          # 若自家降到竞品到手价
    bench_rate = bench_gp / comp_price if comp_price != 0 else None
    delta = bench_gp - cur_gp
    out.update({
        "毛利_当前毛利": round(cur_gp, 4),
        "毛利_当前毛利率": _fmt_rate(cur_rate),
        "毛利_对标毛利": round(bench_gp, 4),
        "毛利_对标毛利率": _fmt_rate(bench_rate),
        "毛利_毛利变化": round(delta, 4),
    })
    return out


# ============================================================
# Task 4.3 图因素（仅判有无）
# ============================================================
def score_image(crow, orow):
    own_has = _clean(orow.get("是否有图片")) == "是"
    comp_has = _clean(crow.get("图片url")) != ""

    if own_has and comp_has:
        status = "双方有图，内容对比需人工"
    elif (not own_has) and comp_has:
        status = "缺图风险，建议优先补图"
    elif own_has and (not comp_has):
        status = "自家有图，竞品无图"
    else:
        status = "双方均无图"
    return {
        "图_竞品有图": "是" if comp_has else "否",
        "图_自家有图": "是" if own_has else "否",
        "图_状态": status,
    }


# ============================================================
# Task 4.4a 标题覆盖度
# ============================================================
def _title_coverage(title, brand, model, power, capacity, size):
    """计算标题对 5 个关键词维度的覆盖数（0-5）。

    每个维度：值非空且（小写后）出现在标题中即 +1。空值维度算未覆盖。
    返回 (覆盖数, 总维度=5)。
    """
    t = _clean(title).lower()
    cnt = 0
    for v in (brand, model, power, capacity, size):
        vs = _clean(v)
        if not vs:
            continue
        if vs.lower() in t:
            cnt += 1
    return cnt, 5


def score_title(crow, orow):
    c_cov, _ = _title_coverage(
        crow.get("商品名字"), crow.get("品牌（抽取）"), crow.get("型号"),
        crow.get("功率"), crow.get("容量"), crow.get("尺寸"),
    )
    o_cov, _ = _title_coverage(
        orow.get("商品名称"), orow.get("品牌（抽取）"), orow.get("型号"),
        orow.get("功率"), orow.get("容量"), orow.get("尺寸"),
    )
    return {
        "标题_竞品覆盖数": c_cov,
        "标题_自家覆盖数": o_cov,
        "标题_自家覆盖度": f"{o_cov}/5",
        "标题_覆盖差距": o_cov - c_cov,   # 正=自家更全，负=竞品更全
    }


# ============================================================
# Task 4.4b 规格组合（同品牌同型号 SKU 数对比）
# ============================================================
def score_spec(crow, orow, comp_counts, own_counts):
    cb = _clean(crow.get("品牌（抽取）"))
    cm = _clean(crow.get("型号"))
    ob = _clean(orow.get("品牌（抽取）"))
    om = _clean(orow.get("型号"))

    c_n = comp_counts.get((cb, cm), 0) if cb else 0
    o_n = own_counts.get((ob, om), 0) if ob else 0
    return {
        "规格_竞品SKU数": c_n,
        "规格_自家SKU数": o_n,
        "规格_档位差": c_n - o_n,   # 正=竞品规格更丰富
        "规格_型号缺失": "是" if (cm == "" or om == "") else "否",
    }


# ============================================================
# Task 4.4c 活动（竞品神价/折扣有无；自家无活动数据）
# ============================================================
def score_activity(crow):
    shen = _num(crow.get("神价"))
    zhe = _clean(crow.get("折扣"))
    has = "有" if shen is not None else "无"
    return {
        "活动_竞品有无": has,
        "活动_竞品神价": shen,
        "活动_竞品折扣": zhe,
        "活动_自家": "自家无活动数据",
    }


# ============================================================
# Task 4.4d 评价（月售/月销量）
# ============================================================
def score_review(crow, orow):
    raw = _clean(crow.get("月售"))
    parsed = parse_month_sales(raw)
    c_val = parsed["数值"]
    o_val = _num(orow.get("月销量"))
    o_missing = o_val is None
    diff = None if o_missing else round(o_val - c_val, 4)
    return {
        "评价_竞品月售原值": raw,
        "评价_竞品月售类型": parsed["类型"],
        "评价_竞品月售值": c_val,
        "评价_自家月销量": o_val,
        "评价_自家月销量缺失": "是" if o_missing else "否",
        "评价_销量差": diff,    # 自家 - 竞品；None=自家缺失
    }


# ============================================================
# Task 4.5 综合建议动作
# ============================================================
def make_suggestion(price_res, img_res, spec_res, title_res):
    """基于各维度初判，按优先级给出动作。

    优先级：调价 > 补图 > 扩规格 > 优化标题；无则 保持。
    """
    actions = []
    if price_res["价格_竞争力"] == "劣势":
        actions.append("调价")
    if "缺图风险" in img_res["图_状态"]:
        actions.append("补图")
    if spec_res["规格_档位差"] > 0:
        actions.append("扩规格")
    if title_res["标题_覆盖差距"] < 0:   # 自家标题覆盖少于竞品
        actions.append("优化标题")
    return "；".join(actions) if actions else "保持"


# ============================================================
# 主流程
# ============================================================
def main():
    print("=" * 70)
    print("阶段4 跨店差异打分")
    print("=" * 70)

    da, dfc, dfo = load_data()

    # 命中校验
    sid_hit = da["竞品Sku_id"].isin(dfc["Sku_id"]).sum()
    oid_hit = da["自家SKUID"].isin(dfo["SKUID"]).sum()
    print(f"[关联校验] 竞品Sku_id 命中 {sid_hit}/{len(da)}  "
          f"自家SKUID 命中 {oid_hit}/{len(da)}")

    # 建索引（精确 1:1）
    comp_idx = dfc.set_index("Sku_id")
    own_idx = dfo.set_index("SKUID")
    comp_counts, own_counts = build_spec_counts(dfc, dfo)

    rows = []
    miss_comp = miss_own = 0
    for i, a in da.iterrows():
        gid = f"MG{i + 1:04d}"
        sid = _clean(a.get("竞品Sku_id"))
        oid = _clean(a.get("自家SKUID"))

        # 关联查全字段
        if sid and sid in comp_idx.index:
            crow = comp_idx.loc[sid]
            if isinstance(crow, pd.DataFrame):
                crow = crow.iloc[0]
        else:
            crow = pd.Series(dtype=object)
            miss_comp += 1
        if oid and oid in own_idx.index:
            orow = own_idx.loc[oid]
            if isinstance(orow, pd.DataFrame):
                orow = orow.iloc[0]
        else:
            orow = pd.Series(dtype=object)
            miss_own += 1

        # 7 维度打分
        price = score_price(crow, orow)
        margin = score_margin(crow, orow, price["竞品到手价"], price["自家售价"])
        image = score_image(crow, orow)
        title = score_title(crow, orow)
        spec = score_spec(crow, orow, comp_counts, own_counts)
        activity = score_activity(crow)
        review = score_review(crow, orow)
        suggestion = make_suggestion(price, image, spec, title)

        row = {
            "匹配组ID": gid,
            "竞品商品ID": _clean(a.get("竞品品ID")),
            "竞品商品名": _clean(a.get("竞品商品名")),
            "自家SKUID": oid,
            "自家商品名": _clean(a.get("自家商品名")),
            "匹配层": a.get("匹配层"),
            "置信度": a.get("置信度"),
            # --- 7 维度汇总列 ---
            "价格_竞争力": price["价格_竞争力"],
            "价格_价差率": price["价格_价差率"],
            "毛利_当前毛利率": margin["毛利_当前毛利率"],
            "毛利_对标毛利率": margin["毛利_对标毛利率"],
            "图_状态": image["图_状态"],
            "标题_自家覆盖度": title["标题_自家覆盖度"],
            "规格_档位差": spec["规格_档位差"],
            "活动_竞品有无": activity["活动_竞品有无"],
            "评价_销量差": review["评价_销量差"],
            "综合建议动作": suggestion,
        }
        # --- 明细列（便于抽检/溯源）---
        detail = {
            "竞品Sku_id": sid,
            "竞品到手价": price["竞品到手价"],
            "竞品到手价来源": price["竞品到手价来源"],
            "自家售价": price["自家售价"],
            "价格_价差": price["价格_价差"],
            "自家采购价": margin["自家采购价"],
            "毛利_可算": margin["毛利_可算"],
            "毛利_当前毛利": margin["毛利_当前毛利"],
            "毛利_对标毛利": margin["毛利_对标毛利"],
            "毛利_毛利变化": margin["毛利_毛利变化"],
            "图_竞品有图": image["图_竞品有图"],
            "图_自家有图": image["图_自家有图"],
            "标题_竞品覆盖数": title["标题_竞品覆盖数"],
            "标题_自家覆盖数": title["标题_自家覆盖数"],
            "标题_覆盖差距": title["标题_覆盖差距"],
            "规格_竞品SKU数": spec["规格_竞品SKU数"],
            "规格_自家SKU数": spec["规格_自家SKU数"],
            "规格_型号缺失": spec["规格_型号缺失"],
            "活动_竞品神价": activity["活动_竞品神价"],
            "活动_竞品折扣": activity["活动_竞品折扣"],
            "评价_竞品月售原值": review["评价_竞品月售原值"],
            "评价_竞品月售类型": review["评价_竞品月售类型"],
            "评价_竞品月售值": review["评价_竞品月售值"],
            "评价_自家月销量": review["评价_自家月销量"],
            "评价_自家月销量缺失": review["评价_自家月销量缺失"],
        }
        row.update(detail)
        rows.append(row)

    out = pd.DataFrame(rows)
    print(f"[关联缺失] 竞品未命中 {miss_comp}  自家未命中 {miss_own}")

    # 导出
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    out.to_excel(OUT_FILE, index=False)
    print(f"\n[导出] {OUT_FILE}  共 {len(out)} 行 {out.shape[1]} 列")

    # ---- 统计 ----
    print("\n" + "=" * 70)
    print("打分统计")
    print("=" * 70)
    print(f"价格竞争力: {out['价格_竞争力'].value_counts().to_dict()}")
    print(f"图状态: {out['图_状态'].value_counts().to_dict()}")
    print(f"活动有无: {out['活动_竞品有无'].value_counts().to_dict()}")
    print(f"毛利可算: {out['毛利_可算'].value_counts().to_dict()}")
    print(f"自家月销量缺失: {out['评价_自家月销量缺失'].value_counts().to_dict()}")
    print(f"综合建议动作: {out['综合建议动作'].value_counts().to_dict()}")

    # ---- 抽检 3 条 ----
    print("\n" + "=" * 70)
    print("抽检（前 3 条匹配记录）")
    print("=" * 70)
    for i in range(min(3, len(out))):
        r = out.iloc[i]
        print(f"\n[{i+1}] 匹配组={r['匹配组ID']} 匹配层={r['匹配层']} 置信={r['置信度']}")
        print(f"    竞品: {r['竞品商品名'][:32]}  (Sku_id={r['竞品Sku_id']})")
        print(f"    自家: {r['自家商品名'][:32]}  (SKUID={r['自家SKUID']})")
        print(f"    价格: 到手价={r['竞品到手价']}({r['竞品到手价来源']}) "
              f"自家售价={r['自家售价']} 价差={r['价格_价差']} "
              f"价差率={r['价格_价差率']} 竞争力={r['价格_竞争力']}")
        if r["毛利_可算"] == "是":
            print(f"    毛利: 采购价={r['自家采购价']} 当前毛利={r['毛利_当前毛利']}"
                  f"(率={r['毛利_当前毛利率']}) 对标毛利={r['毛利_对标毛利']}"
                  f"(率={r['毛利_对标毛利率']}) 变化={r['毛利_毛利变化']}")
        else:
            print(f"    毛利: 不可算(采购价={r['自家采购价']})")
        print(f"    图: {r['图_状态']}")
        print(f"    标题: 竞品覆盖={r['标题_竞品覆盖数']}/5 自家覆盖={r['标题_自家覆盖数']}/5 "
              f"差距={r['标题_覆盖差距']}")
        print(f"    规格: 竞品SKU={r['规格_竞品SKU数']} 自家SKU={r['规格_自家SKU数']} "
              f"档位差={r['规格_档位差']}")
        print(f"    活动: 竞品有无={r['活动_竞品有无']} 神价={r['活动_竞品神价']} 折扣={r['活动_竞品折扣']}")
        print(f"    评价: 竞品月售={r['评价_竞品月售原值']}(值={r['评价_竞品月售值']}) "
              f"自家月销量={r['评价_自家月销量']}(缺失={r['评价_自家月销量缺失']}) "
              f"销量差={r['评价_销量差']}")
        print(f"    => 综合建议: {r['综合建议动作']}")

    print("\n完成。")


if __name__ == "__main__":
    main()

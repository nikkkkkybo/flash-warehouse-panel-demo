# -*- coding: utf-8 -*-
"""
生成「数据准备清单」Excel
输出: d:\\查图\\output\\数据准备清单.xlsx

包含 3 个 sheet:
  1. 数据准备清单（主表，4 份数据 ①②③④）
  2. 执行编排图（阶段 B/C2/D2/E/F1/F2/F3 依赖与并行关系）
  3. 回归验证矩阵（各阶段验证项与失败处理）

依据:
  - 自家商品库_标准版本.xlsx 列结构（SPUID/SKUID/图片url 等 19 列）
  - output/类目映射_待人工确认.xlsx 列结构（三级类目/人工确认映射 等 7 列，共 135 行）
  - scripts/brand_extractor_v2.py 的 BRAND_DICT（set 形式）与黑名单
  - scripts/import_adapter.py 的 STANDARD_FIELDS/KEYWORDS 猜映射
  - scripts/normalize.py 的 COMP 列映射（商品名字/规格/美团后台三级类）
"""
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = r"<项目根目录>"
OUT_DIR = os.path.join(BASE, "output")
OUT_PATH = os.path.join(OUT_DIR, "数据准备清单.xlsx")

# ===========================================================================
# 样式定义
# ===========================================================================
HEADER_FILL = PatternFill("solid", fgColor="305496")
HEADER_FONT = Font(name="微软雅黑", size=11, bold=True, color="FFFFFF")
CELL_FONT = Font(name="微软雅黑", size=10)
NOTE_FONT = Font(name="微软雅黑", size=10, color="C00000", bold=True)
TITLE_FONT = Font(name="微软雅黑", size=14, bold=True, color="1F3864")
WRAP = Alignment(wrap_text=True, vertical="top", horizontal="left")
CENTER_WRAP = Alignment(wrap_text=True, vertical="center", horizontal="center")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# 解锁阶段配色
STAGE_FILLS = {
    "B":  PatternFill("solid", fgColor="FFF2CC"),   # 浅黄
    "C2": PatternFill("solid", fgColor="E2EFDA"),   # 浅绿
    "D2": PatternFill("solid", fgColor="DDEBF7"),   # 浅蓝
    "E":  PatternFill("solid", fgColor="FCE4D6"),   # 浅橙
    "F1": PatternFill("solid", fgColor="EDEDED"),
    "F2": PatternFill("solid", fgColor="EDEDED"),
    "F3": PatternFill("solid", fgColor="EDEDED"),
    "全链路": PatternFill("solid", fgColor="F2F2F2"),
}


def style_header_row(ws, n_cols, row=1):
    for c in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = CENTER_WRAP
        cell.border = BORDER
    ws.row_dimensions[row].height = 28


def style_body(ws, n_rows, n_cols, start_row=2):
    for r in range(start_row, start_row + n_rows):
        for c in range(1, n_cols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = CELL_FONT
            cell.alignment = WRAP
            cell.border = BORDER


def set_widths(ws, widths):
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w


def paint_stage_col(ws, col_idx, stage_values, start_row=2):
    """按阶段值给指定列上色"""
    for i, st in enumerate(stage_values):
        fill = STAGE_FILLS.get(st)
        if fill:
            ws.cell(row=start_row + i, column=col_idx).fill = fill


# ===========================================================================
# Sheet 1: 数据准备清单（主表）
# ===========================================================================
HEADERS = [
    "数据编号", "数据名称", "用途说明", "解锁阶段", "必填字段", "可选字段",
    "格式要求", "文件命名建议", "示例（1-2行）", "提交位置", "当前状态",
]

ROWS = [
    # ① 美团后台带图url商品表
    [
        "①",
        "美团后台带图url商品表",
        "补全「自家商品库_标准版本.xlsx」的 图片url 列（当前该列全为空），"
        "解锁第4层「文本+图片双信号融合」匹配，提升对齐命中率。",
        "B",
        "SPUID（或 SPU id）；SKUID（或 SKU id）；图片url（http 开头的图片地址）",
        "商品名称；类目",
        "xlsx；UTF-8。\n"
        "⚠ SPUID/SKUID 必须保留 19 位精度：用「文本格式」或前置单引号 "
        "（如 '2038460756227977255），切勿用 Excel 默认数字格式，否则会被 float64 截断丢精度。",
        "自家商品库_带图url.xlsx",
        "SPUID=2038460756227977255 | SKUID=2085273863046144013 | 图片url=https://img.meituan.net/.../g12.jpg\n"
        "SPUID=2038460756227977256 | SKUID=2085273863055228929 | 图片url=https://img.meituan.net/.../g34.jpg",
        "d:\\查图\\",
        "待提供",
    ],
    # ② 品牌清单
    [
        "②",
        "品牌清单",
        "扩充品牌词典（brand_extractor_v2.py 的 BRAND_DICT），提升品牌抽取命中率（当前约 34%），"
        "直接改善第3层匹配与 SPU 聚合质量。",
        "C2",
        "品牌名称",
        "来源（营业执照/品牌授权/高频词确认）；英文名；备注",
        "xlsx 或 txt（每行一个品牌名）。\n"
        "⚠ 不要包含 PE/PVC/LED/开关/插座 等材质代号/品类词/营销词（已在 ALL_BLACKLIST 中，会被剔除）。",
        "品牌清单_用户补充.xlsx",
        "品牌名称=红典 | 来源=高频词确认 | 备注=出现36次\n"
        "品牌名称=路易华登 | 来源=营业执照 | 英文名=LUYIHUADENG",
        "d:\\查图\\",
        "待提供",
    ],
    # ③ 类目映射确认结果
    [
        "③",
        "类目映射确认结果",
        "写回类目映射表，统一两表三级类目口径，提升第2层「强组合键」匹配数，"
        "并避免第3层因类目错位而下沉。",
        "D2",
        "原类目名（即待人工确认表中的「三级类目」列）；"
        "标准类目（填入「人工确认映射」列，即用户确认的统一类目名）",
        "备注",
        "在 d:\\查图\\output\\类目映射_待人工确认.xlsx 基础上填回「人工确认映射」列后另存。\n"
        "⚠ 共 135 项需全部填写（106 项仅自家 + 29 项仅竞品）；无对应填「无对应」。",
        "类目映射_确认结果.xlsx（基于原文件另存）",
        "三级类目=开关插座 | 人工确认映射=墙壁开关 | 备注=统一为墙壁开关\n"
        "三级类目=暖贴/暖宝宝 | 人工确认映射=无对应 | 备注=竞品独家类目",
        "d:\\查图\\output\\",
        "待提供",
    ],
    # ④ 第二家竞店Excel表
    [
        "④",
        "第二家竞店Excel表",
        "验证多竞店对齐能力（一个本地 SKU 对应多个竞店品），构建多竞店对齐汇总库，"
        "支撑差异打分与选品对比。",
        "E",
        "商品名；价格（任意一种：原价/折后价/神价）",
        "商品ID；SKU id；条形码；规格；图片url；月销量；类目",
        "xlsx；列名不强制（import_adapter.py 会用 KEYWORDS 自动猜映射）。\n"
        "⚠ 列名可不一致，但至少要有「商品名」和「价格」两类列，否则 CORE_FIELDS 校验失败。",
        "竞品_<店名>_<日期>.xlsx\n（如 竞品_示例竞店B_20260807.xlsx）",
        "商品名字=公牛G12斜插座86型面板五孔 | 原价=40 | 折后价=35 | 神价=30 | 规格=一开单控(灰色) | 条形码=2085273863046144013 | 图片url=https://img.meituan.net/.../g12.jpg\n"
        "商品名字=德力西一开单控开关 | 原价=12 | 折后价=9.9 | 规格=86型 | 月售=320",
        "d:\\查图\\",
        "待提供",
    ],
]

WIDTHS_MAIN = [7, 20, 38, 9, 34, 24, 38, 26, 52, 16, 10]


def build_main_sheet(wb):
    ws = wb.create_sheet("数据准备清单")
    ws.append(HEADERS)
    for row in ROWS:
        ws.append(row)
    set_widths(ws, WIDTHS_MAIN)
    style_header_row(ws, len(HEADERS))
    style_body(ws, len(ROWS), len(HEADERS))
    # 阶段列上色（第4列）
    paint_stage_col(ws, 4, [r[3] for r in ROWS])
    # 当前状态列上色（待提供=浅红）
    status_fill = PatternFill("solid", fgColor="F8CBAD")
    for i in range(len(ROWS)):
        ws.cell(row=2 + i, column=11).fill = status_fill
    # 行高
    for i in range(len(ROWS)):
        ws.row_dimensions[2 + i].height = 110
    ws.freeze_panes = "A2"


# ===========================================================================
# Sheet 2: 执行编排图
# ===========================================================================
ORCH_HEADERS = ["阶段", "依赖数据", "前置阶段", "可并行", "输出", "说明"]
ORCH_ROWS = [
    ["B",  "①",        "无",          "C2 / E", "第4层匹配数对比（补全图片url前后）",
     "图片url 写入自家商品库，重跑第4层双信号融合，输出匹配数增益"],
    ["C2", "②",        "无",          "B / E",  "品牌命中率提升统计",
     "扩充 BRAND_DICT 后重抽品牌，统计命中率从 ~34% 提升幅度"],
    ["D2", "③",        "B + C2 完成", "无",     "第2层强组合匹配数提升",
     "类目映射写回后重跑第2层，验证强组合键匹配数提升且第3层不下沉"],
    ["E",  "④",        "无",          "B / C2", "多竞店对齐汇总库",
     "导入适配层标准化第二家竞店表，构建一对多对齐汇总"],
    ["F1", "全部",     "B+C2+D2+E 完成", "无",  "全链路重跑结果",
     "四份数据齐备后全链路重跑，产出最终归一化与对齐结果"],
    ["F2", "F1 完成",   "F1",          "无",     "回归验证矩阵",
     "按回归验证矩阵逐项校验，任一失败即中止并输出原因"],
    ["F3", "F2 通过",   "F2",          "无",     "最终匹配数对比报告",
     "回归通过后生成最终匹配数对比报告，交付选品决策"],
]
WIDTHS_ORCH = [8, 16, 18, 14, 32, 50]


def build_orch_sheet(wb):
    ws = wb.create_sheet("执行编排图")
    ws.append(ORCH_HEADERS)
    for row in ORCH_ROWS:
        ws.append(row)
    set_widths(ws, WIDTHS_ORCH)
    style_header_row(ws, len(ORCH_HEADERS))
    style_body(ws, len(ORCH_ROWS), len(ORCH_HEADERS))
    paint_stage_col(ws, 1, [r[0] for r in ORCH_ROWS])
    for i in range(len(ORCH_ROWS)):
        ws.row_dimensions[2 + i].height = 46
    ws.freeze_panes = "A2"


# ===========================================================================
# Sheet 3: 回归验证矩阵
# ===========================================================================
VERIFY_HEADERS = ["阶段", "验证项", "失败处理"]
VERIFY_ROWS = [
    ["B",
     "1) 第4层匹配数对比补全前后提升；\n"
     "2) 图片url 同源 +0.1 权重生效；\n"
     "3) 现有 4 层对齐不回归（前3层匹配数不下降）",
     "中止，输出失败原因"],
    ["C2",
     "1) 品牌命中率提升且无材质/品类/营销词误判；\n"
     "2) SPU 聚合无异常（同一 SPU 下品牌一致或为空）",
     "中止"],
    ["D2",
     "1) 第2层强组合匹配数提升；\n"
     "2) 第3层不下沉异常（未被错误归一拉低）",
     "中止"],
    ["E",
     "1) 一对多对齐验证（1 个本地 SKU → 多竞店品）；\n"
     "2) 反向查询多竞店分组正确；\n"
     "3) SKU 精度不回归",
     "中止"],
    ["全链路",
     "差异打分 / 优点分析 / 热销品 / 入库导出 / 效果追踪 全部正常",
     "中止"],
]
WIDTHS_VERIFY = [10, 70, 22]


def build_verify_sheet(wb):
    ws = wb.create_sheet("回归验证矩阵")
    ws.append(VERIFY_HEADERS)
    for row in VERIFY_ROWS:
        ws.append(row)
    set_widths(ws, WIDTHS_VERIFY)
    style_header_row(ws, len(VERIFY_HEADERS))
    style_body(ws, len(VERIFY_ROWS), len(VERIFY_HEADERS))
    paint_stage_col(ws, 1, [r[0] for r in VERIFY_ROWS])
    for i in range(len(VERIFY_ROWS)):
        ws.row_dimensions[2 + i].height = 60
    ws.freeze_panes = "A2"


# ===========================================================================
# 主流程
# ===========================================================================
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = Workbook()
    # 删除默认 Sheet
    wb.remove(wb.active)
    build_main_sheet(wb)
    build_orch_sheet(wb)
    build_verify_sheet(wb)
    wb.save(OUT_PATH)
    size = os.path.getsize(OUT_PATH)
    print(f"已生成: {OUT_PATH}")
    print(f"文件大小: {size} 字节 ({size/1024:.1f} KB)")
    print(f"Sheet 数: {len(wb.sheetnames)} -> {wb.sheetnames}")
    print(f"主表数据行: {len(ROWS)} (含表头共 {len(ROWS)+1} 行)")
    print(f"执行编排图数据行: {len(ORCH_ROWS)}")
    print(f"回归验证矩阵数据行: {len(VERIFY_ROWS)}")


if __name__ == "__main__":
    main()

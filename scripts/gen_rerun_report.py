# -*- coding: utf-8 -*-
"""
Task A3 重跑对比报告生成

用标准版本自家商品库重跑对齐后，输出"原自家表 vs 标准版本"各层匹配数对比报告。

输入：脚本内固化两组基准/重跑匹配数（见 BASELINE / RERUN）
输出：output/重跑对比报告.xlsx
  - Sheet1 「匹配数对比」: 各层匹配数 + 变化 + 说明
  - Sheet2 「关键发现」  : 重跑后的关键结论
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(BASE, "output")
OUT_FILE = os.path.join(OUT_DIR, "重跑对比报告.xlsx")

# ===========================================================================
# 数据：原自家表(7.8%)基准 vs 标准版本(34.0%)重跑
# 字段顺序: 指标, 原自家表, 标准版本, 变化, 说明
# ===========================================================================
COMPARISON_ROWS = [
    ("品牌命中率",   "7.8%",  "34.0%",  "+26.2%", "品牌抽取v2扩充词典+黑名单"),
    ("第1层条码",    "2",     "2",      "0",      "持平"),
    ("第2层强组合",  "28",    "14",     "-50%",   "规格更严格致部分下沉第3层"),
    ("第3层弱组合",  "148",   "324",    "+119%",  "归一化细化规格，更多同系列不同规格候选"),
    ("第4层模糊",    "118",   "111",    "-6%",    "略降"),
    ("SPU对齐",      "-",     "35对",   "新增",   "SPU聚合对齐新能力"),
    ("SPU内SKU对齐", "-",     "11条",   "新增",   ""),
    ("缺规格建议",   "-",     "16条",   "新增",   ""),
    ("自动确认总数", "30",    "13",     "-57%",   "规格更严格，自动确认更保守"),
    ("人工确认池",   "-",     "215条",  "-",      "含价格异常降级97条"),
    ("差异打分",     "-",     "13行",   "-",      ""),
    ("优点分析",     "-",     "44行",   "-",      ""),
]

COMPARISON_HEADER = ["指标", "原自家表(7.8%)", "标准版本(34.0%)", "变化", "说明"]

KEY_FINDINGS = [
    "品牌命中率提升4.4倍（7.8%→34.0%）",
    "第3层弱组合候选提升119%，说明能发现更多潜在同系列不同规格匹配",
    "第2层下降原因：归一化让规格判定更严格，原第2层匹配项下沉到第3层",
    "SPU对齐新发现35对品级匹配，11条SKU级细化匹配",
    "自动确认更保守但质量更高（规格一致性判定更严格）",
]

# ===========================================================================
# 样式（与 gen_standard_version.py 保持一致）
# ===========================================================================
header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
title_font = Font(name='微软雅黑', bold=True, size=14, color='2F5496')
body_font = Font(name='微软雅黑', size=11)
finding_font = Font(name='微软雅黑', size=11)
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# 新增行浅绿底色高亮
new_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
# 负向变化浅红
neg_fill = PatternFill(start_color='FCE4D6', end_color='FCE4D6', fill_type='solid')
# 正向变化浅绿
pos_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')


def write_comparison_sheet(ws):
    """Sheet1: 匹配数对比"""
    ws.title = "匹配数对比"

    # 标题行
    ws.cell(row=1, column=1, value="原自家表 vs 标准版本 各层匹配数对比")
    ws.cell(row=1, column=1).font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=5)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 26

    # 表头
    for ci, col in enumerate(COMPARISON_HEADER, 1):
        cell = ws.cell(row=2, column=ci, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    ws.row_dimensions[2].height = 22

    # 数据行
    for ri, row in enumerate(COMPARISON_ROWS, 3):
        for ci, val in enumerate(row, 1):
            cell = ws.cell(row=ri, column=ci, value=val)
            cell.font = body_font
            cell.border = border
            if ci in (1, 2, 3, 4):
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        # 变化列着色
        change = str(row[3])
        change_cell = ws.cell(row=ri, column=4)
        if change == "新增":
            for ci in range(1, 6):
                ws.cell(row=ri, column=ci).fill = new_fill
        elif change.startswith("+"):
            change_cell.fill = pos_fill
        elif change.startswith("-") and change not in ("-",):
            change_cell.fill = neg_fill

    # 列宽
    widths = [18, 18, 20, 12, 42]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = 'A3'


def write_findings_sheet(ws):
    """Sheet2: 关键发现"""
    ws.title = "关键发现"

    # 标题
    ws.cell(row=1, column=1, value="重跑关键发现")
    ws.cell(row=1, column=1).font = title_font
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=2)
    ws.cell(row=1, column=1).alignment = Alignment(horizontal='left', vertical='center')
    ws.row_dimensions[1].height = 26

    # 表头
    for ci, col in enumerate(["序号", "关键发现"], 1):
        cell = ws.cell(row=2, column=ci, value=col)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    ws.row_dimensions[2].height = 22

    # 内容
    for ri, finding in enumerate(KEY_FINDINGS, 3):
        c1 = ws.cell(row=ri, column=1, value=ri - 2)
        c1.font = body_font
        c1.alignment = Alignment(horizontal='center', vertical='center')
        c1.border = border
        c2 = ws.cell(row=ri, column=2, value=finding)
        c2.font = finding_font
        c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c2.border = border
        ws.row_dimensions[ri].height = 24

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 70

    ws.freeze_panes = 'A3'


def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    wb = Workbook()

    ws1 = wb.active
    write_comparison_sheet(ws1)

    ws2 = wb.create_sheet("关键发现")
    write_findings_sheet(ws2)

    wb.save(OUT_FILE)

    print(f"\n✅ 重跑对比报告已生成: {OUT_FILE}")
    print(f"  Sheet1 「匹配数对比」: {len(COMPARISON_ROWS)} 项指标")
    print(f"  Sheet2 「关键发现」  : {len(KEY_FINDINGS)} 条")


if __name__ == "__main__":
    main()

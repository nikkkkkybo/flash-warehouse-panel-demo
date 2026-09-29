# -*- coding: utf-8 -*-
"""生成自家商品库标准模板"""
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

wb = Workbook()
ws = wb.active
ws.title = '自家商品库'

# 标准列定义：列名 | 说明 | 是否必填 | 示例
columns = [
    ('SPUID',            '美团后台的SPU ID，19位数字，同一品不同规格共用一个SPUID', '必填', '2085208656848453670'),
    ('SKUID',            '美团后台的SKU ID，19位数字，每个规格唯一',                '必填', '2085208656861679667'),
    ('商品条码',          '商品条形码（EAN13/UPC），没有就留空',                      '选填', '6901234567890'),
    ('商品名称',          '完整的商品名称，含品牌+型号+规格+关键词',                  '必填', '公牛开关插座G32四开单控86型面板 金色'),
    ('规格名称',          '这个SKU的规格，如尺寸/颜色/功率/数量',                    '必填', '四开单控 金色'),
    ('品牌名称',          '品牌名纯文字不要带括号英文。没品牌填无，不要留空',          '必填', '公牛'),
    ('图片url',           '商品主图完整网络链接http开头，没有就留空',                '选填', 'https://img.meituan.net/xxx.jpg'),
    ('美团闪购类目一级',   '美团后台一级类目',                                        '必填', '五金日杂'),
    ('美团闪购类目二级',   '美团后台二级类目',                                        '必填', '开关插座'),
    ('美团闪购类目三级',   '美团后台三级类目',                                        '必填', '墙壁开关'),
    ('美团闪购渠道售价',   '售价元纯数字',                                            '必填', '19.9'),
    ('采购价',            '采购成本价元纯数字',                                       '选填', '12.5'),
    ('渠道库存',           '渠道库存数量',                                            '选填', '50'),
    ('实物库存',           '实物库存数量',                                            '选填', '30'),
    ('售卖状态',          '在售或下架',                                              '选填', '在售'),
    ('是否有图片',         '有图或无图',                                              '选填', '有图'),
    ('月销量',            '月销数量',                                                '选填', '100'),
    ('商品卖点',          '商品卖点描述',                                            '选填', '食品级304不锈钢'),
    ('商品描述',          '商品详情描述',                                            '选填', ''),
]

header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
required_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
optional_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
thin_border = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

# 写表头
for col_idx, (col_name, desc, required, example) in enumerate(columns, 1):
    cell = ws.cell(row=1, column=col_idx, value=col_name)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    cell.border = thin_border

# 写说明行
for col_idx, (col_name, desc, required, example) in enumerate(columns, 1):
    tag = f'[{required}] {desc}'
    cell = ws.cell(row=2, column=col_idx, value=tag)
    cell.font = Font(name='微软雅黑', size=9, color='666666', italic=True)
    cell.fill = required_fill if required == '必填' else optional_fill
    cell.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    cell.border = thin_border

# 写示例行
for col_idx, (col_name, desc, required, example) in enumerate(columns, 1):
    cell = ws.cell(row=3, column=col_idx, value=example)
    cell.font = Font(name='微软雅黑', size=9, color='999999')
    cell.alignment = Alignment(horizontal='left', vertical='center')
    cell.border = thin_border

# 列宽
widths = [22, 22, 16, 35, 20, 14, 40, 16, 16, 16, 14, 12, 10, 10, 10, 10, 10, 25, 25]
for i, w in enumerate(widths, 1):
    col_letter = chr(64 + i) if i <= 26 else 'A' + chr(64 + i - 26)
    ws.column_dimensions[col_letter].width = w

ws.freeze_panes = 'A4'

out = r'<项目根目录>\自家商品库_标准模板.xlsx'
wb.save(out)
print(f'模板已生成: {out}')
print(f'共 {len(columns)} 列')
print()
print('必填列:')
for col_name, desc, required, example in columns:
    if required == '必填':
        print(f'  - {col_name}: {desc}')
print()
print('选填列:')
for col_name, desc, required, example in columns:
    if required == '选填':
        print(f'  - {col_name}: {desc}')

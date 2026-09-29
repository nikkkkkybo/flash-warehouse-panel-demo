# -*- coding: utf-8 -*-
"""
将源商品导出表按标准模板产出标准版本自家商品库。
- SPUID/SKUID/商品条码 全程字符串保留19位精度
- 品牌名称归一：去括号英文，空/其他品牌 → 无
- 售卖状态：售卖 → 在售
- 是否有图片：是 → 有图，否 → 无图
"""
import re
import sys
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# 品牌抽取 v2：扩充词典 + 黑名单过滤
sys.path.insert(0, r'<项目根目录>\scripts')
from brand_extractor_v2 import extract_brand_v2

SRC = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'
OUT = r'<项目根目录>\自家商品库_标准版本.xlsx'

# 标准19列（顺序与模板一致）
STD_COLS = [
    'SPUID', 'SKUID', '商品条码', '商品名称', '规格名称', '品牌名称',
    '图片url', '美团闪购类目一级', '美团闪购类目二级', '美团闪购类目三级',
    '美团闪购渠道售价', '采购价', '渠道库存', '实物库存', '售卖状态',
    '是否有图片', '月销量', '商品卖点', '商品描述',
]

def norm_brand(b):
    """品牌归一：去括号英文，空/其他品牌 → 无"""
    if not isinstance(b, str):
        return '无'
    s = b.strip()
    if s in ('', 'nan', 'NaN', '其他品牌'):
        return '无'
    # 去掉中英文括号及内容：公牛（BULL）→ 公牛
    s = re.sub(r"[(（][^()（）]*[)）]", "", s).strip()
    return s if s else '无'

def norm_status(s):
    if not isinstance(s, str):
        return ''
    s = s.strip()
    if s == '售卖':
        return '在售'
    if s == '下架':
        return '下架'
    return s

def norm_img(s):
    if not isinstance(s, str):
        return '无图'
    s = s.strip()
    if s == '是':
        return '有图'
    if s == '否':
        return '无图'
    return s

def to_str(v):
    """转字符串，保留原始精度，nan→空"""
    if v is None:
        return ''
    if isinstance(v, float):
        import math
        if math.isnan(v):
            return ''
        # 整数显示为整数
        if v == int(v):
            return str(int(v))
        return str(v)
    s = str(v).strip()
    return '' if s in ('nan', 'NaN', 'None') else s

print('读取源文件...')
df = pd.read_excel(SRC, sheet_name='门店商品', dtype=str)

out = pd.DataFrame()
out['SPUID'] = df['SPUID'].map(to_str)
out['SKUID'] = df['SKUID'].map(to_str)
out['商品条码'] = df['商品条码'].map(to_str)
out['商品名称'] = df['商品名称'].map(to_str)
out['规格名称'] = df['规格名称'].map(to_str)

# 品牌抽取：原值优先 → 词典匹配 → 填"无"
print('品牌抽取（v2: 扩充词典+黑名单过滤）...')
brands, sources = [], []
for name, raw_brand in zip(df['商品名称'], df['品牌名称']):
    b, src = extract_brand_v2(name if isinstance(name, str) else '',
                              raw_brand if isinstance(raw_brand, str) else '')
    brands.append(b if b else '无')
    sources.append(src)
out['品牌名称'] = brands
out['_品牌来源'] = sources  # 临时列，统计用，不写入最终文件
out['图片url'] = ''  # 源文件无图片url
out['美团闪购类目一级'] = df['美团闪购类目一级'].map(to_str)
out['美团闪购类目二级'] = df['美团闪购类目二级'].map(to_str)
out['美团闪购类目三级'] = df['美团闪购类目三级'].map(to_str)
out['美团闪购渠道售价'] = df['美团闪购渠道售价'].map(to_str)
out['采购价'] = df['采购价'].map(to_str)
out['渠道库存'] = df['渠道库存'].map(to_str)
out['实物库存'] = df['实物库存'].map(to_str)
out['售卖状态'] = df['售卖状态'].map(norm_status)
out['是否有图片'] = df['是否有图片'].map(norm_img)
out['月销量'] = df['月销量'].map(to_str)
out['商品卖点'] = df['商品卖点'].map(to_str)
out['商品描述'] = df['商品描述'].map(to_str)

print(f'总行数: {len(out)}')

# 统计
print('\n=== 品牌统计（v2: 扩充词典+黑名单）===')
brand_vc = out['品牌名称'].value_counts()
print(f'有品牌(非"无"): {(out["品牌名称"] != "无").sum()} / {len(out)} ({(out["品牌名称"]!="无").mean()*100:.1f}%)')
print(f'无品牌: {(out["品牌名称"]=="无").sum()}')
print('Top20品牌:')
print(brand_vc.head(20))
print('\n品牌来源:')
print(out['_品牌来源'].value_counts())

print('\n=== 必填列非空检查 ===')
required = ['SPUID','SKUID','商品名称','规格名称','品牌名称','美团闪购类目一级','美团闪购类目二级','美团闪购类目三级','美团闪购渠道售价']
for c in required:
    non_empty = (out[c].astype(str).str.strip() != '').sum()
    print(f'  {c}: {non_empty}/{len(out)}')

# 写入Excel（用openpyxl保留格式 + 字符串防精度丢失）
print('\n写入标准版本...')
wb = Workbook()
ws = wb.active
ws.title = '自家商品库'

header_font = Font(name='微软雅黑', bold=True, size=11, color='FFFFFF')
header_fill = PatternFill(start_color='2F5496', end_color='2F5496', fill_type='solid')
required_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
optional_fill = PatternFill(start_color='F2F2F2', end_color='F2F2F2', fill_type='solid')
thin = Side(style='thin')
border = Border(left=thin, right=thin, top=thin, bottom=thin)

required_set = {'SPUID','SKUID','商品名称','规格名称','品牌名称','美团闪购类目一级','美团闪购类目二级','美团闪购类目三级','美团闪购渠道售价'}

# 表头
for ci, col in enumerate(STD_COLS, 1):
    cell = ws.cell(row=1, column=ci, value=col)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = border

# 数据行（全部字符串写入，防止19位精度丢失）；仅写标准19列，排除临时来源列
for ri, row in enumerate(out[STD_COLS].itertuples(index=False), 2):
    for ci, val in enumerate(row, 1):
        ws.cell(row=ri, column=ci, value=str(val) if val is not None else '')

# 列宽
widths = [22, 22, 16, 40, 20, 14, 30, 16, 16, 16, 14, 12, 10, 10, 10, 10, 10, 25, 30]
for i, w in enumerate(widths, 1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws.freeze_panes = 'A2'

wb.save(OUT)
print(f'\n✅ 标准版本已生成: {OUT}')
print(f'共 {len(out)} 行, {len(STD_COLS)} 列')

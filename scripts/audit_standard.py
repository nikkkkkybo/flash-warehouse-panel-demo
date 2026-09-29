# -*- coding: utf-8 -*-
"""
对照审核：源表 vs 标准版本
- 行数一致
- 每列可追溯到源表对应列（直传或可解释的归一化）
- 抽样核对：首/中/尾各取若干行，逐字段对照
- 品牌列：原值非空→归一化原值；原值空→从商品名抽取（可追溯）
- 数值字段：售价/采购价/库存 原文一致
"""
import re
import pandas as pd
import numpy as np

SRC = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'
STD = r'<项目根目录>\自家商品库_标准版本.xlsx'

print('读取源表和标准版本...')
src = pd.read_excel(SRC, sheet_name='门店商品', dtype=str)
std = pd.read_excel(STD, dtype=str)

print(f'源表行数: {len(src)}')
print(f'标准版本行数: {len(std)}')
assert len(src) == len(std), f'行数不一致! src={len(src)} std={len(std)}'
print('✅ 行数一致')
print()

# ========== 1. 直传字段一一对照 ==========
# (标准版本列, 源表列)
direct_map = [
    ('SPUID', 'SPUID'),
    ('SKUID', 'SKUID'),
    ('商品条码', '商品条码'),
    ('商品名称', '商品名称'),
    ('规格名称', '规格名称'),
    ('美团闪购类目一级', '美团闪购类目一级'),
    ('美团闪购类目二级', '美团闪购类目二级'),
    ('美团闪购类目三级', '美团闪购类目三级'),
    ('美团闪购渠道售价', '美团闪购渠道售价'),
    ('采购价', '采购价'),
    ('渠道库存', '渠道库存'),
    ('实物库存', '实物库存'),
    ('月销量', '月销量'),
    ('商品卖点', '商品卖点'),
    ('商品描述', '商品描述'),
]

def norm_nan(v):
    if v is None:
        return ''
    s = str(v).strip()
    return '' if s in ('nan', 'NaN', 'None') else s

print('=== 1. 直传字段对照 ===')
mismatch_total = 0
for std_col, src_col in direct_map:
    s = src[src_col].map(norm_nan)
    t = std[std_col].map(norm_nan)
    diff = (s != t).sum()
    if diff > 0:
        print(f'❌ {std_col}: {diff} 处不一致')
        # 显示前3个不一致
        idx = (s != t)
        for i in s[idx].index[:3]:
            print(f'   行{i}: 源=[{s[i]}] 标准=[{t[i]}]')
        mismatch_total += diff
    else:
        print(f'✅ {std_col}: 全部一致')
print(f'直传字段不一致总数: {mismatch_total}')
print()

# ========== 2. 归一化字段对照 ==========
print('=== 2. 归一化字段对照 ===')

# 售卖状态: 售卖→在售, 下架→下架
def norm_status(s):
    if not isinstance(s, str):
        return ''
    s = s.strip()
    if s == '售卖':
        return '在售'
    if s == '下架':
        return '下架'
    return s

src_status = src['售卖状态'].map(norm_status)
std_status = std['售卖状态']
diff = (src_status != std_status).sum()
if diff > 0:
    print(f'❌ 售卖状态: {diff} 处不一致')
    idx = (src_status != std_status)
    for i in src_status[idx].index[:5]:
        print(f'   行{i}: 源=[{src["售卖状态"][i]}] 归一=[{src_status[i]}] 标准=[{std_status[i]}]')
else:
    print(f'✅ 售卖状态: 全部一致（售卖→在售）')
print()

# 是否有图片: 是→有图, 否→无图
def norm_img(s):
    if not isinstance(s, str):
        return '无图'
    s = s.strip()
    if s == '是':
        return '有图'
    if s == '否':
        return '无图'
    return s

src_img = src['是否有图片'].map(norm_img)
std_img = std['是否有图片']
diff = (src_img != std_img).sum()
if diff > 0:
    print(f'❌ 是否有图片: {diff} 处不一致')
else:
    print(f'✅ 是否有图片: 全部一致（是→有图）')
print()

# ========== 3. 品牌名称对照（重点核查） ==========
print('=== 3. 品牌名称对照（重点）===')

def norm_brand_raw(b):
    """仅归一化原值，不抽取"""
    if not isinstance(b, str):
        return ('', '原值空')
    s = b.strip()
    if s in ('', 'nan', 'NaN', '其他品牌'):
        return ('', '原值空')
    s = re.sub(r"[(（][^()（）]*[)）]", "", s).strip()
    return (s, '原值')

# 统计品牌来源
brand_src_count = {'原值非空': 0, '原值空抽取命中': 0, '原值空抽取未命中填无': 0}
mismatch_brand = 0
for i in range(len(src)):
    raw = src['品牌名称'][i] if isinstance(src['品牌名称'][i], str) else ''
    normed, kind = norm_brand_raw(raw)
    std_b = std['品牌名称'][i]

    if normed:
        brand_src_count['原值非空'] += 1
        if normed != std_b:
            mismatch_brand += 1
            if mismatch_brand <= 5:
                print(f'❌ 行{i}: 源原值=[{raw}] 归一=[{normed}] 标准=[{std_b}]')
    else:
        # 原值空，标准版本应从商品名抽取或填"无"
        if std_b == '无':
            brand_src_count['原值空抽取未命中填无'] += 1
        else:
            brand_src_count['原值空抽取命中'] += 1
            # 反查：商品名是否含该品牌
            name = src['商品名称'][i] if isinstance(src['商品名称'][i], str) else ''
            if std_b not in name:
                mismatch_brand += 1
                if mismatch_brand <= 5:
                    print(f'❌ 行{i}: 商品名=[{name[:40]}] 标准品牌=[{std_b}] 商品名不含该品牌')

print(f'品牌来源统计: {brand_src_count}')
print(f'品牌不一致/可疑数: {mismatch_brand}')
print()

# ========== 4. 图片url 核查 ==========
print('=== 4. 图片url 核查 ===')
non_empty_url = (std['图片url'].map(norm_nan) != '').sum()
print(f'标准版本图片url非空数: {non_empty_url}（源表无图片url列，应全空）')
if non_empty_url > 0:
    print('❌ 图片url不应有值')
else:
    print('✅ 图片url全空（源表无此列，符合预期）')
print()

# ========== 5. 抽样逐字段对照 ==========
print('=== 5. 抽样逐字段对照（首3/中3/尾3）===')
sample_idx = list(range(3)) + list(range(len(src)//2 - 1, len(src)//2 + 2)) + list(range(len(src) - 3, len(src)))
for i in sample_idx:
    print(f'--- 行 {i} ---')
    print(f'  SPUID:      源=[{norm_nan(src["SPUID"][i])}] 标准=[{norm_nan(std["SPUID"][i])}]')
    print(f'  SKUID:      源=[{norm_nan(src["SKUID"][i])}] 标准=[{norm_nan(std["SKUID"][i])}]')
    print(f'  商品条码:    源=[{norm_nan(src["商品条码"][i])}] 标准=[{norm_nan(std["商品条码"][i])}]')
    name_src = norm_nan(src['商品名称'][i])[:30]
    name_std = norm_nan(std['商品名称'][i])[:30]
    print(f'  商品名称:    源=[{name_src}] 标准=[{name_std}]')
    print(f'  规格名称:    源=[{norm_nan(src["规格名称"][i])}] 标准=[{norm_nan(std["规格名称"][i])}]')
    raw_brand = norm_nan(src['品牌名称'][i])
    std_brand = std['品牌名称'][i]
    print(f'  品牌名称:    源原值=[{raw_brand}] 标准=[{std_brand}]')
    print(f'  类目三级:    源=[{norm_nan(src["美团闪购类目三级"][i])}] 标准=[{norm_nan(std["美团闪购类目三级"][i])}]')
    print(f'  售价:       源=[{norm_nan(src["美团闪购渠道售价"][i])}] 标准=[{norm_nan(std["美团闪购渠道售价"][i])}]')
    print(f'  采购价:     源=[{norm_nan(src["采购价"][i])}] 标准=[{norm_nan(std["采购价"][i])}]')
    print(f'  售卖状态:    源=[{norm_nan(src["售卖状态"][i])}] 标准=[{norm_nan(std["售卖状态"][i])}]')
print()

# ========== 6. 汇总 ==========
print('========== 审核汇总 ==========')
print(f'行数: {len(src)} == {len(std)} ✅')
print(f'直传字段不一致: {mismatch_total}')
print(f'品牌可疑: {mismatch_brand}')
print(f'图片url非空: {non_empty_url}（应为0）')
if mismatch_total == 0 and mismatch_brand == 0 and non_empty_url == 0:
    print('🎉 全部数据真实转移，无捏造')
else:
    print('⚠️ 存在问题，请查看上方明细')

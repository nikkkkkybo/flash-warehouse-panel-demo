# -*- coding: utf-8 -*-
"""核查品牌抽取误判：列出正则命中+词典命中的所有品牌，识别非品牌词"""
import re
import sys
import pandas as pd
sys.path.insert(0, r'<项目根目录>\scripts')
from normalize import clean_product_name, build_brand_set, extract_brand, normalize_existing_brand

SRC = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'
STD = r'<项目根目录>\自家商品库_标准版本.xlsx'

src = pd.read_excel(SRC, sheet_name='门店商品', dtype=str)
std = pd.read_excel(STD, dtype=str)

# 重新跑一遍抽取，拿到置信度
brand_set = build_brand_set(src)
confs = []
for i in range(len(src)):
    raw = src['品牌名称'][i] if isinstance(src['品牌名称'][i], str) else ''
    ex = normalize_existing_brand(raw)
    if ex:
        confs.append('原值')
        continue
    cleaned = clean_product_name(src['商品名称'][i])
    _, conf = extract_brand(cleaned, brand_set, existing_brand='')
    confs.append(conf)

std['_conf'] = confs

print('=== 各置信度的品牌分布 ===')
for conf in ['原值', '词典命中', '正则命中', '未命中']:
    sub = std[std['_conf'] == conf]
    print(f'\n--- {conf} ({len(sub)}条) ---')
    vc = sub['品牌名称'].value_counts()
    for b, c in vc.items():
        if c >= 5 or conf == '正则命中':
            print(f'  {b}: {c}')

# 重点：正则命中里非品牌的词
print('\n\n=== 正则命中里的可疑非品牌词 ===')
regex_sub = std[std['_conf'] == '正则命中']
suspect_words = {'LED', 'PE', 'PVC', 'PC', 'PPR', 'AB', 'UV', 'RGB', 'DC', 'AC',
                 '优选', '优选好物', '商城', '工厂', '直供', '现货', '正品',
                 'ABS', 'PA', 'PP', 'PE管', 'PVC管', '88', '86', '118', '120'}
for w in suspect_words:
    cnt = (regex_sub['品牌名称'] == w).sum()
    if cnt > 0:
        # 看一个样本商品名
        sample = regex_sub[regex_sub['品牌名称'] == w]['商品名称'].iloc[0]
        print(f'  {w}: {cnt}条  样本: {sample[:40]}')

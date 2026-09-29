# -*- coding: utf-8 -*-
"""
分析源数据商品名结构，定位品牌位置规律
- 统计括号后的开头词（品牌候选）
- 统计原品牌列为空时，商品名的开头模式
- 区分真实品牌 vs 材质/品类/营销词
"""
import re
from collections import Counter
import pandas as pd

SRC = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'
df = pd.read_excel(SRC, sheet_name='门店商品', dtype=str)

# 只看原品牌列为空的行（这些是需要抽取的）
empty_brand = df[df['品牌名称'].isna() | (df['品牌名称'].fillna('').str.strip().isin(['', '其他品牌', 'nan']))]
print(f'原品牌为空的行数: {len(empty_brand)} / {len(df)}')
print()

# === 1. 商品名括号后的开头词（品牌候选）===
# 模式：（品类词）品牌+描述
pat_after_bracket = re.compile(r'[)）]\s*([^\s\d（(]+)')
after_words = []
no_bracket = []
for name in empty_brand['商品名称'].dropna():
    name = str(name).strip()
    m = pat_after_bracket.search(name)
    if m:
        after_words.append(m.group(1))
    else:
        no_bracket.append(name)

print(f'=== 1. 括号后开头词统计（共{len(after_words)}条有括号）===')
wc = Counter(after_words)
print(f'不同开头词数: {len(wc)}')
print('Top50括号后开头词:')
for w, c in wc.most_common(50):
    print(f'  {w}: {c}')
print()

print(f'=== 无括号商品名数: {len(no_bracket)} ===')
print('前20条无括号样本:')
for s in no_bracket[:20]:
    print(f'  {s[:50]}')
print()

# === 2. 整体开头词（含无括号的）===
print('=== 2. 所有原品牌空行的商品名开头词（去括号后前2-4字）===')
all_starts = []
for name in empty_brand['商品名称'].dropna():
    name = str(name).strip()
    # 去掉开头的括号内容
    cleaned = re.sub(r'^[（(][^()（）]*[)）]\s*', '', name)
    # 取开头2-6字（到第一个数字/空格/斜杠前）
    m = re.match(r'^([^\s\d/\-×xX*（(]+)', cleaned)
    if m:
        all_starts.append(m.group(1)[:6])
    else:
        all_starts.append(cleaned[:6])

wc2 = Counter(all_starts)
print(f'不同开头词数: {len(wc2)}')
print('Top60开头词:')
for w, c in wc2.most_common(60):
    print(f'  {w}: {c}')

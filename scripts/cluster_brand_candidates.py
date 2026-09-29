# -*- coding: utf-8 -*-
"""
Task C1：未命中品牌商品名的高频词聚类
- 从品牌为"无"的29346条商品名中提取候选品牌
- 过滤黑名单（ALL_BLACKLIST）和已有词典（BRAND_DICT）
- 输出候选品牌清单供人工确认
"""
import sys
import re
from collections import Counter, defaultdict

import pandas as pd

sys.path.insert(0, r'<项目根目录>\scripts')
from brand_extractor_v2 import ALL_BLACKLIST, BRAND_DICT

SRC = r'<项目根目录>\自家商品库_标准版本.xlsx'
OUT = r'<项目根目录>\output\候选品牌清单_待确认.xlsx'

# 阈值
MIN_COUNT = 6          # 最低出现次数
MAX_BRAND_LEN = 5      # 候选品牌最大长度（>5多为品类描述）

# 补充黑名单：ALL_BLACKLIST 未覆盖的明显非品牌词（材质/品类/描述）
# 这些是从首轮候选中人工识别的明显误判，加入过滤以提高候选质量
SUPPLEMENT_BLACKLIST = {
    # 材质词
    '亚克力', '不锈钢', '纯铜', '黄铜', '全铜', '塑料', '铝合金', '铸铁',
    '全钢', '碳钢', '合金',
    # 品类/描述词
    '一进一出', '卡箍抱箍', '方体球阀', '自来水家用', '下水道疏导',
    '螺口转动e', '单针', '阳角', '阴角', '孔径', '磁面',
    '小卷生料带', '双头呆扳手', '菜盆下水管', '家用M', '双截棍',
    '安全不漏气', '密码锁单体', '秋和带网',
}

# 1. 读取未命中品牌的商品名
df = pd.read_excel(SRC, dtype=str)
miss = df[df['品牌名称'].fillna('').str.strip() == '无'].copy()
print(f'未命中品牌行数: {len(miss)} / {len(df)}')

# 2. 高频开头词聚类
# 模式：（品类词）品牌+描述
pat_bracket = re.compile(r'^[（(][^()（）]*[)）]\s*')
# 取括号后的开头词（到第一个 数字/空格/斜杠/破折号/×/x/*/括号 前）
pat_word = re.compile(r'^([^\s\d/\-×xX*（(]+)')

counter = Counter()
samples = defaultdict(list)
no_match = 0

for name in miss['商品名称'].dropna():
    name = str(name).strip()
    # 去掉开头括号内容
    cleaned = pat_bracket.sub('', name)
    m = pat_word.match(cleaned)
    if m:
        # 截断到6字（后续按长度过滤）
        word = m.group(1)[:6]
        counter[word] += 1
        if len(samples[word]) < 2:
            samples[word].append(name)
    else:
        no_match += 1

print(f'提取到开头词: {sum(counter.values())} 条, 未匹配: {no_match} 条')
print(f'不同开头词数: {len(counter)}')

# 3. 过滤并生成候选清单
candidates = []
for word, cnt in counter.items():
    if cnt < MIN_COUNT:
        continue
    if word in ALL_BLACKLIST:
        continue
    if word in SUPPLEMENT_BLACKLIST:
        continue
    if word in BRAND_DICT:
        continue
    if len(word) < 2:            # 单字不是品牌
        continue
    if len(word) > MAX_BRAND_LEN:  # >5字多为品类描述
        continue
    candidates.append({
        '候选品牌': word,
        '出现次数': cnt,
        '样本商品名': ' | '.join(samples[word]),
        '是否建议加入词典': '',  # 待人工确认
    })

# 按出现次数降序
candidates.sort(key=lambda x: -x['出现次数'])
print(f'候选品牌数: {len(candidates)}')

# 4. 输出 Excel
out_df = pd.DataFrame(candidates, columns=['候选品牌', '出现次数', '样本商品名', '是否建议加入词典'])
out_df.to_excel(OUT, index=False)
print(f'已输出: {OUT}')

# 5. 报告
print()
print('=== Top20 候选品牌 ===')
for c in candidates[:20]:
    print(f'  {c["候选品牌"]:<8} {c["出现次数"]:>4}  {c["样本商品名"][:60]}')

high_freq = [c for c in candidates if c['出现次数'] >= 20]
mid_freq = [c for c in candidates if 6 <= c['出现次数'] < 20]
print()
print(f'高频候选 (>=20次): {len(high_freq)} 个')
print(f'中频候选 (6-19次): {len(mid_freq)} 个')

# -*- coding: utf-8 -*-
"""
建立扩充品牌词典：
1. 从原品牌列提取可靠品牌
2. 从商品名高频开头词中用黑名单过滤出候选品牌
3. 输出候选品牌清单供确认
"""
import re
from collections import Counter
import pandas as pd

SRC = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'
df = pd.read_excel(SRC, sheet_name='门店商品', dtype=str)

# === 1. 原品牌列的可靠品牌 ===
def norm_raw_brand(b):
    if not isinstance(b, str):
        return ''
    s = b.strip()
    if s in ('', 'nan', 'NaN', '其他品牌'):
        return ''
    s = re.sub(r"[(（][^()（）]*[)）]", "", s).strip()
    return s

raw_brands = df['品牌名称'].map(norm_raw_brand)
raw_brands = raw_brands[raw_brands != '']
print(f'=== 1. 原品牌列可靠品牌（{len(raw_brands)}条）===')
print(raw_brands.value_counts().to_string())
print()

# === 2. 材质/品类/规格/描述词黑名单 ===
MATERIAL_BLACKLIST = {
    # 材质代号
    'PE', 'PVC', 'PP', 'PA', 'ABS', 'EVA', 'PC', 'PS', 'PMMA', 'POM',
    'PPR', 'PEX', 'PB', 'PVDF', 'PTFE', 'PET', 'HDPE', 'LDPE',
    'LED', 'RGB', 'UV', 'DC', 'AC', 'PD', 'QC', 'HDMI', 'WIFI',
    'COB', 'SMD', 'NPC', 'UVC', 'RVV', 'BLVVB', 'KBG', 'KCD',
    'DN', 'SC', 'CT', 'ST', 'MG', 'GT', 'IP', 'OA', 'DC', 'TD', 'VE',
    'SMS', 'BATTERYCABLE', 'ANENG', 'WOIMOP', 'WOIMOPT', 'KEWEIDUN',
    'LESSO', 'TCL', 'DTC', 'AB', 'DT',
    # 材质中文
    '不锈钢', '碳钢', '铁', '铜', '铝', '合金', '镀锌', '镀铬', '镀镍',
    '塑料', '尼龙', '亚克力', '硅胶', '橡胶', '陶瓷', '玻璃',
    '纯铜', '全铜', '黄铜', '青铜', '不锈钢钻尾', '304不锈钢', '410不锈钢',
    '316不锈钢', '201不锈钢', '430不锈钢', '不锈钢十字带', '不锈钢螺栓螺',
    '碳钢黑色', '碳钢镀锌',
}
CATEGORY_BLACKLIST = {
    # 品类词
    '开关', '插座', '浴霸', '灯泡', '电池', '面板', '插头', '螺丝', '膨胀',
    '吊环', '灯带', '灯条', '射灯', '轨道', '船型', '翘板', '按钮', '接头',
    '接线', '端子', '继电器', '接触器', '断路器', '漏电', '配电', '插排',
    '锅具', '筷子', '餐具', '水龙头', '门锁', '挂钩', '粘钩', '电线', '电缆',
    '吸顶灯', '角阀', '马桶', '贴纸', '墩头', '密码锁', '线缆',
    '保险丝', '熔断器', '自攻螺丝', '钻尾螺丝', '木工钉子',
    '外六角螺栓', '外六角内膨胀', '全牙内六角螺', '镀镍十字大平',
    '不锈钢钻尾螺', '电动扳手架子', '钻尾螺丝专用', '开关螺丝螺栓',
    '省力叉叉扳手', '家具木门修补', '磷化干壁钉墙', '自攻螺丝碳钢',
    '三和自喷漆黑', '内丝三通', '活络扳手', '双小钩', '管卡钉',
    '浴室单', '空鼓锤', '懒人拖把', '大扫把', '指姆滚筒', '绳子',
    '线博士', '绿林批头', '电烤盘', '四件套', '铁占子',
    # 规格/描述词
    '国标', '黑色', '白色', '红色', '蓝色', '绿色', '黄色', '灰色',
    '国标黑色', '正色', '大号', '长款', '管道款', '机型', '办公文具',
    '家庭', '商城', '工厂', '直供', '现货', '正品', '优选', '优选好物',
    '12.9高强', '8.8级外六', '150-40', '120-32', '12-19*', 'M',
    '长', '短', '大', '小', '高', '低',
}
ALL_BLACKLIST = MATERIAL_BLACKLIST | CATEGORY_BLACKLIST

# === 3. 从商品名括号后高频词过滤候选品牌 ===
empty_brand = df[df['品牌名称'].isna() | (df['品牌名称'].fillna('').str.strip().isin(['', '其他品牌', 'nan']))]

after_words = []
for name in empty_brand['商品名称'].dropna():
    name = str(name).strip()
    # 去括号内容后取开头词
    cleaned = re.sub(r'^[（(][^()（）]*[)）]\s*', '', name)
    # 取开头到第一个数字/空格/斜杠前
    m = re.match(r'^([^\s\d/\-×xX*（(]+)', cleaned)
    if m:
        w = m.group(1)
        after_words.append(w)

wc = Counter(after_words)
# 过滤黑名单，取出现>=6次的高频候选
candidates = []
for w, c in wc.most_common():
    if w in ALL_BLACKLIST:
        continue
    if c < 6:
        break
    # 过滤含数字或特殊字符的
    if re.search(r'[\d/×*\-]', w):
        continue
    # 过滤过长的（>8字可能是描述）
    if len(w) > 8:
        continue
    candidates.append((w, c))

print(f'=== 2. 过滤黑名单后的候选品牌（出现>=6次）===')
print(f'候选数: {len(candidates)}')
for w, c in candidates:
    print(f'  {w}: {c}')

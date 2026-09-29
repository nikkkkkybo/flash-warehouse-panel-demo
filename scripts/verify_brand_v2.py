# -*- coding: utf-8 -*-
"""确认之前的误判品牌已被过滤"""
import pandas as pd

std = pd.read_excel(r'<项目根目录>\自家商品库_标准版本.xlsx', dtype=str)

print('=== 之前误判的品牌是否已清除 ===')
for w in ['PE', 'PVC', 'LED', 'PA', 'PPR', 'ABS', 'EVA', '优选', '吸顶灯', '角阀', '马桶', '国标']:
    cnt = (std['品牌名称'] == w).sum()
    print(f'  {w}: {cnt}条 {"✅已清除" if cnt == 0 else "❌仍存在"}')

print()
print('=== 词典命中品牌抽样验证（商品名是否确实含该品牌）===')
import random
hit = std[std['品牌名称'].isin(['公牛','雷士照明','日丰','德力西','潜水艇','联塑','立邦','欧德宝','起点如日','家有点点','威龙原子','长鹿','顺冠'])]
for b in ['公牛','雷士照明','日丰','德力西','潜水艇','联塑','立邦','欧德宝','起点如日','家有点点','威龙原子','长鹿','顺冠']:
    sub = hit[hit['品牌名称'] == b]
    if len(sub) > 0:
        name = sub['商品名称'].iloc[0][:45]
        print(f'  {b}({len(sub)}条): {name} {"✅" if b in name else "❌商品名不含"}')

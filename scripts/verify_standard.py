# -*- coding: utf-8 -*-
"""验证标准版本：SKU精度、品牌归一、必填列"""
import pandas as pd

df = pd.read_excel(r'<项目根目录>\自家商品库_标准版本.xlsx', dtype=str)
print(f'行数: {len(df)}  列数: {len(df.columns)}')
print(f'列名: {list(df.columns)}')
print()
print('=== 前3行（SPUID/SKUID/品牌/状态/图）===')
print(df[['SPUID','SKUID','商品条码','商品名称','规格名称','品牌名称','美团闪购渠道售价','售卖状态','是否有图片']].head(3).to_string())
print()
print('=== 精度校验 ===')
print(f'SPUID长度分布: {df["SPUID"].str.len().value_counts().to_dict()}')
print(f'SKUID长度分布: {df["SKUID"].str.len().value_counts().to_dict()}')
num_pat = r"^\d+$"
spu_all_num = df["SPUID"].str.match(num_pat).all()
print(f'SPUID是否全为数字: {spu_all_num}')
print()
print('=== 售卖状态分布 ===')
print(df['售卖状态'].value_counts())
print()
print('=== 是否有图片分布 ===')
print(df['是否有图片'].value_counts())

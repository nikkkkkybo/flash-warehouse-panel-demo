# -*- coding: utf-8 -*-
"""检查源数据文件列结构"""
import pandas as pd

src = r'<项目根目录>\自家商品库\门店商品导出_1000000000000000000002.xlsx'

# 读取所有sheet
xls = pd.ExcelFile(src)
print(f'Sheets: {xls.sheet_names}')
print()

for sn in xls.sheet_names:
    df = pd.read_excel(src, sheet_name=sn, dtype=str)
    print(f'=== Sheet: {sn} ===')
    print(f'行数: {len(df)}  列数: {len(df.columns)}')
    print(f'列名:')
    for i, c in enumerate(df.columns, 1):
        non_null = df[c].notna().sum()
        sample = df[c].dropna().astype(str).head(2).tolist()
        print(f'  {i:2d}. {c}  (非空 {non_null}/{len(df)})  示例: {sample}')
    print()
    # 打印前3行完整数据
    print('前3行数据:')
    print(df.head(3).to_string())
    print()

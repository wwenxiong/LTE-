# -*- coding: utf-8 -*-
import pandas as pd, sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
df = pd.read_excel('Objects.xlsx')
print("Columns:", df.columns.tolist())
print("Shape:", df.shape)
print()
if '区域' in df.columns:
    print("区域分布:")
    print(df['区域'].value_counts().to_string())
    print()
print(df.head(20).to_string())

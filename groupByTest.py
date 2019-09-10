import pandas as pd
import numpy as np

df = pd.read_excel('.\\test_excel\\test.xlsx')

# df['date']=pd.to_datetime(df['date'])
# df=df.set_index('date')
# print(df.head())
# # print(df.groupby('name').resample('Q')['ext price'].sum())
# print(df.groupby('name').resample('Q')['ext price'].sum())

print(df[["ext price", "quantity", "unit price"]].agg(['sum', 'mean']))
print(df[["ext price", "quantity","unit price"]].sum())
print(df[["ext price", "quantity","unit price"]].mean())
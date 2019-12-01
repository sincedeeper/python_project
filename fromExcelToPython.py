import pandas as pd
import numpy as np

df = pd.DataFrame({"id": [1001, 1002, 1003, 1004, 1005, 1006],
                   "date": pd.date_range('20130102', periods=6),
                   "city": ['Beijing ', 'SH', ' guangzhou ', 'Shenzhen', 'shanghai', 'BEIJING'],
                   "age": [23, 44, 54, 32, 34, 32],
                   "category": ['100-A', '100-B', '110-A', '110-C', '210-A', '130-F'],
                   "price": [1200, np.nan, 2133, 5433, np.nan, 4432]},
                  columns=['id', 'date', 'city', 'category', 'age', 'price'])
df['city'] = df['city'].str.lower()
print(df)

df1 = pd.DataFrame({"id": [1001, 1002, 1003, 1004, 1005, 1006, 1007, 1008],
                    "gender": ['male', 'female', 'male', 'female', 'male', 'female', 'male', 'female'],
                    "pay": ['Y', 'N', 'Y', 'Y', 'N', 'Y', 'N', 'Y', ],
                    "m-point": [10, 12, 20, 40, 40, 40, 30, 20]})

print(df1)
# 数据表匹配合并
df_inner = pd.merge(df1, df, how='inner')
print(df_inner)
# 如果price列的值>3000，group列显示high，否则显示low
df_inner['group'] = np.where(df_inner['price'] > 3000, 'high', 'low')
# 对复合多个条件的数据进行分组标记
df_inner.loc[(df_inner['city'] == 'beijing') & (df_inner['price'] >= 4000), 'sign'] = 1
print(df_inner)
# #输出某列的唯一值
# print ("输出city列的城市名称，去除重复项 {0}".format(df['city'].unique()))

#判断city列的值是否为beijing，返回值为True或者False
flag=df_inner['city'].isin(['beijing'])
print(flag)
#先判断city列里是否包含beijing和shanghai，然后将符合条件的数据提取出来，满足条件的所有行都会提取出来。
flag=df_inner.loc[df_inner['city'].isin(['beijing','shanghai'])]
print(flag)
#多条件筛选
print(df_inner.loc[(df_inner['age'] > 25) & (df_inner['city'] == 'beijing'), ['id','city','age','category','gender']])

#数据分列，对category字段的值依次进行分列，并创建数据表，索引值为df_inner的索引列，列名称为category和size
df_split=pd.DataFrame((x.split('-') for x in df_inner['category']),index=df_inner.index,columns=['category','size'])
#将完成分列后的数据表与原df_inner数据表进行匹配
df_inner=pd.merge(df_inner,df_split,right_index=True, left_index=True)
print(df_inner)
# df_pivod=pd.pivot_table(df_inner,index=["city"],values=["price"],columns=["size"],aggfunc=[len,np.sum],fill_value=0,margins=True)
df_pivod=pd.pivot_table(df_inner,index=["city"],values=["price"],columns=["size"],fill_value=0,margins=True)
print(df_pivod)
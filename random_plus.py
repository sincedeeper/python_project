import random
import pandas as pd
import numpy as np

# for i in range(1,100,1):
#     j=random.randint(21,49)
#     k=random.randint(50,99)
#     print ("{0}'+'{1}'=''".format(j,k))
# 以下为Random函数的测试
# #4行5列的DF
# print(pd.DataFrame(np.random.randn(4,5)))
#
# #定义Series 注意size的用法
# print(pd.Series(np.random.randint(5,10,size=10)))
# print(pd.Series(np.random.randn(5)))
#df1=pd.DataFrame(np.random.randint(15,99,size=(100,2)))

number1=pd.Series(np.random.randint(49,99,size=100))
number2=pd.Series(np.random.randint(11,49,size=100))
list1=['+']*100
list2=['-']*100
list3=['=']*100
plusOper=pd.Series(list1)
minusOper=pd.Series(list2)
equalOper=pd.Series(list3)

df_plus=pd.concat([number1,plusOper,number2,equalOper],axis=1)
df_minus=pd.concat([number1,minusOper,number2,equalOper],axis=1)
df_result=df_plus.sample(n=50).append(df_minus.sample(n=50))

with pd.ExcelWriter('100以内加减法100道.xlsx') as writer:
   df_result.to_excel(writer, sheet_name='100以内加减法', index=False)

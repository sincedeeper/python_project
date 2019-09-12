import random
import pandas as pd
df_result_plus= pd.DataFrame()
df_result_minus= pd.DataFrame()
for i in range(1,1000,1):
    num1=random.randint(11,99)
    num2=random.randint(11,99)
    if (num1 + num2) < 100:
        str_tmp=str(num1)+'+'+str(num2)+'='
        list_tmp=list()
        list_tmp.append(str_tmp)
        df_result_plus=df_result_plus.append(pd.DataFrame([list_tmp]),ignore_index=True)

for i in range(1,1000,1):
    num1=random.randint(11,99)
    num2=random.randint(11,99)
    if (num1 - num2) > 0:
        str_tmp=str(num1)+'-'+str(num2)+'='
        list_tmp=list()
        list_tmp.append(str_tmp)
        df_result_minus=df_result_minus.append(pd.DataFrame([list_tmp]),ignore_index=True)

df_result=df_result_plus.sample(n=200).append(df_result_minus.sample(n=200))
df_result=df_result.sample(n=100)
print(df_result)
with pd.ExcelWriter('100以内加减法100道.xlsx') as writer:
   df_result.to_excel(writer, sheet_name='100以内加减法', index=False)

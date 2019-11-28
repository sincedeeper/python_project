import pandas as pd

# 当前Q或者指定Q的时间段
from pandas import Grouper

StartTime = '2019-04'
EndTime = '2019-06'
# By Year
# StartTime='2018-4'
# EndTime='2019-3'

# 从一开始有订单到现在为止
AllStartTime = '2017-01'
AllEndTime = '2019-08'

# 财年FY时间开始
FYStartTime = '2019-04'
FYEndTime = '2019-06'

# 加r以防止转义字符
# df为Raw数据，df1为清洗过的数据
# 新的订单表一定要手动删除列名上面的哪一行数据
excelFile = r'QTD订单交付报表-入账明细20190904.xlsx'
costFile = r'cost.xlsx'
df = pd.DataFrame(pd.read_excel(excelFile))
# print (df.head(n=5))
# print ('所有的列名\n{0}'.format(df.columns))
# 摘取部分列，赋值给df1,两种方法相同，用columns似乎看起来更优雅
# df1=df.loc[:,['Input Date','销售订单号','Client Name','管理分区','Unit Price','P/N','LS ManDays',
#              'Project Name','是否完成交付','财务是否入账','财务入账时间','不含税RMB','是否申请解锁']]
df1 = df[['Input Date', '销售订单号', 'Client Name', '管理分区', 'Unit Price', 'P/N', 'LS ManDays',
          'Project Name', '是否完成交付', '财务是否入账', '财务入账时间', '不含税RMB', '是否申请解锁']]

# print ('抽取的列名\n{0}'.format(df1.columns)) 初步清洗数据，将'HPC' 'HPC LICO' 'GSS'合并为HPC;
# 将'ThinkAgile' 'HX-Nutanix' 'Nutanix'
# 合并为'Nutanix' 将 'DPA12000' '数据备份' 合并成 'DPA' 将'虚拟化support' Vmware 合并为 'Vmware'

# 这种方法可以，全文搜索替换
# df1.replace('HPC LICO','HPC',inplace=True)
# df1.replace('GSS','HPC',inplace=True)
# df1.replace('ThinkAgile','Nutanix',inplace=True)
# df1.replace('HX-Nutanix','Nutanix',inplace=True)
# df1.replace( 'DPA12000','DPA',inplace=True)
# df1.replace( '数据备份','DPA',inplace=True)
# df1.replace('虚拟化support','Vmware',inplace=True)

# 初步清洗数据，将'HPC' 'HPC LICO' 'GSS'合并为HPC;将'ThinkAgile' 'HX-Nutanix' 'Nutanix' 合并为'Nutanix'
# 将 'DPA12000' '数据备份' 合并成'DPA'
# 将'虚拟化support' Vmware 合并为 'Vmware'
df1['Project Name'].replace('HPC LICO', 'HPC', inplace=True)
df1['Project Name'].replace('GSS', 'HPC', inplace=True)
# ThinkAgile类型是nutanix的三年7*24 远程技术支持，此类型不应该被替换
# df1['Project Name'].replace('ThinkAgile','Nutanix',inplace=True)
df1['Project Name'].replace('HX-Nutanix', 'Nutanix', inplace=True)
df1['Project Name'].replace('DPA12000', 'DPA', inplace=True)
df1['Project Name'].replace('数据备份', 'DPA', inplace=True)
df1['Project Name'].replace('虚拟化support', 'Vmware', inplace=True)
# 剔除Y4285-nutanix远程技术支持
# df1=df1.drop(df1[df1['Project Name'] == 'ThinkAgile'].index)

# print ('输出第一行\n{0}'.format(df1.loc[0]))
# 取出Project Name列的所有内容，转置为行排列，并转化为list,loc和[]两种方法都可以
# SPLList=df1.loc[:,['Project Name']].drop_duplicates().T.values.tolist()[:][0]
SPLList = df1['Project Name'].drop_duplicates().T.values.tolist()
# print(SPLList)

# 此时，df1为初步清洗过的数据，修改了SPL类别，需要根据情况判断是否需要剔除95Y4285-nutanix远程技术支持，line55
df_ByOrderTime = df1
df_ByFinaceTime = df1

# 开始计算指定（StartTime和ENDTime之间的数据）By 产品线计算订单数
print(
    '\n---------------------------------------开始按照指定日期，以订单数量的方式计算---'
    '------------------------------------------------------------\n')
# 剔除CNNU002的物料号
df_ByOrderTime_WithoutCNNU002 = df_ByOrderTime.drop(df_ByOrderTime[df_ByOrderTime['P/N'] == 'CNNU002'].index)
# 设置订单时间段为索引,为以订单时间为条件筛选数据做准备
df_ByOrderTime_WithoutCNNU002['Input Date'] = pd.to_datetime(df_ByOrderTime_WithoutCNNU002['Input Date'])
df_ByOrderTime_WithoutCNNU002 = df_ByOrderTime_WithoutCNNU002.set_index('Input Date')
# 获取StartTime和EndTime之间的数据；投影SPL（Group by project Name）;计算结果为BySPL 的订单数
result_Spl_Grouped_Count = df_ByOrderTime_WithoutCNNU002[StartTime:EndTime].groupby(['Project Name'])[
    'P/N'].count().rename('Total 订单数').reset_index()
# 打印出By SPL的订单数据，此处应该导出为excel表格
# print ("每个产品线的订单数\n{0}".format(result_Spl_Grouped_Count))

sheet_name1 = 'ByOrderCount_' + StartTime + '_To_' + EndTime
# 结果输出，后续应该写个函数，统一完成结果输出
# result_Spl_Grouped_Count.to_excel('result1.xlsx',sheet_name=sheet_name1,index=False)
# 按照季度，计算从2017年1月开始，到目前为止，以订单数据量方式计算


# 按照指定时间段，以订单Cost的方式计算
df_ByOrderTime.loc[:, 'Input Date'] = pd.to_datetime(df_ByOrderTime.loc[:, 'Input Date'])
df_ByOrderTime = df_ByOrderTime.set_index('Input Date')
print(
    '\n---------------------------------------开始按照指定日期，以Cost'
    '的方式计算---------------------------------------------------------------\n')
# GroupBy的时候已经要对project Name和P/N两个都分组，
df_ByOrderTime = df_ByOrderTime[StartTime:EndTime].groupby(['Project Name', 'P/N'])['LS ManDays'].sum().rename(
    '合计P/N数量').reset_index()
# print (df_ByOrderTime)
# rename sum column name，but did not work
# result_Df_ByOrderTime_Grouped.rename(columns=['Project Name','P/N','Sum P/N'],inplace=True)
# print (df_ByOrderTime)
# 读取cost表格,只需要P/N列和cost列
df_Cost = pd.DataFrame(pd.read_excel(costFile, usecols=[1, 2]))
# 加上.reset_index(drop=True)也没啥效果
df_ByOrderTime = df_ByOrderTime.merge(df_Cost)
# print (df_ByOrderTime)
# 不能用df_ByOrderTime[Cost]的方式访问,只能用iloc+索引的方式
df_ByOrderTime.loc[:, 'Total Cost'] = df_ByOrderTime.iloc[:, 2] * df_ByOrderTime.iloc[:, 3]

# print (df_ByOrderTime)
df_ByOrderTime = df_ByOrderTime.groupby('Project Name')['Total Cost'].sum().rename('合计Cost').reset_index()
# print(df_ByOrderTime)
sheet_name2 = 'ByCost_' + StartTime + '_To_' + EndTime

# 按照季度，计算从2017年1月开始，到目前为止，以Cost方式计算
# 按照季度计算所有数据，By cost方式
df_ByOrderTime_All_From2007 = df1

print(
    '\n---------------------------------------开始从2017年开始，By 季度 '
    '以Cost的方式计算所有数据，---------------------------------------------------------------\n')

df_ByOrderTime_All_From2007 = df_ByOrderTime_All_From2007.merge(df_Cost)
# 添加Total Cost这一列，Total Cost 等于每一行的物料号乘以数量
df_ByOrderTime_All_From2007['Total Cost'] = df_ByOrderTime_All_From2007.iloc[:, 6] * df_ByOrderTime_All_From2007.iloc[
                                                                                     0:, 13]
print(df_ByOrderTime_All_From2007.head())
df_ByOrderTime_All_From2007.loc[:, 'Input Date'] = pd.to_datetime(df_ByOrderTime_All_From2007.loc[:, 'Input Date'])

# 每个季度，SPL的Summary
df_ByOrderTime_All_From2007_ByQ = \
    df_ByOrderTime_All_From2007.groupby([Grouper(key='Input Date', freq='BQ'), 'Project Name'])[
        'Total Cost'].sum().rename('合计Cost').reset_index()

# 每个SPL，在每个季度的Summary
df_ByOrderTime_All_From2007_BySPL = \
    df_ByOrderTime_All_From2007.groupby(['Project Name', Grouper(key='Input Date', freq='BQ')])[
        'Total Cost'].sum().rename('合计Cost').reset_index()
# print(df_ByOrderTime_All_From2007_BySPL.head())
# print(df_ByOrderTime_All_From2007_ByQ.groupby('Project Name')['合计Cost'].sum())
# By Q total
# df_ByOrderTime_All_From2007_Total=df_ByOrderTime_All_From2007.set_index('Input Date').resample('BQ').sum().to_period('Q')
df_ByOrderTime_All_From2007_Total=df_ByOrderTime_All_From2007.groupby([Grouper(key='Input Date', freq='BQ')])[
        'Total Cost'].sum().rename('合计Cost').reset_index()
sheet_name3 = 'Total BySpl By Q'
sheet_name4 = 'Total BySpl By SPL'
sheet_name5= 'Total By Q'
# 结果输出
with pd.ExcelWriter('result.xlsx') as writer:
    result_Spl_Grouped_Count.to_excel(writer, sheet_name=sheet_name1, index=False)
    df_ByOrderTime.to_excel(writer, sheet_name=sheet_name2, index=False)
    df_ByOrderTime_All_From2007_ByQ.to_excel(writer, sheet_name=sheet_name3, index=False)
    df_ByOrderTime_All_From2007_BySPL.to_excel(writer, sheet_name=sheet_name4, index=False)
    df_ByOrderTime_All_From2007_Total.to_excel(writer, sheet_name=sheet_name5, index=False)


# 以下为根据财务解锁时间计算
# By  SPL 投影
# 根据解锁时间段选择数据
# df_ByFinaceTime['财务入账时间'] = pd.to_datetime(df_ByFinaceTime['财务入账时间'])
# df_ByFinaceTime=df_ByFinaceTime.set_index('财务入账时间')
# 获取2019年的数据
# 如下两条语句通用，建议用trancate(),但是此场景中财务入账时间有空值，不能当index,导致truncate报错
# print(df_ByFinaceTime.truncate( before=EndTime, after=StartTime))
# print(df_ByFinaceTime[StartTime:EndTime])

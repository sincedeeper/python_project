import pandas as pd
import matplotlib.pyplot as plt
# 当前Q或者指定Q的时间段
from pandas import Grouper

# 以下开始全局变量的定义
StartTime = '2019-04'
EndTime = '2019-06'

# # 从一开始有订单到现在为止
# AllStartTime = '2017-01'
# AllEndTime = '2019-08'

# # 财年FY时间开始
# FYStartTime = '2019-04'
# FYEndTime = '2019-06'

# 定义输出结果的excel表中的sheet的名字
sheet_name1 = 'Month_' + StartTime + '_To_' + EndTime
sheet_name2 = 'Rev_' + StartTime + '_To_' + EndTime
sheet_name3 = 'Total ByQ & BySQL'
sheet_name4 = 'Total BySPL'
sheet_name5 = 'Total ByQ'

# 按照已经解锁的订单计算
booking_sheet_name3 = 'Booking Total ByQ & BySQL'
booking_sheet_name4 = 'Booking Total BySPL'
booking_sheet_name5 = 'Booking Total ByQ'

# 加r以防止转义字符
# df为Raw数据，df1为清洗过的数据
# 新的订单表一定要手动删除列名上面的哪一行数据


excelFile = r'QTD订单交付报表-入账明细20190911.xlsx'
costFile = r'cost.xlsx'
df = pd.DataFrame(pd.read_excel(excelFile))
df_Cost = pd.DataFrame(pd.read_excel(costFile, usecols=[1, 2]))

# 摘取部分列，赋值给df1,两种方法相同，用columns似乎看起来更优雅
# df1=df.loc[:,['Input Date','销售订单号','Client Name','管理分区','Unit Price','P/N','LS ManDays',
#              'Project Name','是否完成交付','财务是否入账','财务入账时间','不含税RMB','是否申请解锁']]
df1 = df[['Input Date', '销售订单号', 'Client Name', '管理分区', 'Unit Price', 'P/N', 'LS ManDays',
          'Project Name', '是否完成交付', '财务是否入账', '财务入账时间', '不含税RMB', '是否申请解锁']]

# 初步清洗数据，将'HPC' 'HPC LICO' 'GSS'合并为HPC;
# 将'HX-Nutanix' 'Nutanix' 合并为'Nutanix'
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

# 获取所有的SPL 类别
# SPLList = df1['Project Name'].drop_duplicates().T.values.tolist()
# print(SPLList)

# 此时，df1为初步清洗过的数据，修改了SPL类别，需要根据情况判断是否需要剔除95Y4285-nutanix远程技术支持，line55

# 以下开始各个功能函数的定义
# 开始计算指定（StartTime和ENDTime之间的数据）By 产品线计算订单数
def specify_Month_By_OrderCount(mydf, myStartTime, myEndTime):
    print(
        '\n---------------------------------------开始按照指定日期，以订单数量的方式计算---'
        '------------------------------------------------------------\n')
    # 剔除CNNU002的物料号
    mydf = mydf.drop(mydf[mydf['P/N'] == 'CNNU002'].index)
    # 设置订单时间段为索引,为以订单时间为条件筛选数据做准备
    mydf['Input Date'] = pd.to_datetime(mydf['Input Date'])
    mydf = mydf.set_index('Input Date')
    # 获取StartTime和EndTime之间的数据；投影SPL（Group by project Name）;计算结果为BySPL 的订单数
    mydf = mydf[StartTime:EndTime].groupby(['Project Name'])[
        'P/N'].count().rename('Total 订单数').reset_index()
    # print ("每个产品线的订单数\n{0}".format(mydf))

    return mydf


#  指定订单时间内的Cost统计
def specify_Month_By_Cost(mydf, myStartTime, myEndTime):
    mydf.loc[:, 'Input Date'] = pd.to_datetime(mydf.loc[:, 'Input Date'])
    mydf = mydf.set_index('Input Date')
    print(
        '\n---------------------------------------开始按照指定日期，以Cost'
        '的方式计算---------------------------------------------------------------\n')
    # GroupBy的时候已经要对project Name和P/N两个都分组，
    mydf = mydf[myStartTime:myEndTime].groupby(['Project Name', 'P/N'])['LS ManDays'].sum().rename(
        '合计P/N数量').reset_index()
    mydf = mydf.merge(df_Cost)
    # 用索引的方式计算每个订单的总价，单价乘以物料号数量
    mydf.loc[:, 'Total Cost'] = mydf.iloc[:, 2] * mydf.iloc[:, 3]
    mydf = mydf.groupby('Project Name')['Total Cost'].sum().rename('合计Cost').reset_index()
    return mydf


# 按照季度，计算从2017年1月开始，到目前为止，以Cost方式计算
# 按照季度计算所有数据，By cost方式
def all_Month_By_Cost(mydf):
    print(
        '\n---------------------------------------开始从2017年开始，By 季度 '
        '以Cost的方式计算所有数据，---------------------------------------------------------------\n')

    mydf = mydf.merge(df_Cost)
    # 添加Total Cost这一列，Total Cost 等于每一行的物料号乘以数量
    mydf['Total Cost'] = mydf.iloc[:, 6] * mydf.iloc[0:, 13]
    mydf.loc[:, 'Input Date'] = pd.to_datetime(mydf.loc[:, 'Input Date'])

    # 每个季度，SPL的Summary
    df_ByOrderTime_All_From2007_ByQ = \
        mydf.groupby([Grouper(key='Input Date', freq='BQ'), 'Project Name'])[
            'Total Cost'].sum().rename('合计Cost').reset_index()

    # 每个SPL，在每个季度的Summary
    df_ByOrderTime_All_From2007_BySPL = \
        mydf.groupby(['Project Name', Grouper(key='Input Date', freq='BQ')])[
            'Total Cost'].sum().rename('合计Cost').reset_index()

    df_ByOrderTime_All_From2007_Total = mydf.groupby([Grouper(key='Input Date', freq='BQ')])[
        'Total Cost'].sum().rename('合计Cost').reset_index()

    return (df_ByOrderTime_All_From2007_ByQ,
            df_ByOrderTime_All_From2007_BySPL,
            df_ByOrderTime_All_From2007_Total,
            )


# 以解锁时间统计，按照季度，计算从2017年1月开始，到目前为止，以Cost方式计算，
# 按照季度计算所有数据，By cost方式
def all_Month_By_Cost_booking(mydf):
    print(
        '\n---------------------------------------以Booking时间，开始从2017年开始，By 季度 '
        '以Booking时间,以Cost的方式计算所有数据，---------------------------------------------------------------\n')
    mydf = mydf.dropna(subset=["财务入账时间"])
    mydf = mydf.merge(df_Cost)
    # 添加Total Cost这一列，Total Cost 等于每一行的物料号乘以数量
    mydf['Total Cost'] = mydf.iloc[:, 6] * mydf.iloc[0:, 13]
    mydf.loc[:, '财务入账时间'] = pd.to_datetime(mydf.loc[:, '财务入账时间'])
    # 每个季度，SPL的Summary
    df_ByBookingTime_All_From2007_ByQ = \
        mydf.groupby([Grouper(key='财务入账时间', freq='BQ'), 'Project Name'])[
            'Total Cost'].sum().rename('合计Cost').reset_index()

    # 每个SPL，在每个季度的Summary
    df_ByBookingTime_All_From2007_BySPL = \
        mydf.groupby(['Project Name', Grouper(key='财务入账时间', freq='BQ')])[
            'Total Cost'].sum().rename('合计Cost').reset_index()

    df_ByBookingTime_All_From2007_Total = mydf.groupby([Grouper(key='财务入账时间', freq='BQ')])[
        'Total Cost'].sum().rename('合计Cost').reset_index()

    return (df_ByBookingTime_All_From2007_ByQ,
            df_ByBookingTime_All_From2007_BySPL,
            df_ByBookingTime_All_From2007_Total)


# 结果输出
# 输入df1,调用三个主函数，输出的sheet_name1--5,用了全局变量
# 指定时间内的订单数统计
#result_specify_Month_By_OrderCount=specify_Month_By_OrderCount(df1,StartTime,EndTime)
# 指定订单时间内的Cost统计
result_specify_Month_By_Cost = specify_Month_By_Cost(df1, StartTime, EndTime)
(resultl_ALL_Month_By_Cost_ByQ, result_ALL_Month_By_Cost_BySPL,
 result_ALL_Month_By_Cost_Total) = all_Month_By_Cost(df1)
(result_ALL_Month_By_Cost_ByQ_Booking, result_ALL_Month_By_Cost_BySPL_Booking,
 result_ALL_Month_By_Cost_Total_Booking) = all_Month_By_Cost_booking(df1)
#柱状图，not yet work
#result_specify_Month_By_OrderCount.plot(kind='bar')
#plt.show()
# 输出到result.xlsx 文件，如果改文件存在，则覆盖
# booking开头的sheet是以入账时间进行计算
with pd.ExcelWriter('result.xlsx') as writer:
    # 指定订单时间内的订单数统计
    # result_specify_Month_By_OrderCount.to_excel(writer, sheet_name=sheet_name1, index=False)
    # 指定订单时间内的Cost统计
    result_specify_Month_By_Cost.to_excel(writer, sheet_name=sheet_name2, index=False)
    resultl_ALL_Month_By_Cost_ByQ.to_excel(writer, sheet_name=sheet_name3, index=False)
    result_ALL_Month_By_Cost_BySPL.to_excel(writer, sheet_name=sheet_name4, index=False)
    result_ALL_Month_By_Cost_Total.to_excel(writer, sheet_name=sheet_name5, index=False)
    result_ALL_Month_By_Cost_ByQ_Booking.to_excel(writer, sheet_name=booking_sheet_name3, index=False)
    result_ALL_Month_By_Cost_BySPL_Booking.to_excel(writer, sheet_name=booking_sheet_name4, index=False)
    result_ALL_Month_By_Cost_Total_Booking.to_excel(writer, sheet_name=booking_sheet_name5, index=False)

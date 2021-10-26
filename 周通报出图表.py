# -*- coding: utf-8 -*-
"""
Created on Tue Sep 29 09:15:39 2020

@author: Administrator
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib
from datetime import datetime, date, timedelta
matplotlib.rcParams['font.sans-serif'] = ['SimHei']  # 用黑体显示中文
#设置字体、图形样式
sns.set_style("whitegrid")
matplotlib.rcParams['font.sans-serif'] = ['SimHei']
matplotlib.rcParams['font.family']='sans-serif'
matplotlib.rcParams['axes.unicode_minus'] = False


#导入周期需要刷的累积数据
fig_data = pd.read_excel('【0921-0927】周报_2866.xlsx')
Before_week_fig_data = pd.read_excel('【0914-0920】周报_2866.xlsx')
data = pd.DataFrame(fig_data)
Bdata = pd.DataFrame(Before_week_fig_data)
print('本周数据行列数：',data.shape,'上周数据行列数：', Bdata.shape,'\n')


Unit_group_count = pd.DataFrame(data['单位分组'].value_counts())
print(Unit_group_count,'\n')

data[data['单位分组']== '运输单位'].count()

Unit_group_total_diatance = pd.DataFrame(data.groupby('单位分组')['总行驶里程km'].sum())
Unit_group_day_diatance = pd.DataFrame(data.groupby('单位分组')['日均运行里程km'].sum())

result = pd.merge(Unit_group_total_diatance,Unit_group_day_diatance,on = '单位分组')
result = pd.concat([Unit_group_count,result],axis = 1)
result.rename(columns={'单位分组': '总车数','总行驶里程km': '本周行驶里程（公里）','日均运行里程km': '日均行驶里程（公里）'
                      }, inplace=True)

result['总出班趟次'] = None
result['日均出班趟次'] = None
Weekly_trips = data.groupby('单位分组')['周期出班趟次'].sum()
Dayly_trips = Weekly_trips / 7
result['总出班趟次'] = Weekly_trips
result['日均出班趟次'] = Dayly_trips
result.loc['其他单位'] = result.loc['专业局'] + result.loc['其他']
result['单车日均行驶里程（公里）'] =result['日均行驶里程（公里）'] /result['总车数']
result = result[['总车数','本周行驶里程（公里）','日均行驶里程（公里）','单车日均行驶里程（公里）','总出班趟次','日均出班趟次']]
result.drop(['专业局' ,'其他'],axis=0, index=None, columns=None, inplace=True)
result.index.name = '单位'
#基本运行情况用输入ppt数据
result = result.reindex(index=['运输单位','经营单位','其他单位'])
print('本周共行驶:',result['本周行驶里程（公里）'].sum()/10000,'万公里')
print('本周日均行驶:',result['日均行驶里程（公里）'].sum()/10000,'万公里')
print('本周单车日均行驶里程（公里）:',result['日均行驶里程（公里）'].sum()/result['总车数'].sum())
print('本周减去上周总行驶里程：',(result['本周行驶里程（公里）'].sum() - Bdata['总行驶里程km'].sum())/10000,'万公里')
print('本周减去上周——日均运行里程km:',data['日均运行里程km'].sum() - Bdata['日均运行里程km'].sum(),'公里')
print('中心局行驶',result.loc['运输单位','本周行驶里程（公里）']/10000,'万公里\n'
      ,'日均行驶',result.loc['运输单位','日均行驶里程（公里）']/10000,'万公里\n'
      ,'单车日均',result.loc['运输单位','单车日均行驶里程（公里）'],'公里\n')
print('区分公司行',result.loc['经营单位','本周行驶里程（公里）']/10000,'万公里\n'
      ,'日均行驶',result.loc['经营单位','日均行驶里程（公里）']/10000,'万公里\n'
      ,'单车日均',result.loc['经营单位','单车日均行驶里程（公里）'],'公里\n')
print('其他单位（含专业局和其他单位）行驶',result.loc['其他单位','本周行驶里程（公里）']/10000,'万公里\n'
      ,'日均行驶',result.loc['其他单位','日均行驶里程（公里）']/10000,'万公里\n'
      ,'单车日均',result.loc['其他单位','单车日均行驶里程（公里）'],'公里\n')
#显示为两位小数
print(result.round(decimals=2),'\n')

#基本运行情况用输入ppt图片
#取做图数据
x=['中心局','区分公司','其他单位']
y1=result['本周行驶里程（公里）']
y2=result['日均行驶里程（公里）']
y3=result['单车日均行驶里程（公里）']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1 = fig.add_subplot(111)
#设置标题
ax1.set_title("本周行驶里程",fontsize='24')  
plt.bar(x=range(len(x)), height=y1, label='本周行驶里程（公里）', color='steelblue', alpha=0.8, width=bar_width)
plt.bar(x=np.arange(len(x)) + bar_width, height=y2, label='日均行驶里程（公里）', color='indianred', alpha=0.8, width=bar_width)
# 显示图例
plt.legend()
# 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
for x1, yy in enumerate(y1):
    plt.text(x1, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
for x1, yy in enumerate(y2):
    plt.text(x1 + bar_width, yy + 1, '%.2f'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
#画折线图
ax2 = ax1.twinx() 
ax2.plot(x, y3, 'g',marker='*',ms=10,label='单车日均行驶里程（公里）')
# 显示图例
plt.legend(loc='upper right', fontsize=10,bbox_to_anchor=(1,0.9))
ax2.set_xlim([-0.7,2.7])
ax2.set_ylim([0,200])
ax2.grid(None)

for x1, yy in enumerate(y3):
    plt.text(x1, yy + 4, '%.2f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)

plt.savefig('./本周行驶里程.png',dpi=600,bbox_inches = 'tight')


#本周与上周各车型公里数情况制表
This_week = data.groupby('车型')['总行驶里程km'].sum()
Last_week = Bdata.groupby('车型')['总行驶里程km'].sum()
This_week= pd.DataFrame(This_week).T
Last_week = pd.DataFrame(Last_week).T
DistanceVS_week = pd.concat([This_week, Last_week],axis = 0)
DistanceVS_week.columns.name = ''
DistanceVS_week.index = [ '本周', '上周']
DistanceVS_week.loc['变化幅度'] = DistanceVS_week.loc['本周'] - DistanceVS_week.loc['上周']
DistanceVS_week.index.name = '周期'
DistanceVS_week = DistanceVS_week[['轻型车','3吨','5吨','8吨','12吨','牵引头']]
print(DistanceVS_week,'\n')



#本周与上周各车型公里数情况制图
#取做图数据
result2 = DistanceVS_week.T
CarModule = ['轻型车','3吨','5吨','8吨','12吨','牵引头']
x_CarModule   = range(len(CarModule ))

y1_cycle_distance=result2['本周']
y2_cycle_distance=result2['上周']
y3_cycle_distance=result2['变化幅度']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1_Moduledistance_cycleVS = fig.add_subplot(111)
#设置标题
ax1_Moduledistance_cycleVS.set_title("本周与上周各车型公里数情况",fontsize='24')  
plt.bar(x=range(len(CarModule)), height=y1_cycle_distance, label='本周', color='blue', alpha=0.5, width=bar_width)
plt.bar(x=np.arange(len(CarModule)) + bar_width, height=y2_cycle_distance, label='上周', color='orange', alpha=0.5, width=bar_width)
# 显示图例
plt.legend()
#画折线图
ax2_Moduledistance_cycleVS = ax1_Moduledistance_cycleVS.twinx()  # 这个很重要
ax2_Moduledistance_cycleVS.plot(x_CarModule ,y3_cycle_distance, 'g',marker='*',ms=10,label='变化幅度')
# 显示图例
plt.xticks(x_CarModule,CarModule)
plt.legend(loc='right', fontsize=10,bbox_to_anchor=(0.9,0.94))
ax2_Moduledistance_cycleVS.grid(None)
plt.savefig('./本周与上周各车型公里数情况.png',dpi=600,bbox_inches = 'tight')

#取做图数据
x = ['中心局','区分公司','其他单位']
x_shift_times = range(len(x))
y1_all_shift_times=result['总出班趟次']
y2_dayly_shift_times=result['日均出班趟次']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1_Shift_times_this_week = fig.add_subplot(111)
#设置标题
ax1_Shift_times_this_week.set_title("本周出班趟次",fontsize='24')  
#绘制条形图
plt.bar(x=range(len(x_shift_times)), height=y1_all_shift_times, label='总出班趟次', color='steelblue', alpha=0.8, width=bar_width)
plt.bar(x=np.arange(len(x_shift_times)) + bar_width, height=y2_dayly_shift_times, label='日均出班趟次', color='green', alpha=0.8, width=bar_width)
# 显示图例
plt.legend()
# 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
for x1_all_shift_times, yy_all_shift_times in enumerate(y1_all_shift_times):
    plt.text(x1_all_shift_times, yy_all_shift_times + 1, '%.0f'%yy_all_shift_times, ha='center', va='bottom', fontsize=16, rotation=0)
for x1_dayly_shift_times, yy_dayly_shift_times in enumerate(y2_dayly_shift_times):
    plt.text(x1_dayly_shift_times + bar_width, yy_dayly_shift_times + 1, '%.0f'%yy_dayly_shift_times, ha='center', va='bottom', fontsize=16, rotation=0)

#修改x轴名称
plt.xticks(x_shift_times,x)
plt.savefig('./本周出班趟次.png',dpi=600,bbox_inches = 'tight')
    

result.loc['合计'] = result.loc['运输单位'] + result.loc['经营单位'] + result.loc['其他单位']
#显示为两位小数
print(result.round(decimals=2),'\n')



print('本周共出班', result.loc['合计','总出班趟次'],'趟次')
print('本周共出班', result.loc['合计','日均出班趟次'],'趟次')
print('中心局共出班',result.loc['运输单位','总出班趟次'],'趟次,','日均出班',result.loc['运输单位','日均出班趟次'],'趟次')
print('区分公司共出班',result.loc['经营单位','总出班趟次'],'趟次,','日均出班',result.loc['经营单位','日均出班趟次'],'趟次')
print('其他单位（含专业局和其他单位）出班',result.loc['其他单位','总出班趟次'],'趟次,','日均出班',result.loc['其他单位','日均出班趟次'],'趟次')
print('本周比(减去)上周',result.loc['合计','总出班趟次'] - Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['专业局'] -
      Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['其他'] - Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['经营单位']-
      Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['运输单位']
      )
print('日均出班增加',(result.loc['合计','总出班趟次'] - Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['专业局'] -
      Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['其他'] - Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['经营单位']-
      Bdata.groupby('单位分组')['周期出班趟次'].sum().loc['运输单位'])/7
      )


#本周与上周各车型出班趟次情况制表
This_week_module_trips = data.groupby('车型')['周期出班趟次'].sum()
Last_week_module_trips = Bdata.groupby('车型')['周期出班趟次'].sum()
This_week_module_trips = pd.DataFrame(This_week_module_trips).T
Last_week_module_trips = pd.DataFrame(Last_week_module_trips).T
Module_tripVS_week = pd.concat([This_week_module_trips, Last_week_module_trips],axis = 0)
Module_tripVS_week.columns.name = ''
Module_tripVS_week.index = [ '本周', '上周']
Module_tripVS_week.loc['变化幅度'] = Module_tripVS_week.loc['本周'] - Module_tripVS_week.loc['上周']
Module_tripVS_week.index.name = '周期'
Module_tripVS_week = Module_tripVS_week[['轻型车','3吨','5吨','8吨','12吨','牵引头']]
print(Module_tripVS_week)

#本周与上周各车型公里数情况制图
#取做图数据
result3 = Module_tripVS_week.T
CarModule = ['轻型车','3吨','5吨','8吨','12吨','牵引头']

x_CarModule   = range(len(CarModule ))

y1_cycle_trip =result3['本周']
y2_cycle_trip =result3['上周']
y3_cycle_trip =result3['变化幅度']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1_Module_tripVS_week = fig.add_subplot(111)

#设置标题
ax1_Module_tripVS_week.set_title("本周与上周各车型出班趟次情况",fontsize='24')  
plt.bar(x=range(len(CarModule)), height=y1_cycle_trip, label='本周', color='green', alpha=0.9, width=bar_width)
plt.bar(x=np.arange(len(CarModule)) + bar_width, height=y2_cycle_trip, label='上周', color='orange', alpha=0.9, width=bar_width)
# 显示图例
plt.legend()

#画折线图
ax2_Module_tripVS_week = ax1_Module_tripVS_week.twinx()  # !!!这个很重要
ax2_Module_tripVS_week.plot(x_CarModule ,y3_cycle_trip, 'g',marker='*',ms=10,label='变化幅度')

# 显示图例

plt.xticks(x_CarModule,CarModule)
plt.legend(loc='right', fontsize=10,bbox_to_anchor=(0.9,0.94))

ax2_Module_tripVS_week.grid(None)

plt.savefig('./本周与上周各车型出班趟次情况.png',dpi=600,bbox_inches = 'tight')


#本周行驶里程小于50公里制表
data.loc[(data['总行驶里程km']<50),'本周'] = 1
This_week_distance50_count = data.groupby('配属单位')['本周'].sum()
This_week_distance50_count = pd.DataFrame(This_week_distance50_count)
This_week_distance50_count.loc['区分公司合计'] = This_week_distance50_count.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
'北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
'北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
'北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
'北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
This_week_distance50_count.loc['中心局合计'] = This_week_distance50_count.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
This_week_distance50_count.loc['其他单位合计'] = This_week_distance50_count.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
'国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司' ,'业务代理局']].apply(lambda x: x.sum())
This_week_distance50_count.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局'],axis = 0,inplace = True)
print('区分公司',This_week_distance50_count.loc['区分公司合计','本周'])
print('中心局',This_week_distance50_count.loc['中心局合计','本周'])
print('其他单位（含专业局和其他单位）',This_week_distance50_count.loc['其他单位合计','本周'],'各个占比看饼图')

#本周行驶里程小于50公里制图
This_week_distance50_count_for_plot1 = This_week_distance50_count.loc[['区分公司合计', '中心局合计', '其他单位合计']]
This_week_distance50_count_for_plot1.rename(columns = {'本周':''},inplace = True)
explode = (0.1, 0, 0) 
This_week_distance50_count_for_plot1.plot.pie(explode=explode,labels=['区分公司合计', '中心局合计', '其他单位合计'], colors=['r', 'g', 'b'],
                 autopct='%.1f%%', fontsize=16, figsize=(8, 8),pctdistance=0.6,subplots=True,shadow=True)
plt.title('本周行驶里程小于50公里',fontsize=24 )
plt.savefig('./本周行驶里程小于50公里.png',dpi=600,bbox_inches = 'tight')


Bdata.loc[(Bdata['总行驶里程km']<50),'上周'] = 1
Last_week_distance50_count = Bdata.groupby('配属单位')['上周'].sum()
Last_week_distance50_count = pd.DataFrame(Last_week_distance50_count)
Last_week_distance50_count.loc['区分公司合计'] = Last_week_distance50_count.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
'北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
'北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
'北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
'北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
Last_week_distance50_count.loc['中心局合计'] = Last_week_distance50_count.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
Last_week_distance50_count.loc['其他单位合计'] = Last_week_distance50_count.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
'国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局' ]].apply(lambda x: x.sum())
Last_week_distance50_count.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局'],axis = 0,inplace = True)

Distance50_VS_week = pd.concat([This_week_distance50_count, Last_week_distance50_count],axis = 1)
Distance50_VS_week['差额'] = Distance50_VS_week['本周'] - Distance50_VS_week['上周']
Distance50_VS_week.loc['总计'] = Distance50_VS_week.loc[['区分公司合计','中心局合计','其他单位合计']].apply(lambda x: x.sum())

Distance50_VS_week.rename(index={'北京市朝阳区邮电局':'北京市朝阳区分公司','北京市西城区邮电局':'北京市西城区分公司','北京市东城区邮电局':'北京市东城区分公司','北京市海淀区邮电局':'北京市海淀区分公司',
'北京市大兴区邮政局':'北京市大兴区分公司','北京市通州区邮政局':'北京市通州区分公司','北京市丰台区邮电局':'北京市丰台区分公司',
'北京市延庆县邮政局':'北京市延庆县分公司','北京市房山区邮政局':'北京市房山区分公司','北京市昌平区邮政局':'北京市昌平区分公司',
'北京市门头沟区邮政局':'北京市门头沟区分公司','北京市密云区邮政局':'北京市密云区分公司','北京市顺义区邮政局':'北京市顺义区分公司',
'北京市石景山区邮电局':'北京市石景山区分公司','北京市怀柔区邮政局':'北京市怀柔区分公司','北京市平谷区邮政局':'北京市平谷区分公司'
    }, inplace=True)

#ppt需要展示的表格
print(Distance50_VS_week)
print('（二）行驶里程小于50公里车辆对比分析')
print('本周累计行驶里程小于50公里的车辆数与上周相比%s辆，' %  Distance50_VS_week.loc['总计','差额'])
print(round((Distance50_VS_week.loc['总计','差额'] /Distance50_VS_week.loc['总计','上周'])*100,1),'%')
print('各单位组具体情况如下：')


This_week_distance50_count_for_plot2 = This_week_distance50_count_for_plot1.loc[['区分公司合计', '中心局合计', '其他单位合计']]
This_week_distance50_count_for_plot2_sum = int(This_week_distance50_count_for_plot2[:].sum())
This_week_distance50_count_for_plot2 = This_week_distance50_count_for_plot2.apply(lambda x: x/This_week_distance50_count_for_plot2_sum)

This_week_distance50_count_for_plot2 .rename(columns = {'':'本周占比'},inplace = True)

Last_week_distance50_count_for_plot2 = Last_week_distance50_count.loc[['区分公司合计', '中心局合计', '其他单位合计']]
Last_week_distance50_count_for_plot2.rename(columns = {'上周':'上周占比'},inplace = True)
Last_week_distance50_count_for_plot2_sum = int(Last_week_distance50_count_for_plot2[:].sum())
Last_week_distance50_count_for_plot2 = Last_week_distance50_count_for_plot2.apply(lambda x: x/Last_week_distance50_count_for_plot2_sum)

Distance50_week_percentage= pd.merge(This_week_distance50_count_for_plot2,Last_week_distance50_count_for_plot2,on='配属单位',how= 'left')
Distance50_week_percentage['变化幅度'] = Distance50_week_percentage['本周占比'] - Distance50_week_percentage['上周占比']



#取做图数据
x=['区分公司','中心局','其他单位']
y1_Distance50_week_percentage = 100*Distance50_week_percentage['本周占比']
y2_Distance50_week_percentage = 100*Distance50_week_percentage['上周占比']
y3_Distance50_week_percentage = 100*Distance50_week_percentage['变化幅度']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig_Distance50_week_percentage = plt.figure()
ax1_Distance50_week_percentage= fig_Distance50_week_percentage.add_subplot(111)
#设置标题
ax1_Distance50_week_percentage.set_title("行驶里程小于50公里占比对比",fontsize='18')  
plt.bar(x=range(len(x)), height=y1_Distance50_week_percentage, label='本周占比', color='blue', alpha=0.7, width=bar_width)
plt.bar(x=np.arange(len(x)) + bar_width, height=y2_Distance50_week_percentage, label='上周占比', color='orange', alpha=0.7, width=bar_width)
# 显示图例
plt.legend(loc='upper left')

# 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
for x1, yy in enumerate(y1_Distance50_week_percentage):
    plt.text(x1, yy+ 1, '%.1f%%'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
for x1, yy in enumerate(y2_Distance50_week_percentage):
    plt.text(x1 + bar_width, yy+ 1, '%.1f%%'%yy, ha='center', va='bottom', fontsize=16, rotation=0)
    
#画折线图
ax2_Distance50_week_percentage= ax1_Distance50_week_percentage.twinx()  # 这个很重要
ax2_Distance50_week_percentage.plot(x, y3_Distance50_week_percentage, 'g',marker='*',ms=10,label='变化幅度')


# 显示图例
plt.legend(loc='upper right', fontsize=10)
ax2_Distance50_week_percentage.set_xlim([-0.7,2.7])
ax2_Distance50_week_percentage.set_ylim([-10,10])
ax2_Distance50_week_percentage.grid(None)


for x1, yy in enumerate(y3_Distance50_week_percentage):
    plt.text(x1, yy + 1, '%.1f%%'%yy, ha='center', va='bottom', fontsize=15, rotation=0)

plt.savefig('./行驶里程小于50公里占比对比.png',dpi=600,bbox_inches = 'tight')



Distance50_week_percentage['本周占比'] = Distance50_week_percentage['本周占比'] .apply(lambda x: format(x, '.2%'))
Distance50_week_percentage['上周占比'] = Distance50_week_percentage['上周占比'] .apply(lambda x: format(x, '.2%'))
Distance50_week_percentage['变化幅度'] = Distance50_week_percentage['变化幅度'].apply(lambda x: format(x, '.2%'))

print(Distance50_week_percentage)
print('区分公司环比上周%s；' %  Distance50_week_percentage.loc['区分公司合计','变化幅度'])
print('中心局环比上周%s；' % Distance50_week_percentage.loc['中心局合计','变化幅度'])
print('其他单位（含专业局和其他单位）环比上周%s。' % Distance50_week_percentage.loc['其他单位合计','变化幅度'] )


Distance50_VS_week.loc['总计'] = Distance50_VS_week.loc['区分公司合计'] + Distance50_VS_week.loc['中心局合计'] +  Distance50_VS_week.loc['其他单位合计'] 
Distance50_VS_week_data = pd.DataFrame(Distance50_VS_week.loc['总计',['本周','上周']])  
# 设置图框的大小
fig = plt.figure(figsize=(10,6))

x_Distance50_VS_week_data=['本周','上周']
y_Distance50_VS_week_data = Distance50_VS_week_data['总计']
# 绘图
plt.plot(x_Distance50_VS_week_data, 
         y_Distance50_VS_week_data, 
         linestyle = '-', 
         linewidth = 5, 
         color = 'orange', 
         marker = 'o', 
         markersize = 2, 
         markeredgecolor='black', 
         markerfacecolor='steelblue') 

# 添加标题和坐标轴标签
plt.title('本周与上周里程少于50公里的车辆数')
plt.ylabel('车量数(量)')

# 剔除图框上边界和右边界的刻度
plt.tick_params(top = 'off', right = 'off')

# 获取图的坐标信息
ax = plt.gca()
ax.set_xlim([-0.4,1.4])
ax.set_ylim([0,600])


for x1, yy in enumerate(y_Distance50_VS_week_data):
    plt.text(x1, yy + 10, '%.0f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)
# 保存图形
plt.savefig('./本周与上周里程少于50公里的车辆数.png',dpi=600,bbox_inches = 'tight')


#出班少于3天车数对比
data.loc[(data['周期出班趟次']<3),'本周<3'] = 1
This_week_trip3_count = data.groupby('配属单位')['本周<3'].sum()
This_week_trip3_count = pd.DataFrame(This_week_trip3_count)
This_week_trip3_count.loc['区分公司合计'] = This_week_trip3_count.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
'北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
'北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
'北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
'北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
This_week_trip3_count.loc['中心局合计'] = This_week_trip3_count.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
This_week_trip3_count.loc['其他单位合计'] = This_week_trip3_count.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
'国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局' ]].apply(lambda x: x.sum())
#注意这里业务代理局是否有！！！
This_week_trip3_count.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局'],axis = 0,inplace = True)
print(This_week_trip3_count)


Bdata.loc[(Bdata['周期出班趟次']<3),'上周<3'] = 1
Last_week_trip3_count = Bdata.groupby('配属单位')['上周<3'].sum()
Last_week_trip3_count = pd.DataFrame(Last_week_trip3_count)
Last_week_trip3_count.loc['区分公司合计'] = Last_week_trip3_count.loc[['北京市朝阳区邮电局','北京市西城区邮电局','北京市东城区邮电局','北京市海淀区邮电局',
'北京市大兴区邮政局','北京市通州区邮政局','北京市丰台区邮电局',
'北京市延庆县邮政局','北京市房山区邮政局','北京市昌平区邮政局',
'北京市门头沟区邮政局','北京市密云区邮政局','北京市顺义区邮政局',
'北京市石景山区邮电局','北京市怀柔区邮政局','北京市平谷区邮政局']].apply(lambda x: x.sum())
Last_week_trip3_count.loc['中心局合计'] = Last_week_trip3_count.loc[['北京邮区中心局','运输公司']].apply(lambda x: x.sum())
Last_week_trip3_count.loc['其他单位合计'] = Last_week_trip3_count.loc[['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心',
'国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局' ]].apply(lambda x: x.sum())
Last_week_trip3_count.drop(['北京国际邮电局','北京市通信局','北京报刊零售公司','北京邮政电商分销局',
'北京邮政商业信函局','北京邮政信息技术局','电商营销中心','国际分公司','航空邮件处理中心','科研院','同城运营中心','物流分公司','业务代理局'],axis = 0,inplace = True)


Trip3_VS_week = pd.concat([This_week_trip3_count, Last_week_trip3_count],axis = 1)
Trip3_VS_week['差额'] = Trip3_VS_week['本周<3'] - Trip3_VS_week['上周<3']

Trip3_VS_week.loc['总计'] = Trip3_VS_week.loc[['区分公司合计','中心局合计','其他单位合计']].apply(lambda x: x.sum())
Trip3_VS_week.rename(index={'北京市朝阳区邮电局':'北京市朝阳区分公司','北京市西城区邮电局':'北京市西城区分公司','北京市东城区邮电局':'北京市东城区分公司','北京市海淀区邮电局':'北京市海淀区分公司',
'北京市大兴区邮政局':'北京市大兴区分公司','北京市通州区邮政局':'北京市通州区分公司','北京市丰台区邮电局':'北京市丰台区分公司',
'北京市延庆县邮政局':'北京市延庆县分公司','北京市房山区邮政局':'北京市房山区分公司','北京市昌平区邮政局':'北京市昌平区分公司',
'北京市门头沟区邮政局':'北京市门头沟区分公司','北京市密云区邮政局':'北京市密云区分公司','北京市顺义区邮政局':'北京市顺义区分公司',
'北京市石景山区邮电局':'北京市石景山区分公司','北京市怀柔区邮政局':'北京市怀柔区分公司','北京市平谷区邮政局':'北京市平谷区分公司'
    }, inplace=True)
#ppt需要展示的表格
Trip3_VS_week.rename(columns={'本周<3':'本周','上周<3':'上周'},inplace=True)
print('区分公司%s辆，占比%s' % (Trip3_VS_week.loc['区分公司合计','本周'],100*Trip3_VS_week.loc['区分公司合计','本周']/Trip3_VS_week.loc['总计','本周']),'%；')
print('中心局%s辆，占比%s' % (Trip3_VS_week.loc['中心局合计','本周'],100*Trip3_VS_week.loc['中心局合计','本周']/Trip3_VS_week.loc['总计','本周']),'%；')
print('其他单位合计%s辆，占比%s' % (Trip3_VS_week.loc['其他单位合计','本周'],100*Trip3_VS_week.loc['其他单位合计','本周']/Trip3_VS_week.loc['总计','本周']),'%；')





Trip3_VS_week = Trip3_VS_week.T

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1_trip3 = fig.add_subplot(111)
x_trip3_VS_week_data = ['本周','上周','差额']

y1_trip3_VS_week_data = Trip3_VS_week['区分公司合计']
y2_trip3_VS_week_data  = Trip3_VS_week['中心局合计']
y3_trip3_VS_week_data  = Trip3_VS_week['其他单位合计']
#plt.plot(x, y, 'ro-')
#plt.plot(x, y1, 'bo-')
plt.xlim(-0.2, 2.2)  # 限定横轴的范围
plt.ylim(-100, 300)  # 限定纵轴的范围
 
 
plt.plot(x_trip3_VS_week_data,y1_trip3_VS_week_data , marker='o', mec='r', mfc='w',label='区分公司合计')
plt.plot(x_trip3_VS_week_data,y2_trip3_VS_week_data , marker='*', ms=10,label='中心局合计')
plt.plot(x_trip3_VS_week_data,y3_trip3_VS_week_data , marker='.', ms=10,label='其他单位合计')
plt.legend()  # 让图例生效

 
plt.margins(0)
plt.subplots_adjust(bottom=0.3)

#编注文字
for x1, yy in enumerate(y1_trip3_VS_week_data):
    plt.text(x1-0.05, yy + 10, '%.0f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)
for x1, yy in enumerate(y2_trip3_VS_week_data):
    plt.text(x1+0.05, yy  , '%.0f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)
for x1, yy in enumerate(y3_trip3_VS_week_data):
    plt.text(x1, yy - 30, '%.0f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)
plt.title('出班少于3天车数对比',fontsize = 18)

# 保存图形
plt.savefig('./出班少于3天车数对比.png',dpi=600,bbox_inches = 'tight')

data.loc[(data['周期出班趟次']<3),'本周<3'] = 1
data_pivot = data[['单位分组','车型','本周<3']]
data_pivot = data_pivot.groupby(['单位分组','车型'])['本周<3'].count()
data_pivot = pd.DataFrame(data_pivot)
data_pivot = pd.pivot_table(data_pivot,index=['单位分组'],columns=['车型'])
data_pivot = data_pivot.fillna(0)
data_pivot.loc['其他单位'] = data_pivot.loc['专业局'] + data_pivot.loc['其他']
data_pivot.drop(['专业局' ,'其他'],axis=0, index=None, columns=None, inplace=True)
print(data_pivot )
data_pivot.loc['总计'] = data_pivot.apply(lambda x: x.sum())



data_pivot.loc['占比%'] = round(100*data_pivot.loc['总计']/Trip3_VS_week.T.loc['总计','本周'],1)
print(data_pivot)
print('查看上表取数据（总计和占比）：轻型车辆146辆，占比51.6%；3吨42辆，占比14.8%；5吨49辆，占比17.3%；8吨28辆，占比9.9%；12吨10辆，占比3.5%；牵引头8辆，占比2.8%。')
#取做图数据
x_trip3 = ['12吨','3吨','5吨','8吨','牵引头','轻型车']

x_trip3_len = range(len(x_trip3 ))

y1_trip3 =data_pivot.loc['经营单位']
y2_trip3 =data_pivot.loc['运输单位']
y3_trip3 =data_pivot.loc['其他单位']
bar_width = 0.25

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig = plt.figure()
ax1_trip3 = fig.add_subplot(111)
#设置标题
ax1_trip3.set_title("本周出班趟次",fontsize='18')  
#绘制条形图
plt.bar(x=range(len(x_trip3)), height=y1_trip3, label='经营单位', color='steelblue', alpha=0.8, width=bar_width)
plt.bar(x=np.arange(len(x_trip3)) + bar_width, height=y2_trip3, label='运输单位', color='green', alpha=0.8, width=bar_width)
plt.bar(x=np.arange(len(x_trip3)) + 2*bar_width, height=y3_trip3, label='其他单位', color='orange', alpha=0.8, width=bar_width)
# 显示图例
plt.legend()

# 在柱状图上显示具体数值, ha参数控制水平对齐方式, va控制垂直对齐方式
for x1_all_shift_times, yy_all_shift_times in enumerate(y1_trip3):
    plt.text(x1_all_shift_times, yy_all_shift_times + 1, '%.0f'%yy_all_shift_times, ha='center', va='bottom', fontsize=12, rotation=0)
for x1_dayly_shift_times, yy_dayly_shift_times in enumerate(y2_trip3):
    plt.text(x1_dayly_shift_times + bar_width, yy_dayly_shift_times + 1, '%.0f'%yy_dayly_shift_times, ha='center', va='bottom', fontsize=12, rotation=0)
for x1_dayly_shift_times, yy_dayly_shift_times in enumerate(y3_trip3):
    plt.text(x1_dayly_shift_times + 2*bar_width, yy_dayly_shift_times + 1, '%.0f'%yy_dayly_shift_times, ha='center', va='bottom', fontsize=12, rotation=0)

#修改x轴名称
plt.xticks(x_trip3_len ,x_trip3)

plt.savefig('./本周出班少于3天车型.png',dpi=600,bbox_inches = 'tight')



#本周与上周各车型出班趟次少于3天情况制表
data = data.loc[(data['周期出班趟次']<3)]
Bdata = Bdata.loc[(Bdata['周期出班趟次']<3)]
This_week_module_trips_low3 = data.groupby('车型')['周期出班趟次'].count()
Last_week_module_trips_low3 = Bdata.groupby('车型')['周期出班趟次'].count()
This_week_module_trips_low3 = pd.DataFrame(This_week_module_trips_low3).T
Last_week_module_trips_low3 = pd.DataFrame(Last_week_module_trips_low3).T
Module_tripVS_week_low3 = pd.concat([This_week_module_trips_low3, Last_week_module_trips_low3],axis = 0)
Module_tripVS_week_low3.columns.name = ''
Module_tripVS_week_low3.index = [ '本周', '上周']
Module_tripVS_week_low3.loc['变化幅度'] = Module_tripVS_week_low3.loc['本周'] - Module_tripVS_week_low3.loc['上周']
Module_tripVS_week_low3.index.name = '周期'
Module_tripVS_week_low3 = Module_tripVS_week_low3[['轻型车','3吨','5吨','8吨','12吨','牵引头']]

#本周与上周各车型公里数情况制图
#取做图数据
result4 = Module_tripVS_week_low3.T
CarModule = ['轻型车','3吨','5吨','8吨','12吨','牵引头']

x_CarModule   = range(len(CarModule ))

y1_cycle_trip_low3 =result4['本周']
y2_cycle_trip_low3 =result4['上周']
y3_cycle_trip_low3 =result4['变化幅度']
bar_width = 0.3

#设置图形大小
plt.rcParams['figure.figsize'] = (12.0,6.0) 
#绘图画板
fig_low3  = plt.figure()
ax1_Module_tripVS_week_low3 = fig_low3.add_subplot(111)

#设置标题
ax1_Module_tripVS_week_low3.set_title("本周与上周出班少于3天车型对比",fontsize='18')  
plt.bar(x=range(len(CarModule)), height=y1_cycle_trip_low3, label='本周', color='green', alpha=0.9, width=bar_width)
plt.bar(x=np.arange(len(CarModule)) + bar_width, height=y2_cycle_trip_low3, label='上周', color='orange', alpha=0.9, width=bar_width)
# 显示图例
plt.legend()

for x1_Module_tripVS_week_low3, yy_Module_tripVS_week_low3 in enumerate(y1_cycle_trip_low3):
    plt.text(x1_Module_tripVS_week_low3, yy_Module_tripVS_week_low3 +1  , '%.0f'%yy_Module_tripVS_week_low3, ha='center', va='bottom', fontsize=12, rotation=0)
for x1_Module_tripVS_week_low3, yy_Module_tripVS_week_low3 in enumerate(y2_cycle_trip_low3):
    plt.text(x1_Module_tripVS_week_low3 + bar_width, yy_Module_tripVS_week_low3 + 1, '%.0f'%yy_Module_tripVS_week_low3, ha='center', va='bottom', fontsize=12, rotation=0)

#画折线图
ax2_Module_tripVS_week_low3 = ax1_Module_tripVS_week_low3.twinx()  
ax2_Module_tripVS_week_low3.plot(x_CarModule ,y3_cycle_trip_low3, 'g',marker='.',ms=10,label='变化幅度')


# 显示图例

plt.xticks(x_CarModule,CarModule)
plt.legend(loc='right', fontsize=10,bbox_to_anchor=(0.6,0.94))


for x1, yy in enumerate(y3_cycle_trip_low3):
    plt.text(x1, yy + 0.2, '%.0f'%yy, ha='center', va='bottom', fontsize=15, rotation=0)

ax2_Module_tripVS_week_low3.grid(None)

plt.savefig('./本周与上周出班少于3天车型对比.png',dpi=600,bbox_inches = 'tight')


print('/n（三）出班小于3天车辆对比分析')
print('本周累计出班小于3天的车辆环比上周 %s 辆，%s' %  (Trip3_VS_week.T.loc['总计','差额'],100*Trip3_VS_week.T.loc['总计','差额']/Trip3_VS_week.T.loc['总计','上周']),'%，各单位组具体情况如下：')
print('区分公司%s '  % Trip3_VS_week.T.loc['区分公司合计','差额'])
print('中心局%s '  % Trip3_VS_week.T.loc['中心局合计','差额'])
print('其他单位%s '  % Trip3_VS_week.T.loc['其他单位合计','差额'])

#车型利用率
fig_data['车型计数'] = 1
Total_occupy_car_type = fig_data.groupby('车型')['车型计数'].count()
Actual_occupy_car_type  = fig_data.loc[fig_data['总行驶里程km']>1].groupby('车型')['车型计数'].count()
occupy_car_type = round(Actual_occupy_car_type/Total_occupy_car_type,3)
print('本周个车型利用率',occupy_car_type)

Before_week_fig_data['车型计数'] = 1
B_Total_occupy_car_type = Before_week_fig_data.groupby('车型')['车型计数'].count()
B_Actual_occupy_car_type  = Before_week_fig_data.loc[Before_week_fig_data['总行驶里程km']>1].groupby('车型')['车型计数'].count()
B_occupy_car_type = round(B_Actual_occupy_car_type/B_Total_occupy_car_type,3)
print('上周个车型利用率',B_occupy_car_type)
'''
看表分析
2、对本周出班趟次小于3天的车辆的统计，区分公司减少明显，最突出的是北京市朝阳区分公司。
3、本周轻型车的利用率为93.3%，其中区分公司占主体，说明本周各区分公司对车辆的使用进行了有效调整。
4、本周除牵引头外，其他车型出班率均大于90%，牵引头车辆利用率相比于上周减少3.4%，主要集中在运输单位。
'''

#静驶时长小于20%前10

fig_data['静驶时长占比全天运行时长比例'] = fig_data['全天静驶时长'] / fig_data['总运行时长（点火时长）hh']
print(fig_data.loc[(fig_data['静驶时长占比全天运行时长比例'] <  0.2) & (fig_data['静驶时长占比全天运行时长比例'] !=  0.2) ].sort_values(by = '静驶时长占比全天运行时长比例',ascending =False).head(11))
Top_10_stop = fig_data.loc[(fig_data['静驶时长占比全天运行时长比例'] <  0.2) & (fig_data['静驶时长占比全天运行时长比例'] !=  0.2) ].sort_values(by = '静驶时长占比全天运行时长比例',ascending =False).head(11)
Top_10_stop['静驶时长占比全天运行时长比例']  = Top_10_stop['静驶时长占比全天运行时长比例'].apply(lambda x :'%.2f%%' % (x*100))


Before_week_fig_data['静驶时长占比全天运行时长比例'] = Before_week_fig_data['全天静驶时长'] / Before_week_fig_data['总运行时长（点火时长）hh']
print(Before_week_fig_data.loc[(Before_week_fig_data['静驶时长占比全天运行时长比例'] <  0.2) & (Before_week_fig_data['静驶时长占比全天运行时长比例'] !=  0.2) ].sort_values(by = '静驶时长占比全天运行时长比例',ascending =False).head(11))
B_Top_10_stop = Before_week_fig_data.loc[(Before_week_fig_data['静驶时长占比全天运行时长比例'] <  0.2) & (Before_week_fig_data['静驶时长占比全天运行时长比例'] !=  0.2) ].sort_values(by = '静驶时长占比全天运行时长比例',ascending =False).head(11)
B_Top_10_stop['静驶时长占比全天运行时长比例']  = B_Top_10_stop['静驶时长占比全天运行时长比例'].apply(lambda x :'%.2f%%' % (x*100))


writer = pd.ExcelWriter('给ppt需要提取的表格.xlsx')
Distance50_VS_week.to_excel(writer,'本周行驶里程小于50公里车辆',index=True)
Trip3_VS_week.T.to_excel(writer,'本周出班天数小于3天的车辆',index=True)
Top_10_stop[['车牌号','配属单位','车辆名称','排放标准','登记年份','吨位','能源类型','总行驶里程km','总运行时长（点火时长）hh','全天静驶时长','静驶时长占比全天运行时长比例']].to_excel(writer,'本周重点车辆（静驶时间过长）',index=True)
B_Top_10_stop[['车牌号','配属单位','车辆名称','排放标准','登记年份','吨位','能源类型','总行驶里程km','总运行时长（点火时长）hh','全天静驶时长','静驶时长占比全天运行时长比例']].to_excel(writer,'上周重点车辆（静驶时间过长）',index=True)
writer.save()




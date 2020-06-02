# _*_ coding.utf-8 _*_
# Editor: Ziqi LIU
# Date: 2020/6/2
# File name: StepTwo.py
# Tool: PyCharm

import pandas as pd
df = pd.read_excel('D:\pycharm\Deloitte\Step1.xlsx')
df1 = pd.DataFrame({"主体新预警":df["债券新预警"].groupby(df['主体']).min(),
                   "原预警时间":df["原预警时间"].groupby(df['主体']).min()})
df1['差值'] =pd.to_datetime(df1['主体新预警'])-pd.to_datetime(df1['原预警时间'])
# 输出汇总表格
print(df1)

# Result 保存全部
writer = pd.ExcelWriter('D:\pycharm\Deloitte\Question 1 result.xlsx')
df.to_excel(writer)
df1.to_excel(writer, startcol=11)
writer.save()
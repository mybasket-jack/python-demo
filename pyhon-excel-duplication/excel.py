import pandas as pd
import os
import datetime

rootpath = os.getcwd() # 获取当前路径
timeStr = str(datetime.datetime.now().strftime('%Y-%m-%d-%H%M%S')) #保存文件的时间撮

startTime = datetime.datetime.now()  #开始时间

excel_dir = rootpath+'\\excel'  # 创建一个excel的文件夹，把要去重的文件放在excel里面
os.chdir(excel_dir) # 切换到excel路径
li = []
for i in os.listdir(excel_dir):
    li.append(pd.read_excel(i))
    print('合并: '+i)
# 合并完成后的excel，是放在excel文件夹下的
combine_file = rootpath+'\\bak.xlsx'
writer = pd.ExcelWriter(combine_file)
pd.concat(li).to_excel(writer,'Sheet1',index=False)
writer.save()
print('文件合并完成,开始解析....')

# 进行解析
df = pd.read_excel(rootpath+'\\bak.xlsx')
df['is_duplicated'] = df.duplicated(['序号'])  #设置重复列
counts = df['序号'].value_counts()  #统计重复的次数
dupNum = 3  # 出现次数
filterArray = (counts[counts > dupNum].index).values #符合条件的列值数组
print(filterArray)
print(counts[counts > dupNum]) # 显示各列值符合条件的情况
df_dup = df.loc[(df['is_duplicated'] == True) & (df['序号'].isin(filterArray) == True)]  # 获取重复3次的列


result_file = rootpath+'\\result\\'+'out_'+timeStr+'.xlsx' #解析生成的文件
df_dup.to_excel(result_file)
os.remove(combine_file) # 移除合并的文件
print('解析完成,解析后的文件： '+result_file)

endTime = datetime.datetime.now() #结束时间
print('耗时：'+str((endTime-startTime).seconds)+'秒')


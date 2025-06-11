"""
Created on Mon Nov 18 17:14:59 2019
coding: utf-8 
"""

#-----------Preparation-----------
# 1、Use Excel to create a static pool table 
# 2. Check the date time format
# 3. Change '拖欠91-120天金额' to "91-120"
# 4. Transform the '-' to an emptyspace
# 5. Change the header of the table to the same format: 
##其中一种表头为：静态池（改），报告期末（改），期初笔数（改），期初金额（改），早偿金额，正常贷款金额，拖欠1-30天，拖欠31-60天，拖欠61-90天，91-120（改），拖欠120天以上
##另一种表头为：静态池（改），报告期末（改），期初笔数（改），期初金额（改），部分早偿（改），全部早偿（改），正常贷款金额，拖欠1-30天，拖欠31-60天，拖欠61-90天，91-120（改），拖欠120天以上


import pandas as pd
import numpy as np
import math


#export the excel file
writer = pd.ExcelWriter("D:\\ABS Files\\output111.xlsx")
#import the excel file
df0 = pd.read_excel("D:\\ABS Files\\111.xlsx", sheet_name="Table1")

RowNum1 = df0.shape[0]

guanceqishu=[]
xinzengdaikuan=[]
qichuzhengchang=[]
buzhengchang=[]
CPR=[]

##############################Processing the initial static pool data###############################
# Add the watching period number, (the current cash flow is the nth period of the static pool)
for i in range(RowNum1):
    if df0.iloc[i,0]==df0.iloc[i,1]:
        guanceqishu.append(0)
    else:
        guanceqishu.append(guanceqishu[i-1]+1)

df0.insert(2,'观测期数',guanceqishu)

# Add the new loan amount, which is the initial capital of the static pool
df1=pd.pivot_table(df0,index=["静态池"],values=["期初金额"],
               aggfunc=[np.max])

df0 = pd.merge(df0, df1, how='left', on='静态池')
for i in range(RowNum1):
    xinzengdaikuan.append(df0.iloc[i,-1])
    
df0.insert(4,'新增贷款额',xinzengdaikuan)

         
# Add the amount of normal loans at the beginning of the period
for i in range(RowNum1):
    # ！！！！Substract the abnormal repayment amount
    for j in range(-6,-1):
        if math.isnan(df0.iloc[i,j])==True:
            df0.iloc[i,j]=0
   # if the abnormal repayment amount = nan, replace it with 0, then sum up
    buzhengchang.append(df0.iloc[i,-6]+df0.iloc[i,-5]+df0.iloc[i,-4]+df0.iloc[i,-3]+df0.iloc[i,-2])
   
    if df0.iloc[i,2]==0:
        qichuzhengchang.append(df0.iloc[i,5]/df0.iloc[i,4])
    else:
        qichuzhengchang.append((df0.iloc[i,5]-buzhengchang[i])/df0.iloc[i,4]) 
   
df0.insert(6,'期初金额（正常类）占比',qichuzhengchang)


for i in range(RowNum1):
    #replace the '0' with 'nan'
    for j in range(-6,-1):
        if df0.iloc[i,j]==0:
            df0.iloc[i,j]=np.nan


#add CPR 
if '部分早偿' in df0.columns:
    monthsum=[]
    for i in range(RowNum1):
        if math.isnan(df0.iloc[i,7])==True and math.isnan(df0.iloc[i,8])==False:
            monthsum.append(df0.iloc[i,8])
        elif math.isnan(df0.iloc[i,7])==False and math.isnan(df0.iloc[i,8])==True:
            monthsum.append(df0.iloc[i,7])
        elif math.isnan(df0.iloc[i,8])==False and math.isnan(df0.iloc[i,7])==False:
            monthsum.append(df0.iloc[i,8]+df0.iloc[i,7])
        else:
            monthsum.append(np.nan)
        # CPR.append(1-(1-monthsum[i]/df0.iloc[i,5])**12)
        CPR.append(monthsum[i] / df0.iloc[i, 5])
    df0.insert(9,'月早偿总额',monthsum)    
    df0.insert(10,'CPR',CPR)
else:
# in case there is no partial prepayment amount
    for i in range(RowNum1):
        # CPR.append(1-(1-df0.iloc[i,7]/df0.iloc[i,5])**12)
        CPR.append(df0.iloc[i, 7] / df0.iloc[i, 5])
    df0.insert(8,'CPR',CPR)
            

df0.to_excel(writer,sheet_name='静态池')


################################## Add a sheet 'analysis of prepayment rate' ############################
df2=pd.pivot_table(df0,index=["静态池"],values=["CPR"],
               columns=["观测期数"],aggfunc=[np.sum])
df2.loc['average'] = np.nan
RowNum2 = df2.shape[1]
for i in range(RowNum2):
    df2.loc['average'].iloc[i]=df2.iloc[:,i].mean()

df2.to_excel(writer,sheet_name='早偿率分析')

#######################################Add a sheet 'Remaining initial capital at the beginning of the month' #####################
df3=pd.pivot_table(df0,index=["观测期数"],values=["期初金额（正常类）占比"],
               columns=["静态池"],aggfunc=[np.sum])
RowNum3 = df3.shape[0]

# Add the average value to the first column
average=[]
for i in range(RowNum3):
    average.append(df3.iloc[i,:].mean())
df3.insert(0,'平均值',average)

df3.to_excel(writer,sheet_name='月初本金余额（正常类）')

##################################### Add a sheet 'overdue rate (>90 days)' #########################
df4=pd.pivot_table(df0,index=["观测期数"],values=["91-120"],
               columns=["静态池"],aggfunc=[np.sum])
df5=pd.pivot_table(df0,values=["期初金额"],
               columns=["静态池"],aggfunc=[np.max])
df4.to_excel(writer,sheet_name='90天以上拖欠率')
df5.to_excel(writer,sheet_name='期初金额最大值')

RowNum4 = df4.shape[0] #row no.
RowNum5=df4.shape[1] #column no.

df6=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(RowNum5):
        if math.isnan(df4.iloc[i,j])==True:
            df6.iloc[i,j]=np.nan
        else:
            df6.iloc[i,j]=df4.iloc[i,j]/df5.iloc[0,j]
# in the above sheet, if there is a data, replace it with '1', if no data, replace it with '0'.
df7=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(RowNum5):
        if math.isnan(df6.iloc[i,j])==True:
            df7.iloc[i,j]=0
        else:
            df7.iloc[i,j]=1
df6.to_excel(writer,sheet_name='增量拖欠')
df7.to_excel(writer,sheet_name='增量拖欠0-1')

# Add arithemetic mean and weighted average of incremental overdue
df8=pd.DataFrame(index=np.arange(RowNum4),columns=['算术平均增量','加权平均增量','违约时间分布'])    
for i in range(RowNum4):
    df8.iloc[i].loc['算术平均增量']=df6.iloc[i,:].mean()
for i in range(RowNum4):
    for j in range(RowNum5):
        if math.isnan(df6.iloc[i,j])==True:
            df6.iloc[i,j]=0
    df8.iloc[i].loc['加权平均增量']=np.sum(np.array(df5.iloc[0,:])*np.array(df6.iloc[i,:]))/np.sum(np.array(df5.iloc[0,:])*np.array(df7.iloc[i,:]))    
for i in range(RowNum4):
    df8.iloc[i].loc['违约时间分布']=df8.iloc[i].loc['算术平均增量']/df8.loc[:,'算术平均增量'].sum()
df8.loc['sum'] = np.nan
df8.loc['sum'].iloc[0]=df8.iloc[:,0].sum()
df8.loc['sum'].iloc[1]=df8.iloc[:,1].sum()
df8.to_excel(writer,sheet_name='增量拖欠基准')

# add the cumulative overdue amount  
df9=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(RowNum5):
        df9.iloc[0,j]=df6.iloc[0,j]
        if df7.iloc[i,j]==0:
            df9.iloc[i,j]=np.nan
        else:
            df9.iloc[i,j]=df6.iloc[i,j]+df9.iloc[i-1,j]           

df10=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(RowNum5):
        if math.isnan(df9.iloc[i,j])==True:
            df10.iloc[i,j]=df10.iloc[i-1,j]+df8.iloc[i].loc['加权平均增量']
        else:
            df10.iloc[i,j]=df9.iloc[i,j]    
            
df9.to_excel(writer,sheet_name='累计拖欠')
df10.to_excel(writer,sheet_name='累计拖欠（平均增量补足）')
            
# add the arithemetic mean to the cumulative overdue amount
df11=pd.DataFrame(index=np.arange(RowNum4),columns=['算术平均增量','加权平均增量'])    
sum=np.sum(np.array(df5.iloc[0,:]))
for i in range(RowNum4):
    df11.iloc[i].loc['算术平均增量']=df10.iloc[i,:].mean()
    df11.iloc[i].loc['加权平均增量']=np.sum(np.array(df5.iloc[0,:])*np.array(df10.iloc[i,:]))/sum
df11.to_excel(writer,sheet_name='累计拖欠基准')

#########################Add a sheet 'Incremental Default Rate by Aging Bucket' #############################
df12=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(1,RowNum5+1):
        if i+j>RowNum4:
            df12.loc[i,j]=0
        else:
            df12.loc[i,j]=df8.iloc[i+j-1].loc['算术平均增量']
df12.to_excel(writer,sheet_name='匹配账龄违约率增量')

############################ Add a sheet 'Incremental Default Rate Matched to Beginning-of-Month Principal' ######################
df13=pd.DataFrame(index=np.arange(RowNum4),columns=np.arange(1,RowNum5+1))
for i in range(RowNum4):
    for j in range(1,RowNum5+1):
        if df3.iloc[i,0]!=0:
            df13.loc[i,j]=df12.loc[i,j]/df3.iloc[i,0]
        else:
            df13.loc[i,j]=0
SUM=[]
for i in range(RowNum4):
    SUM.append(df13.iloc[i,:].sum())
df13.insert(0,'SUM',SUM)
df13.to_excel(writer,sheet_name='匹配月初本金后违约率增量')

############################ Add a sheet 'Cumulative Default and Prepayment Rate Summary Table' ######################
try:
    df14 = pd.DataFrame(columns=['静态池', '平均早偿率', '累计违约率不补数', '累计违约率补数'])
    statistic_pools = list(set(df0['静态池'].to_list()))
    statistic_pools.sort()
    for i in range(len(statistic_pools)):
        np_prepayment_list = df2.iloc[i, :].to_list()
        sum_smm = 0
        count_smm = 0
        for smm_i in np_prepayment_list:
            if math.isnan(smm_i) == 0:
                count_smm = count_smm + 1
                sum_smm = sum_smm + smm_i
        if count_smm > 0:
            average_SMM = sum_smm / count_smm
        else:
            average_SMM = 0
        dr_or = max(df9.iloc[:, i])
        dr_add = max(df10.iloc[:, i])
        add_row_i = pd.DataFrame([[statistic_pools[i], average_SMM, dr_or, dr_add]], columns=df14.columns)
        df14 = pd.concat([df14, add_row_i], axis=0, sort=False)
    df14.to_excel(writer, sheet_name='统计数据', index=False)
except:
    print("统计错误！") # ERROR

writer.close()
#The export sheet is 'output_excel.xlsx'



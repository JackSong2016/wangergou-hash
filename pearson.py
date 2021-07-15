import xlrd
import numpy as np
import pandas as pd
from scipy import stats
from scipy.stats import pearsonr
import matplotlib as plt

#从excel文件中读取数据
def read(file):
    wb = xlrd.open_workbook(filename=file)#打开文件
    sheet = wb.sheet_by_index(0)#通过索引获取表格
    rows = sheet.nrows # 获取行数
    all_content = []        #存放读取的数据
    for j in range(0, 11):       #取第1~第11列的数据
        temp = []
        for i in range(1,rows) :
            cell = sheet.cell_value(i, j)   #获取数据
            temp.append(cell)
        all_content.append(temp)    #按列添加到结果集中
        temp = []
    return np.array(all_content)

#统计描述
def calculate(datas):
    MIN = np.min(datas,axis = 1)    #计算最小值
    MAX = np.max(datas,axis = 1)    #计算最大值
    AVG = np.average(datas,axis = 1)    #计算平均值
    MEDIAN = np.median(datas,axis = 1)  #计算中位数
    SKEWNESS =stats.skew(datas,axis = 1)    #计算偏度
    KURTOSIS = stats.kurtosis(datas,axis = 1)   #计算峰度
    STD = np.std(datas,axis = 1)    #计算标准差
    result = np.array([MIN,MAX,AVG,MEDIAN,SKEWNESS,KURTOSIS,STD])   #形成一个矩阵
    return result

#将统计描述输出到excel文件中
def write(answer_data):
    writer = pd.ExcelWriter('D:\\pycharm\\project\\example\\pearson.xlsx')       # 写入Excel文件
    answer_data.to_excel(writer, 'pearson', float_format='%.5f')     # ‘pearson’是写入excel的sheet名
    writer.save()
    writer.close()

datas=read('D:\\pycharm\\project\\example\\HDH.xlsx')
#result = calculate(datas)   #统计描述
#corre=datas.corr(method='pearson',periods=1)
#方法选择person相关性
corre=np.corrcoef(datas)
#print(corre)
# for i in range(11):
#     list=corre[:,i]
#     print(list[0])
# plt.plot(figsize=(20,12))
# plt.show()
#zidian={"Sr":list[0],"Tamb":list[1],"Uair":list[2],"Twci":list[3],"Twce":list[4],"Twhi":list[5],"Twhe":list[6],"Taci":list[7],"Tace":list[8],"Mair":list[9],"Output":list[10],}
df=pd.DataFrame(corre,index=['Sr','Tamb','Uair','Twci','Twce','Twhi','Twhe','Taci','Tace','Mair','Output'])
#corrcoe = np.corrcoef(result)   #计算皮尔逊相关系数
#answer_data = pd.DataFrame(zidian,index=[Sr,Tamb,Uair,Twci,Twce,Twhi,Twhe,Taci,Tace,Mair,Output])        #将ndarry转换为DataFrame
write(df)  #输出结果
print(df)
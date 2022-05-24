import scipy.stats as ss
import numpy as np
import os
import xlrd
from openpyxl import Workbook
import matplotlib.pyplot as plt
import seaborn as sns


def autoNorm(data):         # 传入一个矩阵
    mins = data.min(0)      # 返回data矩阵中每一列中最小的元素，返回一个列表
    maxs = data.max(0)      # 返回data矩阵中每一列中最大的元素，返回一个列表
    ranges = maxs - mins    # 最大值列表 - 最小值列表 = 差值列表
    normData = np.zeros(np.shape(data))     # 生成一个与 data矩阵同规格的normData全0矩阵，用于装归一化后的数据
    row = data.shape[0]                     # 返回 data矩阵的行数
    normData = data - np.tile(mins, (row, 1))   # data矩阵每一列数据都减去每一列的最小值
    normData = normData / np.tile(ranges, (row, 1))   # data矩阵每一列数据都除去每一列的差值（差值 = 某列的最大值- 某列最小值）
    return normData


# 读取Excel文件
# 男生 BMI	肺活量	1000米	坐位体前屈	立定跳远	50米	引体向上
# 女生 BMI	肺活量	800米	坐位体前屈	立定跳远	50米跑	一分钟仰卧起坐
data_path = r'../IJERPH/女生'
filenames = os.listdir(data_path)
#y_ticklabels = ['BMI',	'vital capacity',	'1000m',	'sit & reach',	'standing '+'\n' +'long jump',	'50m',	'pull up', 'VCWI']
#x_ticklabels = ['BMI',	'vital capacity',	'1000m',	'sit & reach',	'standing '+'\n' +'long jump',	'50m',	'pull up', 'VCWI']
y_ticklabels = ['BMI',	'vital capacity',	'800m',	'sit & reach',	'standing '+'\n' +'long jump',	'50m',	'one minute'+'\n'+'sit ups', 'VCWI']
x_ticklabels = ['BMI',	'vital capacity',	'800m',	'sit & reach',	'standing '+'\n' +'long jump',	'50m',	'one minute '+'\n'+'sit ups', 'VCWI']
font1 = {'family': 'Times New Roman', 'size': 20}


def preprocess(filenames=filenames):
    for file in filenames:
        print(file)
        workbook = Workbook()
        worksheet = workbook.active  # 每个workbook创建后，默认会存在一个worksheet，对默认的worksheet进行重命名
        worksheet.title = "Sheet1"

        xlsx = xlrd.open_workbook(data_path + '/' + file)
        print('All sheets:%s' % xlsx.sheet_names())
        sheet1 = xlsx.sheets()[0]  #  获得第1张sheet，索引从0开始
        sheet1_name = sheet1.name  #  获得名称
        sheet1_cols = sheet1.ncols  #  获得列数
        sheet1_n_rows = sheet1.nrows  #  获得行数

        x = []
        for n in range(sheet1_n_rows):  #  逐行遍历sheet1数据

            n = np.array(sheet1.row_values(n))
            I = n.astype(np.float64)
            x1 = I[:][:].tolist()

            x.append(x1)
        print(len(x))
    return x


x = preprocess()
x = np.array(x)
for i in range(len(x[0])):      # 把x的每一列归一化
    x[:, i] = autoNorm(x[:, i])[0]


corr = []
for n in range(len(x[0])):
    corr_ = []
    for i in range(len(x[0])):
        X = x[:, n]
        Y = x[:, i]
        result = ss.pearsonr(X, Y)
        corr_.append(result[0])  # result[0]是相关关系数值
    corr.append(corr_)

corr = np.array(corr)

figure, ax = plt.subplots(figsize=(18, 14))
sns.heatmap(corr, square=True, annot=True, ax=ax,
            xticklabels=x_ticklabels, yticklabels=y_ticklabels,
            linewidths=1, linecolor='w', annot_kws=font1)
cax = plt.gcf().axes[-1]
cax.tick_params(labelsize=20)
plt.tick_params(labelsize=20)
plt.xticks(rotation=-30)
plt.yticks(rotation=0)
plt.show()


corr2 = []
for n in range(len(x[0])):
    corr_2 = []
    for i in range(len(x[0])):
        X = x[:, n]
        Y = x[:, i]
        result = ss.pearsonr(X, Y)
        if result[1] < 0.001:  # result[1]是两者不相关的概率，即p值
            corr_2.append(1)
        elif 0.001 <= result[1] < 0.01:
            corr_2.append(0.7)
        elif 0.01 <= result[1] < 0.05:
            corr_2.append(0.5)
        else:
            corr_2.append(0)
    corr2.append(corr_2)

corr2 = np.array(corr2)

figure, ax = plt.subplots(figsize=(8, 7))
sns.heatmap(corr2, square=True, annot=False, ax=ax,
            yticklabels=y_ticklabels, xticklabels=x_ticklabels,
            linewidths=1, linecolor='b')
plt.tick_params(labelsize=9)
plt.xticks(rotation=-30)
plt.yticks(rotation=0)
plt.show()


corr3 = []
for n in range(len(x[0])):
    corr_3 = []
    for i in range(len(x[0])):
        X = x[:, n]
        Y = x[:, i]
        result = ss.pearsonr(X, Y)
        corr_3.append(result[1])  # result[1]是两者不相关的概率，即p值
    corr3.append(corr_3)

corr3 = np.array(corr3)

figure, ax = plt.subplots(figsize=(18, 14))

sns.heatmap(corr3, square=True, annot=True, ax=ax,
            xticklabels=x_ticklabels, yticklabels=y_ticklabels,
            linewidths=1, linecolor='w', annot_kws=font1)

plt.tick_params(labelsize=18)
plt.xticks(rotation=-30)
plt.yticks(rotation=0)
plt.show()

print('pearsonr:', corr3)

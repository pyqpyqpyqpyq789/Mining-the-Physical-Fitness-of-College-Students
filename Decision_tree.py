from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
import os
import xlrd
from openpyxl import Workbook
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
import numpy as np
from sklearn.metrics import mean_squared_error


def autoNorm(data):         # 传入一个矩阵
    mins = data.min(0)      # 返回data矩阵中每一列中最小的元素，返回一个列表
    maxs = data.max(0)      # 返回data矩阵中每一列中最大的元素，返回一个列表
    ranges = maxs - mins    # 最大值列表 - 最小值列表 = 差值列表
    normData = np.zeros(np.shape(data))     # 生成一个与 data矩阵同规格的normData全0矩阵，用于装归一化后的数据
    row = data.shape[0]                     # 返回 data矩阵的行数
    normData = data - np.tile(mins, (row, 1))   # data矩阵每一列数据都减去每一列的最小值
    normData = normData / np.tile(ranges, (row, 1))   # data矩阵每一列数据都除去每一列的差值（差值 = 某列的最大值- 某列最小值）
    return normData


data_path = r'../IJERPH/神经网络_女'
filenames = os.listdir(data_path)

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

    x, y = [], []
    for n in range(sheet1_n_rows):  #  逐行遍历sheet1数据

        n = np.array(sheet1.row_values(n))
        I = n.astype(np.float64)
        x1 = I[:][:-1]
        y1 = [I[:][-1]]
        x.append(x1)
        y.append(y1)

y2 = []
for i in y:
    y2.append(i[0])
y2 = np.array(y2)
t = np.arange(len(y2))
'''
plt.scatter(t, y2, s=10, c='b', alpha=0.5)
plt.show()
plt.close()
print(x)
'''
#   归一化
x = np.array(x)
for i in range(len(x[0])):      # 把x的每一列归一化
    x[:, i] = autoNorm(x[:, i])[0]
# print(x)

train_y = y2[:780]  # 训练集测试集划分
train_x = x[:780]
val_y = y2[780:]
val_x = x[780:]
DT = DecisionTreeRegressor(max_depth=4)
DT.fit(train_x, train_y)
pred = DT.predict(val_x)
#可视化决策树
from dtreeviz.trees import *
os.environ["PATH"] += os.pathsep + 'C:/Program Files/Graphviz/bin/'  #千万要有这句
feature_names = ['BMI',	'vital capacity',	'sit & reach',	'standing '+'\n' +'long jump',	'50m',	'pull up']
viz = dtreeviz(DT, x, y2, target_name='time', feature_names=feature_names, )
viz.view()

# 可视化决策树预测结果
plt.figure(figsize=(10, 5))
plt.scatter(np.arange(0, len(val_x), 1), val_y, c='b', s=2)
plt.scatter(np.arange(0, len(val_x), 1), pred, c='r', s=2)
plt.show()

loss = mean_squared_error(pred, val_y)
print("Loss=", loss)
mean_diff = sum(abs(val_y - pred)) / len(val_y)
print('mean_diff=', mean_diff)

# 随机森林
RF = RandomForestRegressor(n_estimators=100, random_state=0)
RF.fit(train_x, train_y)
pred2 = RF.predict(val_x)

# 可视化随机森林预测结果
plt.figure(figsize=(10, 5))
#plt.title("Performance prediction of 1000m running", fontsize='15')
plt.xlabel('Participant serial number', fontsize='15')
#plt.ylabel("Prediction & ture value", fontsize="15")
plt.ylabel("Time (second)", fontsize="15")
plt.scatter(np.arange(0, len(train_x), 1), train_y, c='b', s=4, alpha=0.5)
plt.scatter(np.arange(len(train_x), len(x), 1), val_y, c='r', s=4, marker='o', alpha=0.6)

plt.scatter(np.arange(0, len(train_x), 1), RF.predict(train_x), c='y', s=4, alpha=1, marker='x')
plt.scatter(np.arange(len(train_x), len(x), 1), pred2, c='g', s=4, alpha=0.5, marker='v')
plt.legend(['train true', 'validation true', 'train pred', 'validation pred'], frameon=True, loc="upper right", markerscale=3)
plt.show()

##
plt.figure(figsize=(10, 10))
fig, ax = plt.subplots()
plt.xlabel('True', fontsize='15')
plt.ylabel('Pred', fontsize="15")
#plt.xticks(ticks=range(100, 600, 20))
#plt.yticks(ticks=range(100, 600, 20))
ax.set_xlim(xmin=5, xmax=18)
ax.set_ylim(ymin=5, ymax=18)
plt.scatter(val_y, pred2, c='r', s=1, alpha=1, marker='x')
plt.plot([0, 400], [0, 400])
plt.show()

loss2 = mean_squared_error(pred2, val_y)
print("Loss2=", loss2)
mean_diff2 = sum(abs(val_y - pred2)) / len(val_y)
print('mean_diff2=', mean_diff2)

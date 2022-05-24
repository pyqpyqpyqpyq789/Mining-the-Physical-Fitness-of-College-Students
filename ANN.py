'''
https://blog.csdn.net/qq_39567427/article/details/111935833?utm_medium=distribute.pc_relevant.none-task-blog-2%7Edefault%7ECTRLIST%7Edefault-4.no_search_link&depth_1-utm_source=distribute.pc_relevant.none-task-blog-2%7Edefault%7ECTRLIST%7Edefault-4.no_search_link
'''
import os
import xlrd
from openpyxl import Workbook
import torch
from torch import nn
import matplotlib.pyplot as plt
plt.rcParams['font.sans-serif'] = ['SimHei']  # 用来正常显示中文标签
import numpy as np

from sklearn.preprocessing import scale
from torchsummary import summary


device = torch.device("cuda:0" if torch.cuda.is_available() else "cpu")

data_path = r'../IJERPH/神经网络_男'
filenames = os.listdir(data_path)

for file in filenames:
    print(file)

    workbook = Workbook()
    worksheet = workbook.active  # 每个workbook创建后，默认会存在一个worksheet，对默认的worksheet进行重命名
    worksheet.title = "Sheet1"

    xlsx = xlrd.open_workbook(data_path+'/'+file)
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
plt.scatter(t, y2, s=1, c='b', alpha=0.5)
plt.show()
plt.close()
print('y', x)


class Net(nn.Module):
    def __init__(self, input_num, hidden_num1, hidden_num2, output_num):
        super(Net, self).__init__()
        self.net = nn.Sequential(
            nn.Linear(input_num, hidden_num1),
            nn.ReLU(),
            nn.Linear(hidden_num1, hidden_num2),
            nn.ReLU(),
            nn.Linear(hidden_num2, output_num),
            nn.ReLU()
        )

    def forward(self, inputs):
        return self.net(inputs)


net = Net(input_num=6, hidden_num1=12, hidden_num2=12, output_num=1).to(device)
summary(net, input_size=(6,))
print(net)  # 打印网络结构

x_data, y_data = torch.FloatTensor(scale(x)).to(device), torch.unsqueeze(torch.FloatTensor(y2), dim=1).to(device)
train_x, train_y = x_data[:780, :], y_data[:780]
valid_x, valid_y = x_data[780:, :], y_data[780:]

epochs = 200
learning_rate = 0.005
batch_size = 30
total_step = int(train_x.shape[0] / batch_size)

optimizer = torch.optim.Adam(net.parameters(), lr=learning_rate)
#optimizer = torch.optim.SGD(net.parameters(), lr=learning_rate, momentum=0.9)
loss_func = torch.nn.MSELoss()


def weight_reset(m):
    if isinstance(m, nn.Conv2d) or isinstance(m, nn.Linear):
        m.reset_parameters()


net.apply(weight_reset)

epoch_train_loss_value = []
step_train_loss_value = []
epoch_valid_loss_value = []
pre_list = []

for i in range(epochs):
    pre_list.clear()
    for step in range(total_step):
        xs = train_x[step * batch_size:(step + 1) * batch_size, :]
        ys = train_y[step * batch_size:(step + 1) * batch_size]
        prediction = net(xs)

        loss = loss_func(prediction, ys)
        optimizer.zero_grad()
        loss.backward()
        optimizer.step()
        step_train_loss_value.append(loss.cpu().detach().numpy().item())
        pre_list.append(prediction.cpu().detach().numpy())

    val_pred = net(valid_x).cpu().detach().numpy()
    valid_loss = loss_func(net(valid_x), valid_y)
    epoch_valid_loss_value.append(valid_loss)
    epoch_train_loss_value.append(np.mean(step_train_loss_value))
    print('epoch={:3d}/{:3d}, train_loss={:.4f}, valid_loss={:.4f}'.format(i + 1, epochs, np.mean(step_train_loss_value), valid_loss))

    pre_list2 = []
    for pre1 in pre_list:
        for pre2 in pre1:
            pre_list2.append(pre2)
    print(len(pre_list2))

    plt.figure(figsize=(10, 5))
    plt.title("Performance prediction of 1000m running", fontsize='15')
    plt.xlabel('Participant serial number', fontsize='15')
    plt.ylabel("Time (second)", fontsize="15")
    plt.scatter(np.arange(0, len(train_y), 1), train_y.cpu().numpy(), c='b', s=4, marker='o', alpha=0.4)
    plt.scatter(np.arange(len(train_y), len(y2), 1), valid_y.cpu().numpy(), c='r', s=4, marker='o', alpha=0.6)
    plt.scatter(np.arange(0, len(pre_list2), 1), pre_list2, c='y', s=4, marker='x', alpha=1)

    plt.scatter(np.arange(len(pre_list2), len(t), 1), val_pred, c='g', s=4, marker='v', alpha=1)
    plt.legend(['train true', 'validation true', 'train pred', 'validation pred'], frameon=True, loc="upper right", markerscale=3)
    plt.savefig("../IJERPH/img_output/" + str(i) + '.jpg')
    plt.show()
    plt.close()

    mean_diff2 = sum(abs(valid_y.cpu().numpy() - val_pred)) / len(valid_y)
    print('mean_diff2=', mean_diff2)


plt.figure(figsize=(10, 5))
fig = plt.gcf()
fig.set_size_inches(10, 5)

plt.xlabel('Epochs', fontsize=15)
plt.ylabel('MSE Loss', fontsize=15)
plt.plot(epoch_train_loss_value, c='blue', label='Train loss')
plt.plot(epoch_valid_loss_value, c='red', label='Validation loss')
plt.legend(loc='best')
plt.title('Training and Validation loss', fontsize=15)
plt.savefig("../IJERPH/img_output/" + 'loss_.jpg')
plt.show()


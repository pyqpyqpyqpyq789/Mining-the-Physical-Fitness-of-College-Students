import os
from openpyxl import Workbook
import xlrd


# coding=utf-8
def createC1(dataSet):  # 构建所有1项候选项集的集合
    C1 = []
    for transaction in dataSet:
        for item in transaction:
            if [item] not in C1:
                C1.append([item])  # C1添加的是列表，对于每一项进行添加，[[1], [2], [3], [4], [5]]
    # print('C1:',C1)
    return list(map(frozenset, C1))  # 使用frozenset，被“冰冻”的集合，为后续建立字典key-value使用。


###由候选项集生成符合最小支持度的项集L。参数分别为数据集、候选项集列表，最小支持度
###如
###C3: [frozenset({1, 2, 3}), frozenset({1, 3, 5}), frozenset({2, 3, 5})]
###L3: [frozenset({2, 3, 5})]
def scanD(D, Ck, minSupport):
    ssCnt = {}
    for tid in D:  # 对于数据集里的每一条记录
        for can in Ck:  # 每个候选项集can
            if can.issubset(tid):  # 若是候选集can是作为记录的子集，那么其值+1,对其计数
                if not ssCnt.__contains__(can):  # ssCnt[can] = ssCnt.get(can,0)+1一句可破，没有的时候为0,加上1,有的时候用get取出，加1
                    ssCnt[can] = 1
                else:
                    ssCnt[can] += 1
    numItems = float(len(D))
    retList = []
    supportData = {}
    for key in ssCnt:
        support = ssCnt[key] / numItems  # 除以总的记录条数，即为其支持度
        if support >= minSupport:
            retList.insert(0, key)  # 超过最小支持度的项集，将其记录下来。
            supportData[key] = support
    return retList, supportData


###由Lk生成K项候选集Ck
###如由L2: [frozenset({3, 5}), frozenset({2, 5}), frozenset({2, 3}), frozenset({1, 3})]
###生成
###C3: [frozenset({1, 2, 3}), frozenset({1, 3, 5}), frozenset({2, 3, 5})]
def aprioriGen(Lk, k):
    retList = []
    lenLk = len(Lk)
    for i in range(lenLk):
        for j in range(i + 1, lenLk):
            if len(Lk[i] | Lk[j]) == k:
                retList.append(Lk[i] | Lk[j])
    return list(set(retList))


####生成所有频繁子集
def apriori(dataSet, minSupport=0.4):
    C1 = createC1(dataSet)
    D = list(map(set, dataSet))
    L1, supportData = scanD(D, C1, minSupport)
    L = [L1]  # L将包含满足最小支持度，即经过筛选的所有频繁n项集，这里添加频繁1项集
    k = 2
    while (len(L[k - 2]) > 0):  # k=2开始，由频繁1项集生成频繁2项集，直到下一个打的项集为空
        Ck = aprioriGen(L[k - 2], k)
        Lk, supK = scanD(D, Ck, minSupport)
        supportData.update(supK)  # supportData为字典，存放每个项集的支持度，并以更新的方式加入新的supK
        L.append(Lk)
        k += 1
    return L, supportData


data_path = r'../IJERPH/男生'
filenames = os.listdir(data_path)


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
            n = sheet1.row_values(n)
            for i in range(len(n)):
                n[i] = float(n[i])

            if n[0] <= 17.1:
                n_0 = "低体重"
            if 17.1 < n[0] <= 23.9:
                n_0 = "正常"
            if 23.9 < n[0] <= 28.0:
                n_0 = "超重"
            if 28.0 < n[0]:
                n_0 = "肥胖"
            n[0] = 'BMI' + n_0

            if n[1] >= 3300:
                n1 = '优秀'
            if 3300 > n[1] >= 3000:
                n1 = '良好'
            if 3000 > n[1] >= 2000:
                n1 = '及格'
            if 2000 > n[1]:
                n1 = '不及格'
            n[1] = '肺活量' + n1

            if n[2] <= 207:
                n2 = '优秀'
            if 207 < n[2] <= 222:
                n2 = '良好'
            if 222 < n[2] <= 272:
                n2 = '及格'
            if 272 < n[2]:
                n2 = '不及格'
            n[2] = '1000m' + n2

            if n[3] >= 21.3:
                n3 = '优秀'
            if 21.3 > n[3] >= 17.7:
                n3 = '良好'
            if 17.7 > n[3] >= 3.7:
                n3 = '及格'
            if 3.7 > n[3]:
                n3 = "不及格"
            n[3] = '体前屈' + n3

            if 263 <= n[4]:
                n4 = '优秀'
            if 263 > n[4] >= 248:
                n4 = "良好"
            if 248 > n[4] >= 208:
                n4 = '及格'
            if n[4] < 208:
                n4 = '不及格'
            n[4] = '立定跳远' + n4

            if 6.9 >= n[5]:
                n5 = '优秀'
            if 6.9 < n[5] <= 7.1:
                n5 = '良好'
            if 7.1 < n[5] <= 9.1:
                n5 = '及格'
            if 9.1 < n[5]:
                n5 = '不及格'
            n[5] = '50m' + n5

            if 17 <= n[6]:
                n6 = '优秀'
            if 17 > n[6] >= 15:
                n6 = '良好'
            if 15 > n[6] >= 10:
                n6 = '及格'
            if 10 > n[6]:
                n6 = '不及格'
            n[6] = '引体向上' + n6

            if n[7] >= 78:
                n7 = '优秀'
            if 68 <= n[7] < 78:
                n7 = '良好'
            if 55 <= n[7] < 68:
                n7 = '及格'
            if 55 > n[7]:
                n7 = '不及格'
            n[7] = 'VCWI' + n7

            x.append(n)
            print(n)
    return x


if __name__ == "__main__":
    x = preprocess(filenames=filenames)
    dataSet = x
    D = list(map(set, dataSet))
    L, suppData = apriori(dataSet)
    print('L:', L)
    print('suppData:', suppData)


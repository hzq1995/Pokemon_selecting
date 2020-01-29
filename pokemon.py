import xlrd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib import rcParams
rcParams['font.family'] = 'simhei'

TYPE_NUM = 18  # 共存在18种属性的宝可梦
PARTY_NUM = 6  # 一个队伍中共包含6个宝可梦


def next_indices(now, index_deal=PARTY_NUM-1):
    if now[index_deal] + 1 > TYPE_NUM - PARTY_NUM + index_deal:
        if index_deal == 0:
            return None
        else:
            return next_indices(now, index_deal-1)
    else:
        now[index_deal] += 1
        for i in range(index_deal+1, PARTY_NUM):
            now[i] = now[index_deal] + i - index_deal
        return now


if __name__ == '__main__':
    excel = xlrd.open_workbook('属性克制表.xlsx')
    sheet = excel.sheets()[0]

    type_table = list()  # 存放18种属性
    matrix = np.zeros((TYPE_NUM, TYPE_NUM))  # 存放克制矩阵
    result_no = list()
    result_score = list()
    result_party = list()
    result_weak = list()

    for i in range(2, 2+TYPE_NUM):
        type_table.append(sheet.cell_value(1, i))
    type_table = np.array(type_table)  # 转化为numpy数组
    for i in range(2, 2+TYPE_NUM):
        for j in range(2, 2+TYPE_NUM):
            matrix[i-2, j-2] = sheet.cell_value(i, j)

    selected_indices = np.array([0, 1, 2, 3, 4, 5])
    cnt = 0
    while True:
        cnt += 1

        score = np.sum(np.max(matrix[selected_indices], axis=0))  # 计算每个组合的得分
        weak = np.argwhere(np.max(matrix[selected_indices], axis=0) < 2.)[:, 0]

        result_no.append(cnt-1)
        result_score.append(score)
        result_weak.append(weak.copy())
        result_party.append(selected_indices.copy())

        selected_indices = next_indices(selected_indices)  # 计算下一个组和
        if selected_indices is None:
            print('组合的总数量是:', cnt)
            break

    result_score = np.array(result_score)
    result_party = np.squeeze(np.array(result_party))
    result_weak = np.squeeze(np.array(result_weak))
    print('分数的范围是:', np.sort(np.unique(result_score)))

    optimal_indices = np.argwhere(result_score == 35.)
    sub_optimal_indices = np.argwhere(result_score == 34.)
    for i, val in enumerate(optimal_indices):
        print('组合{}, 得分：{}，组成：{}, 无法克制：{}'.format(i+1, result_score[val],
                                              type_table[result_party[val]],
                                              type_table[result_weak[val][0]]))
    print('\n')
    for i, val in enumerate(sub_optimal_indices):
        print('组合{}, 得分：{}，组成：{}, 无法克制：{}'.format(i+1, result_score[val],
                                              type_table[result_party[val]],
                                              type_table[result_weak[val][0]]))

    print('\n')
    t = np.concatenate((np.ravel(result_party[optimal_indices]), np.ravel(result_party[sub_optimal_indices])))
    appear_cnt = np.bincount(t)
    rank = np.argsort(appear_cnt)[::-1]
    for i, val in enumerate(rank):
        print('属性：{}，出现次数：{}'.format(type_table[rank[i]], appear_cnt[rank[i]]))

    data = appear_cnt[rank]
    labels = type_table[rank]
    plt.axes(aspect='equal')
    explode = [0.1, 0.05, 0.02, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]

    plt.gca().spines['right'].set_color('none')
    plt.gca().spines['top'].set_color('none')
    plt.gca().spines['left'].set_color('none')
    plt.gca().spines['bottom'].set_color('none')

    plt.pie(x=data, labels=labels, explode=explode, autopct='%3.2f%%')
    plt.show()

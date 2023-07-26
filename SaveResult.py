import os
import numpy as np
import pandas as pd


def save_result(save_path, sheet_name, acc, AMean, AStd, OAMean, OAStd, AAMean, AAStd, kMean, kStd, train_time,
                test_time, CLASS_NUM):
    rows = []
    best_iter = 0
    result = np.zeros((len(acc) + CLASS_NUM + 5, 2))

    rows.append("Method")

    for i in range(len(acc)):
        result[i, 0] = acc[i]
        rows.append(str(i + 1))
        print('{}:{}'.format(i, acc[i]))
        if acc[i] > acc[best_iter]:
            best_iter = i

    for i in range(CLASS_NUM):
        result[i + len(acc), 0] = 100 * AMean[i]
        result[i + len(acc), 1] = 100 * AStd[i]
        rows.append("Class " + str(i + 1))
        print("Class " + str(i) + ": " + "{:.2f}".format(100 * AMean[i]) + " ± " + "{:.2f}".format(100 * AStd[i]))

    print("average OA: " + "{:.2f}".format(OAMean) + " ± " + "{:.2f}".format(OAStd))
    result[len(acc) + CLASS_NUM, 0] = OAMean
    result[len(acc) + CLASS_NUM, 1] = OAStd
    print("average AA: " + "{:.2f}".format(100 * AAMean) + " ± " + "{:.2f}".format(100 * AAStd))
    result[len(acc) + CLASS_NUM + 1, 0] = 100 * AAMean
    result[len(acc) + CLASS_NUM + 1, 1] = 100 * AAStd
    print("average kappa: " + "{:.4f}".format(100 * kMean) + " ± " + "{:.4f}".format(100 * kStd))
    result[len(acc) + CLASS_NUM + 2, 0] = 100 * kMean
    result[len(acc) + CLASS_NUM + 2, 1] = 100 * kStd
    print("train time per DataSet(s): " + "{:.5f}".format(train_time))
    print("test time per DataSet(s): " + "{:.5f}".format(test_time))
    result[len(acc) + CLASS_NUM + 3, 0] = train_time
    result[len(acc) + CLASS_NUM + 4, 0] = test_time

    # print('best acc all={}'.format(acc[best_iDataset]))
    # result[len(acc) + CLASS_NUM + 5, 0] = acc[best_iDataset]

    rows = rows + ["OA", "AA", "Kappa", "train_time", "test_time"]

    result = np.around(result, 2)

    lists = [sheet_name]
    for i in range(result.shape[0]):
        if len(acc) <= i < result.shape[0] - 2:
            lists.append(str(result[i, 0]) + " ± " + str(result[i, 1]))
        else:
            lists.append(str(result[i, 0]))

    data = pd.DataFrame(lists)
    data.index = rows

    if not os.path.exists(save_path):
        writer = pd.ExcelWriter(save_path)  # 写入Excel文件
        data.to_excel(writer, sheet_name, header=False)  # 保存文件
    else:
        writer = pd.ExcelWriter(save_path, mode='a', engine='openpyxl', if_sheet_exists='new')
        data.to_excel(writer, sheet_name, header=False)  # 保存文件

    writer._save()

    print("Save finished!")

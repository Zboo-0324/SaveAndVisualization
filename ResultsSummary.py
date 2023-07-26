# -*- coding: utf-8 -*-
import os

import numpy as np
import openpyxl
import pandas as pd


def get_worksheet_count(file_path):
    # 打开Excel文件
    workbook = openpyxl.load_workbook(file_path)

    # 获取工作表列表
    worksheets = workbook.sheetnames

    # 获取工作表个数
    worksheet_count = len(worksheets)

    return worksheets, worksheet_count


# 将各方法的实验结果汇总到Excel文件中
def get_result(file_path, save_path, sheet_name="5"):
    # 设置保存的工作表名称
    # sheet_name = str(per_class)

    # 获取Excel文件中的工作表名称列表和工作表个数
    methods, method_count = get_worksheet_count(file_path)

    # 读取数据
    values = pd.read_excel(file_path, sheet_name=methods[0]).values
    rows, data = values[:, 0].tolist(), values[:, 1]

    # 循环遍历工作表，将数据逐一合并到data数组中
    for i in range(1, method_count):
        data = np.column_stack((data, pd.read_excel(file_path, sheet_name=methods[i]).values[:, 1]))

    # 创建DataFrame，并设置列名和行索引
    df = pd.DataFrame(data, columns=methods, index=rows)

    # 判断保存路径是否存在，如果不存在则创建新的Excel文件，将DataFrame写入其中
    if not os.path.exists(save_path):
        with pd.ExcelWriter(save_path) as writer:
            df.to_excel(writer, sheet_name=sheet_name)
    else:
        # 如果文件已存在，则追加写入数据到新的工作表中
        with pd.ExcelWriter(save_path, mode='a', if_sheet_exists='new') as writer:
            df.to_excel(writer, sheet_name=sheet_name)

    print("已将各方法的实验结果汇总到Excel文件:", save_path)


# 将各方法的实验结果汇总到Excel文件中
def generalization_experiments_summary(file_path, save_path, dataset='IP'):
    # 设置保存的工作表名称，根据数据集名称命名工作表
    sheet_name = dataset

    # 获取Excel文件中的工作表名称列表和工作表个数
    samples_per_class, count = get_worksheet_count(file_path)

    # 读取第一个工作表的数据
    data = pd.read_excel(file_path, sheet_name=samples_per_class[0])

    # 获取列名（表头）列表，并去除第一个列（第一列通常是样本索引）
    headers = data.columns.tolist()[1:]

    # 获取第一个工作表的OA数据，从第2列（索引为1）开始取，并将其转换为1行n列的数组
    values = data.values[-5, :][1:].reshape(1, len(headers))

    # 循环遍历其他工作表，将每个工作表的OA数据逐一合并到values数组中
    if count > 1:
        for i in range(1, count):
            values = np.concatenate((values, pd.read_excel(
                file_path, sheet_name=samples_per_class[i]).values[-5, :][1:].reshape(1, len(headers))), axis=0)

    # 创建DataFrame，并设置列名和行索引，每行对应一个工作表的最后5行数据
    df = pd.DataFrame(values, columns=headers, index=samples_per_class)

    # 判断保存路径是否存在，如果不存在则创建新的Excel文件，将DataFrame写入其中
    if not os.path.exists(save_path):
        with pd.ExcelWriter(save_path) as writer:
            df.to_excel(writer, sheet_name=sheet_name)
    else:
        # 如果文件已存在，则追加写入数据到新的工作表中
        with pd.ExcelWriter(save_path, mode='a', if_sheet_exists='new') as writer:
            df.to_excel(writer, sheet_name=sheet_name)

    print("已将消融实验结果汇总到Excel文件:", save_path)


if __name__ == '__main__':
    # 获取文件夹下文件的个数
    folder_path = r'D:\Desktop\Result'
    for file in os.listdir(folder_path):
        get_result(os.path.join(folder_path, file), r'D:\Desktop\result.xlsx', file)

    generalization_experiments_summary(r'D:\Desktop\result.xlsx', r'D:\Desktop\xxx.xlsx', 'IP')

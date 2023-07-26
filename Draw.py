import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.ticker as mtick


def draw(filepath, savepath, draw=True):
    # 读取 Excel 表格数据
    accuracy = pd.read_excel(filepath, sheet_name='Sheet1', index_col=0)
    std = pd.read_excel(filepath, sheet_name='Sheet2', index_col=0)

    # 获取方法和数据集的列表
    methods = accuracy.columns.tolist()
    perclass = np.array(accuracy.index.tolist())

    # 自定义颜色列表
    colors = np.array([[255, 183, 3],  # 黄色
                       [240, 113, 103],  # 红色
                       [72, 202, 228],  # 蓝色
                       [157, 78, 221],  # 紫色
                       [56, 176, 0],  # 绿色
                       [21, 97, 109],  # 深蓝色
                       [255, 0, 0],  # 红色
                       [0, 0, 0]  # 黑色
                       ]) / 255.0

    markers = ['o', 's', '^', 'p', '*', 'd', 'v', '<', '>', 'h', 'x', '+']
    """
    o：圆   d：菱形   *：星形   ^：上三角形   v：下三角形   <：左三角形   >：右三角形
    s：正方形   p：五边形   H：六边形   x：叉号   +：加号   .：点标记   ,：像素标记
    """

    lines = []
    for i in range(len(methods)):  # 绘图线条的样式
        lines.append('dotted')  # solid, dashed, dashdot, dotted

    # 设置折线的宽度
    line_width = 4

    # 绘制折线图
    for i, method in enumerate(methods):
        plt.plot(perclass, np.array(accuracy.loc[:, method]), marker=markers[i % len(markers)], linewidth=line_width,
                 linestyle=lines[i], label=method, color=colors[i % len(colors)], markersize=15)
        plt.errorbar(perclass, np.array(accuracy.loc[:, method]), fmt=markers[i % len(markers)],
                     yerr=np.array(std.loc[:, method]), color=colors[i],
                     ecolor=colors[i], elinewidth=3, capsize=4)  # 绘制误差棒

    # 设置横坐标和纵坐标的范围、刻度和格式
    plt.xlim(xmin=np.min(perclass) - 1, xmax=np.max(perclass) + 1)
    plt.xticks(np.array(perclass))
    fmt = '%.0f'
    plt.gca().yaxis.set_major_formatter(mtick.FormatStrFormatter(fmt))

    plt.tick_params(labelsize=25)
    plt.xlabel(u'Traning samples per class', fontsize=25, labelpad=10)
    plt.ylabel(u'OA (%)', fontsize=25, labelpad=10)
    plt.legend(methods, loc="lower right", fontsize=22)

    plt.grid()
    fig = plt.gcf()
    fig.set_size_inches(12, 9)

    # 保存图形
    if draw:
        plt.savefig(savepath, bbox_inches='tight')
        print('Save to {}'.format(savepath))

    # 展示图形
    plt.show()


if __name__ == '__main__':
    datasets = 'IP'
    per_class = 5
    draw(filepath='D:/Desktop/result-IP.xlsx', savepath='D:/Desktop/{}-{}.png'.format(datasets, per_class), draw=False)

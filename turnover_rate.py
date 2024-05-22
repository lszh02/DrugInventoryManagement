# encoding=utf-8
import os

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def process_excel(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 获取药品名称、规格、单位、厂家
    drug_name = df['药品名称'].loc[0]
    drug_specifications = df['规格'].loc[0]
    unit = df['单位'].loc[0]
    manufacturer = df['厂家'].loc[0]

    # 选择需要的列
    selected_columns = ['类型', '入出库数量', '库存量', '操作日期']
    df = df[selected_columns].copy()

    # 将“操作日期”列转换为日期格式
    df['操作日期'] = pd.to_datetime(df['操作日期']).dt.date

    # 当有"住院摆药"记录时，就有“当日销量”
    if '住院摆药' in df['类型'].values:
        # 筛选有"住院摆药"的记录，按照操作日期分组，并计算每日的净销量（摆药-退药）
        daily_sales = df[df['类型'] == '住院摆药'].groupby('操作日期')['入出库数量'].sum().apply(lambda x: -x).reset_index()
        daily_sales.rename(columns={'入出库数量': '当日销量'}, inplace=True)

        # 按照操作日期分组，并计算每日的结余库存量
        daily_last_stock = df.groupby('操作日期')['库存量'].last().reset_index()
        daily_last_stock.rename(columns={'库存量': '日结库存'}, inplace=True)

        # 合并日结库存和每日销量的数据
        merged_df = pd.merge(daily_last_stock, daily_sales, on='操作日期', how='outer')

        # 生成一个包含所有日期的序列
        all_dates = pd.date_range(start='2023-01-01', end='2023-12-31')

        # 将 '操作日期' 列转换为日期时间格式,并作为键进行合并
        merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])
        merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')

        # 将日期时间格式的 '操作日期' 列转换为日期格式
        merged_df['操作日期'] = merged_df['操作日期'].dt.date

        # 使用前一个有效值填充每日最低库存的缺失值
        merged_df['日结库存'].fillna(method='ffill', inplace=True)

        # 使用0填充每日销量的缺失值
        merged_df['当日销量'].fillna(0, inplace=True)

        # 根据前7日销量计算预期10日计划量
        merged_df['预期10日计划量'] = merged_df['当日销量'].rolling(window=7, min_periods=1).mean().round(2) * 10

        if not merged_df.empty:
            # 画图
            draw_a_graph(merged_df, drug_name, unit)
            # 保存图像
            export_img(drug_name, drug_specifications)

    else:
        print("警告：选定周期内，无'住院摆药'记录！")


def draw_a_graph(df, drug_name, unit):
    # 设置matplotlib字体为通用字体
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 绘制当日最低库存的柱状图
    plt.figure(figsize=(20, 10))
    plt.bar(df['操作日期'], df['日结库存'], color='lightblue', label='日结库存')

    # 绘制折线图
    plt.plot(df['操作日期'], df['当日销量'], color='red', label='当日销量')
    plt.plot(df['操作日期'], df['预期10日计划量'], color='blue', label='预期10日计划量')

    # 设置图表标题和坐标轴标签
    plt.title(f'{drug_name}库存与销量分析（单位：{unit}）')
    plt.xlabel('日期')
    plt.ylabel('数量')

    # 设置x轴日期格式
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=14))
    plt.xticks(rotation=45)

    # 添加图例
    plt.legend()

    # 显示图表
    # plt.tight_layout()
    # plt.show()


def export_img(drug_name, drug_specifications):
    # 确保所有父文件夹都存在
    os.makedirs(export_path, exist_ok=True)
    # 导出图片命名
    export_img_file_name = f"{drug_name}_{drug_specifications}"  # 使用下划线连接药物名称和规格
    # 定义非法字符列表
    illegal_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    # 遍历非法字符列表，将每个非法字符替换为下划线
    for char in illegal_chars:
        export_img_file_name = export_img_file_name.replace(char, '_')

    # 使用下划线连接药物名称和规格并替换掉文件名中的非法字符
    export_img_file = os.path.join(export_path, f"{export_img_file_name}.png")  # 使用os.path.join来构造路径

    plt.savefig(export_img_file)


if __name__ == '__main__':
    for filename in os.listdir(directory_path):
        # 获取所有的Excel文件
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(directory_path, filename)
            process_excel(file_path)

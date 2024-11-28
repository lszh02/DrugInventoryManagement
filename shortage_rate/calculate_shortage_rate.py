import os
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger
from extract_data.extract_sales_data import extract_sales_data

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def calculate_shortage_rate(sales_df, start_date=None, end_date=None):  # start_date和end_date为空时，默认分析所有数据
    # 筛选出在start_date和end_date之间的数据
    start_date = max(datetime.strptime(start_date, '%Y-%m-%d').date() if start_date else sales_df['操作日期'].min(),
                     sales_df['操作日期'].min())
    end_date = min(datetime.strptime(end_date, '%Y-%m-%d').date() if end_date else sales_df['操作日期'].max(),
                   sales_df['操作日期'].max())
    filtered_df = sales_df[(sales_df['操作日期'] >= start_date) & (sales_df['操作日期'] <= end_date)]

    # 如果筛选出的数据为空，则返回None
    if filtered_df.empty:
        app_logger.warning(f"文件 {file_name} 中没有在 {start_date} 到 {end_date} 之间的数据")
        return None

    # df新增两列：是否短缺（日结库存<当日销量）、是否纳入统计（当日销量和日结库存均为0则为False）
    filtered_df['是否短缺'] = filtered_df['日结库存'] < filtered_df['当日销量']
    filtered_df['是否在售'] = (filtered_df['当日销量'] != 0) & (filtered_df['日结库存'] != 0)
    print(filtered_df)

    # 计算短缺率（短缺率 = 短缺天数 / 纳入统计天数）
    shortage_rate = filtered_df['是否短缺'].sum() / filtered_df['是否在售'].sum()
    print(filtered_df['是否短缺'].sum())
    print(filtered_df['是否在售'].sum())
    print(filtered_df['操作日期'].count())
    app_logger.info(f"短缺率: {shortage_rate}")

    # 找出短期日期前后一周的记录，并打印
    for index, row in filtered_df.iterrows():
        if row['是否短缺']:
            print(f"短缺日期: {row['操作日期']}")
            print(filtered_df[(filtered_df['操作日期'] >= row['操作日期'] - pd.Timedelta(days=7)) & (
                    filtered_df['操作日期'] <= row['操作日期'] + pd.Timedelta(days=7))])

    return shortage_rate


if __name__ == '__main__':
    # 测试代码
    file_name = r'D:\个人文件\张思龙\1.药事\5.降低静配中心药品供应短缺率\0消耗记录\202303_202402\4.xls'
    start_date = '2023-02-01'
    end_date = '2024-03-22'
    sales_data = extract_sales_data(file_name, start_date, end_date).get('销量数据')
    # print(sales_data)
    shortage_rate = calculate_shortage_rate(sales_data)
    print(shortage_rate)

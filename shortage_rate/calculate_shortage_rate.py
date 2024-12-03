import os
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger
from extract_data.extract_sales_data import extract_sales_data
from utils import filter_date_range

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def calculate_shortage_rate(sales_df, start_date=None, end_date=None):  # start_date和end_date为空时，默认分析所有数据
    """计算短缺率"""
    # 筛选日期范围
    filtered_df, start_date, end_date = filter_date_range(sales_df, start_date, end_date)

    # 检查筛选后的数据是否为空
    if filtered_df.empty:
        app_logger.warning(f"文件 {file_name} 中没有在 {start_date} 到 {end_date} 之间的数据")
        return None

    # df新增两列：是否短缺（日结库存<当日销量）、是否在售（当日销量不为0或者日结库存不为0）
    filtered_df['是否短缺'] = filtered_df['日结库存'] < filtered_df['当日销量']
    filtered_df['是否在售'] = (filtered_df['当日销量'] != 0) | (filtered_df['日结库存'] != 0)

    # 计算短缺率（短缺率 = 短缺天数 / 在售天数）
    shortage_rate = filtered_df['是否短缺'].sum() / filtered_df['是否在售'].sum()

    # 找出短期日期前后一周的记录，并打印
    # for index, row in filtered_df.iterrows():
    #     if row['是否短缺']:
    #         print(f"短缺日期: {row['操作日期']}")
    #         print(filtered_df[(filtered_df['操作日期'] >= row['操作日期'] - pd.Timedelta(days=7)) & (
    #                 filtered_df['操作日期'] <= row['操作日期'] + pd.Timedelta(days=7))])

    # 返回短缺天数、在售天数、统计天数、短缺率
    return {'短缺天数': filtered_df['是否短缺'].sum(),
            '在售天数': filtered_df['是否在售'].sum(),
            '短缺率': shortage_rate,
            '起始日期': start_date,
            '结束日期': end_date}


def process_file(file_path, start_date, end_date):
    # 处理单个文件
    try:
        data = extract_sales_data(file_path)
        if data is None:
            app_logger.warning(f"文件 {file_path} 提取数据失败")
            return None
    except Exception as e:
        error_logger.error(f"处理文件 {file_path} 时发生错误: {e}")
        return None

    app_logger.info(f"开始分析短缺率: {file_name}")

    drug_name = data.get('药品基本信息').get('药品名称')
    drug_spec = data.get('药品基本信息').get('规格')
    sales_data = data.get('销量数据')
    shortage_rate = calculate_shortage_rate(sales_data, start_date, end_date)

    if shortage_rate is None:
        return None

    return {
        '药品名称': drug_name,
        '规格': drug_spec,
        '短缺天数': shortage_rate.get('短缺天数'),
        '在售天数': shortage_rate.get('在售天数'),
        '短缺率': shortage_rate.get('短缺率'),
        '起始日期': shortage_rate.get('起始日期'),
        '结束日期': shortage_rate.get('结束日期'),
        '文件名': os.path.basename(file_path)
    }


if __name__ == '__main__':
    start_date = '2024-04-01'
    end_date = '2024-11-30'
    results = []

    # 获取所有文件，过滤出Excel文件
    all_files = os.listdir(directory_path)
    excel_files = [file for file in all_files if file.endswith(('.xlsx', '.xls'))]

    # 按文件名中的数字部分排序
    try:
        sorted_excel_files = sorted(excel_files, key=lambda x: int(x.split('.')[0]))
    except ValueError:
        sorted_excel_files = sorted(excel_files, key=lambda x: str(x.split('.')[0]))

    # 遍历所有Excel文件，计算短缺率
    for file_name in sorted_excel_files:
        file_path = os.path.join(directory_path, file_name)
        result = process_file(file_path, start_date, end_date)
        if result:
            results.append(result)

    # 将数据列表转换为DataFrame
    df = pd.DataFrame.from_records(results)

    # 导出结果到Excel文件
    os.makedirs(export_path, exist_ok=True)
    export_xls_file = os.path.join(export_path, "短缺率分析结果.xlsx")
    df.to_excel(export_xls_file, index=False)
    app_logger.info(f"短缺率分析结果已导出到 {export_xls_file}")


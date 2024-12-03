# encoding=utf-8
import os
from datetime import datetime

import pandas as pd

from config import app_logger, error_logger
from utils import filter_date_range, read_excel_file

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def extract_basic_info(df):
    return df[['自定义码', '药品名称', '规格', '单位', '入出库数量', '购入金额']].iloc[-1]


def filter_and_transform(df):
    selected_columns = ['类型', '入出库数量', '库存量', '操作日期']
    df = df[selected_columns].copy().dropna()
    df['操作日期'] = pd.to_datetime(df['操作日期']).dt.date
    return df


def calculate_daily_sales(df):
    daily_sales = df[df['类型'] == '住院摆药'].groupby('操作日期')['入出库数量'].sum().apply(lambda x: -x).reset_index()
    daily_sales.rename(columns={'入出库数量': '当日销量'}, inplace=True)
    return daily_sales


def calculate_daily_stock(df):
    daily_last_stock = df.groupby('操作日期')['库存量'].last().reset_index()
    daily_last_stock.rename(columns={'库存量': '日结库存'}, inplace=True)
    return daily_last_stock


def merge_and_fillna(daily_sales, daily_last_stock, start_date, end_date):
    merged_df = pd.merge(daily_sales, daily_last_stock, on='操作日期', how='outer')
    merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])
    all_dates = pd.date_range(start=start_date, end=end_date, freq='D')
    merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')
    merged_df['操作日期'] = merged_df['操作日期'].dt.date
    merged_df['日结库存'] = merged_df['日结库存'].ffill()
    merged_df['当日销量'] = merged_df['当日销量'].fillna(0)
    return merged_df


def extract_sales_data(file_path, start_date=None, end_date=None):
    """
    提取销量信息
    :param file_path: 文件路径
    :param start_date: 开始日期
    :param end_date: 结束日期
    :return: 销量信息
    """
    file_name = os.path.basename(file_path)
    app_logger.info(f"开始提取销量信息: {file_name}")

    df = read_excel_file(file_path)
    if df is None:
        return None

    basic_info = extract_basic_info(df)
    df = filter_and_transform(df)

    if '住院摆药' in df['类型'].values:
        daily_sales = calculate_daily_sales(df)
        daily_last_stock = calculate_daily_stock(df)
        _, start_date, end_date = filter_date_range(daily_sales, start_date, end_date)

        merged_df = merge_and_fillna(daily_sales, daily_last_stock, start_date, end_date)
        return {'文件名': file_name, '药品基本信息': basic_info, '销量数据': merged_df}
    else:
        app_logger.warning(
            f"警告：{basic_info['药品名称']}_{basic_info['规格']}在{start_date}到{end_date}期间没有住院摆药记录!文件名：{file_name}")
        return None


if __name__ == '__main__':
    # 测试代码
    file_name = r'D:\个人文件\张思龙\1.药事\5.降低静配中心药品供应短缺率\0消耗记录\202303_202402\155.xls'
    start_date = '2023-03-01'
    end_date = '2023-11-30'
    # start_date = None
    # end_date = None
    print(extract_sales_data(file_name, start_date, end_date))

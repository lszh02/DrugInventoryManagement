import os

import pandas as pd

from config import app_logger, error_logger


def extract_sales_info(file_path, start_date=None, end_date=None):
    """提取销量信息

    Args:
        file_path (str): Excel文件路径
        start_date (str, optional): 开始日期. Defaults to None.
        end_date (str, optional): 结束日期. Defaults to None.

    Returns:
        DataFrame: 包含销量信息的DataFrame
    """

    # 提取Excel文件名
    file_name = os.path.basename(file_path)
    app_logger.info(f"开始提取销量信息: {file_name}")

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 提取药品基本信息(基于最后一行数据)
    basic_info = df[['自定义码', '药品名称', '规格', '单位', '入出库数量', '购入金额']].iloc[-1]

    # 选择需要的列
    selected_columns = ['类型', '入出库数量', '库存量', '操作日期']
    df = df[selected_columns].copy()

    # 将“操作日期”列转换为日期格式
    df['操作日期'] = pd.to_datetime(df['操作日期']).dt.date

    # 当有"住院摆药"记录时，就有“当日销量”
    if '住院摆药' in df['类型'].values:
        # 筛选有"住院摆药"的记录，按照操作日期分组，并计算每日的净销量（摆药-退药）
        daily_sales = df[df['类型'] == '住院摆药'].groupby('操作日期')['入出库数量'].sum().apply(
            lambda x: -x).reset_index()
        daily_sales.rename(columns={'入出库数量': '当日销量'}, inplace=True)

        # 按照操作日期分组，并计算每日的结余库存量
        daily_last_stock = df.groupby('操作日期')['库存量'].last().reset_index()
        daily_last_stock.rename(columns={'库存量': '日结库存'}, inplace=True)

        # 合并日结库存和每日销量的数据
        merged_df = pd.merge(daily_sales, daily_last_stock, on='操作日期', how='outer')

        # 生成一个包含所有日期的序列(从开始日期到结束日期)
        if start_date is None:
            start_date = merged_df['操作日期'].min()
        if end_date is None:
            end_date = merged_df['操作日期'].max()

        all_dates = pd.date_range(start=start_date, end=end_date, freq='D')

        # 将 '操作日期' 列转换为日期时间格式,才能作为键进行合并
        merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])
        merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')

        # 将日期时间格式的 '操作日期' 列转换为日期格式
        merged_df['操作日期'] = merged_df['操作日期'].dt.date

        # 使用前一个有效值填充每日最低库存的缺失值
        merged_df['日结库存'] = merged_df['日结库存'].ffill()

        # 使用0填充每日销量的缺失值
        merged_df['当日销量'] = merged_df['当日销量'].fillna(0)

        return {'文件名': file_name,
                '药品基本信息': basic_info,
                '销量信息': merged_df
                }

    else:
        app_logger.info(f"{basic_info['药品名称']}_{basic_info['规格']}没有住院摆药记录!")
        return None

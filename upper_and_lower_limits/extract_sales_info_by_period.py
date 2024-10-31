import os

import pandas as pd

from config import app_logger, error_logger, directory_path

pd.set_option('display.unicode.ambiguous_as_wide', True)
pd.set_option('display.unicode.east_asian_width', True)

def extract_sales_info_by_period(file_path, start_date=None, end_date=None):
    """按区间提取销量信息

    Args:
        file_path (str): Excel文件路径
        start_date (str, optional): 开始日期. Defaults to None.
        end_date (str, optional): 结束日期. Defaults to None.

    Returns:
        DataFrame: 包含销量信息的DataFrame
    """

    # 提取Excel文件名
    file_name = os.path.basename(file_path)
    # app_logger.info(f"开始提取区间销量信息: {file_name}")

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 选择需要的列,舍弃无数据的行
    selected_columns = ['消耗日期', '自定义码', '药品名称', '规格', '数量', '单位', '销售金额']
    df = df[selected_columns].copy().dropna()

    # 将“消耗日期”列转换为日期格式
    df['消耗日期'] = pd.to_datetime(df['消耗日期']).dt.date
    # print(df['消耗日期'])

    # 查看最后一行数据
    print(df.tail(3))


if __name__ == '__main__':
    extract_sales_info_by_period(directory_path)

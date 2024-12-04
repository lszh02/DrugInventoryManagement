import os
from datetime import datetime

import pandas as pd

from config import app_logger


def read_excel_file(file_path, max_rows=65535):
    try:
        df = pd.read_excel(file_path)
        if len(df) >= max_rows:
            app_logger.warning(f"文件 {file_path} 的行数超过 {max_rows}，请选择续读文件")
            return continue_read_excel_file(file_path, df, max_rows)
        return df
    except Exception as e:
        app_logger.error(f"读取文件 {file_path} 时发生错误: {e}")
        return None


def continue_read_excel_file(file_path, df, max_rows):
    directory = os.path.dirname(file_path)
    base_name = os.path.basename(file_path)
    base_name_without_extension = os.path.splitext(base_name)[0]

    # 获取当前文件名的后缀
    suffix = base_name_without_extension.split('_')[-1]
    if not suffix.isdigit() or suffix == base_name_without_extension:
        suffix = '0'

    try:
        next_suffix = str(int(suffix) + 1).zfill(len(suffix))
        new_file_name = f"{base_name_without_extension}_{next_suffix}.xls"
        new_file_path = os.path.join(directory, new_file_name)

        if not os.path.exists(new_file_path):
            app_logger.info(f"文件 {new_file_path} 不存在，停止续读")
            return df

        new_df = pd.read_excel(new_file_path)
        df = pd.concat([df, new_df], ignore_index=True)

        if len(new_df) >= max_rows:
            continue_read_excel_file(new_file_path, df, max_rows)
        else:
            return df
    except Exception as e:
        app_logger.error(f"读取文件 {new_file_path} 时发生错误: {e}")
    return df


def parse_date(date_str):
    """解析日期字符串，如果为空则返回None"""
    return datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else None


def filter_date_range(df, start_date, end_date):
    """根据给定的日期范围筛选数据"""
    start_date = max(parse_date(start_date) or df['操作日期'].min(), df['操作日期'].min())
    end_date = min(parse_date(end_date) or df['操作日期'].max(), df['操作日期'].max())
    return df[(df['操作日期'] >= start_date) & (df['操作日期'] <= end_date)], start_date, end_date

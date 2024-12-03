from datetime import datetime

import pandas as pd

from config import app_logger


def read_excel_file(file_path):
    try:
        return pd.read_excel(file_path)
    except Exception as e:
        app_logger.error(f"读取文件 {file_path} 时发生错误: {e}")
        return None


def parse_date(date_str):
    """解析日期字符串，如果为空则返回None"""
    return datetime.strptime(date_str, '%Y-%m-%d').date() if date_str else None


def filter_date_range(df, start_date, end_date):
    """根据给定的日期范围筛选数据"""
    start_date = max(parse_date(start_date) or df['操作日期'].min(), df['操作日期'].min())
    end_date = min(parse_date(end_date) or df['操作日期'].max(), df['操作日期'].max())
    return df[(df['操作日期'] >= start_date) & (df['操作日期'] <= end_date)], start_date, end_date

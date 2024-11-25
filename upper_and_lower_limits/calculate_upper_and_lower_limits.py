import os
from datetime import datetime

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger
from extract_sales_info import extract_sales_info

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def analyze_sales_data(sales_info, start_date=None, end_date=None):  # start_date和end_date为空时，默认分析所有数据
    file_name = sales_info.get('文件名')
    basic_info = sales_info.get('药品基本信息')
    sales_df = sales_info.get('销量数据')

    app_logger.info(f"开始分析药品: {basic_info.get('药品名称')} 的销售数据，文件名: {file_name}")

    # 筛选出在start_date和end_date之间的数据
    start_date = max(datetime.strptime(start_date, '%Y-%m-%d').date() if start_date else sales_df['操作日期'].min(),
                     sales_df['操作日期'].min())
    end_date = min(datetime.strptime(end_date, '%Y-%m-%d').date() if end_date else sales_df['操作日期'].max(),
                   sales_df['操作日期'].max())
    filtered_df = sales_df[(sales_df['操作日期'] >= start_date) & (sales_df['操作日期'] <= end_date)]

    # 计算近5日日均销量、累计销量及其95百分位数
    filtered_df['近5日日均销量'] = filtered_df['当日销量'].rolling(window=5, min_periods=1).mean()
    filtered_df['近5日累计销量'] = filtered_df['当日销量'].rolling(window=5, min_periods=1).sum()
    percentile_95_5 = filtered_df['近5日累计销量'].quantile(0.95)

    # 计算近7日日均销量、累计销量及其95百分位数
    filtered_df['近7日日均销量'] = filtered_df['当日销量'].rolling(window=7, min_periods=1).mean()
    filtered_df['近7日累计销量'] = filtered_df['当日销量'].rolling(window=7, min_periods=1).sum()
    percentile_95_7 = filtered_df['近7日累计销量'].quantile(0.95)

    # 计算近10日日均销量、累计销量及其95百分位数
    filtered_df['近10日日均销量'] = filtered_df['当日销量'].rolling(window=10, min_periods=1).mean()
    filtered_df['近10日累计销量'] = filtered_df['当日销量'].rolling(window=10, min_periods=1).sum()
    percentile_95_10 = filtered_df['近10日累计销量'].quantile(0.95)

    # 如果筛选出的数据不为空，则进行后续分析
    if not filtered_df.empty:
        daily_avg_stock = filtered_df['日结库存'].mean()  # 日均库存
        daily_avg_sales = filtered_df['当日销量'].mean()  # 日均销量
        relative_std = filtered_df['当日销量'].std() / daily_avg_sales  # 相对标准差

        # 0销量天数占比
        zero_sales_days = filtered_df[filtered_df['当日销量'] == 0].shape[0]
        zero_sales_days_ratio = zero_sales_days / filtered_df.shape[0]

        # 设置库存上下限
        upper_limit, lower_limit, value_level = set_the_upper_and_lower_limits(basic_info,
                                                                               percentile_95_5=percentile_95_5,
                                                                               percentile_95_7=percentile_95_7,
                                                                               percentile_95_10=percentile_95_10,
                                                                               relative_std=relative_std,
                                                                               zero_sales_days_ratio=zero_sales_days_ratio)
        # 画图
        draw_a_graph(filtered_df, basic_info['药品名称'], basic_info['规格'], value_level=value_level,
                     upper_limit=round(upper_limit, 2), lower_limit=round(lower_limit, 2))

        # 导出图片
        export_img(file_name, basic_info['药品名称'], basic_info['规格'])

        return {'文件名': file_name,
                '自定义码': basic_info['自定义码'],
                '药品名称': basic_info['药品名称'],
                '规格': basic_info['规格'],
                '单位': basic_info['单位'],
                '拟设下限': round(lower_limit, 2),
                '拟设上限': round(upper_limit, 2),
                '10日销售额P95': round(percentile_95_10 * abs(basic_info['购入金额'] / basic_info['入出库数量']), 2),
                '销量价值等级': value_level,
                '销量波动': round(relative_std, 2),
                '库存天数': round(daily_avg_stock / daily_avg_sales, 2),
                '日均销量': round(daily_avg_sales, 2),
                '0销量天数占比': zero_sales_days_ratio,
                '起始日期': start_date,
                '结束日期': end_date
                }


def set_the_upper_and_lower_limits(basic_info, **kwargs):
    price = abs(basic_info['购入金额'] / basic_info['入出库数量'])
    percentile_95_5 = kwargs.get('percentile_95_5')
    percentile_95_7 = kwargs.get('percentile_95_7')
    percentile_95_10 = kwargs.get('percentile_95_10')
    relative_std = kwargs.get('relative_std')
    zero_sales_days_ratio = kwargs.get('zero_sales_days_ratio')

    # 基于10日销售额P95，分段设置上下限
    if 0 < percentile_95_10 * price < 1000:
        value_level = 1
        upper_limit = percentile_95_10 * 1.3
        lower_limit = percentile_95_7 * 1.3

    elif 1000 <= percentile_95_10 * price < 5000:
        value_level = 2
        upper_limit = percentile_95_10 * 1.2
        lower_limit = percentile_95_7 * 1.2

    elif 5000 <= percentile_95_10 * price < 10000:
        value_level = 3
        upper_limit = percentile_95_10 * 1.1
        lower_limit = percentile_95_7 * 1.1

    elif 10000 <= percentile_95_10 * price < 50000:
        value_level = 4
        upper_limit = percentile_95_10
        lower_limit = percentile_95_7

    elif percentile_95_10 * price > 50000:
        value_level = 5
        upper_limit = percentile_95_7
        lower_limit = percentile_95_5

    else:
        value_level = 0
        upper_limit = 0
        lower_limit = 0

    # 针对中等价值,通过波动性设置上下限
    # if relative_std > 3:
    #     upper_limit = percentile_95_10 * 1.3
    #     lower_limit = percentile_95_7 * 1.3
    # elif relative_std > 1:
    #     upper_limit = percentile_95_10 * 1.2
    #     lower_limit = percentile_95_7 * 1.2
    # else:
    return upper_limit, lower_limit, value_level


def draw_a_graph(df, drug_name, unit, **kwargs):
    value_level = kwargs.get('value_level')
    upper_limit = kwargs.get('upper_limit')
    lower_limit = kwargs.get('lower_limit')

    # 设置matplotlib字体为通用字体
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 绘制柱状图
    plt.figure(figsize=(20, 10))
    plt.bar(df['操作日期'], df['日结库存'], color='lightblue', label='日结库存')

    # 绘制折线图
    plt.plot(df['操作日期'], df['当日销量'], color='red', label='当日销量')
    if value_level == 5:
        plt.plot(df['操作日期'], df['近5日累计销量'], color='blue', label='近5日累计销量')
        plt.plot(df['操作日期'], df['近7日累计销量'], color='green', label='近7日累计销量')
    else:
        plt.plot(df['操作日期'], df['近7日累计销量'], color='blue', label='近7日累计销量')
        plt.plot(df['操作日期'], df['近10日累计销量'], color='green', label='近10日累计销量')

    # 添加水平线及其数值
    plt.axhline(y=lower_limit, color='blue', linestyle='--')
    plt.axhline(y=upper_limit, color='green', linestyle='--')
    plt.text(df['操作日期'].iloc[-1], lower_limit * 0.95, f'拟设下限：{lower_limit:.2f}', ha='left', va='center')
    plt.text(df['操作日期'].iloc[-1], upper_limit * 1.05, f'拟设上限：{upper_limit:.2f}', ha='left', va='center')

    # 显示文字
    if value_level == 1:
        text = '10日销售额：极低值'
    elif value_level == 2:
        text = '10日销售额：低值'
    elif value_level == 3:
        text = '10日销售额：中等值'
    elif value_level == 4:
        text = '10日销售额：高值'
    elif value_level == 5:
        text = '10日销售额：极高值'
    else:
        text = '10日销售额：未知'

    plt.text(df['操作日期'].iloc[-1], (upper_limit + lower_limit) / 2, text, ha='left', va='center')

    # 设置图表标题和坐标轴标签
    plt.title(f'{drug_name}库存与销量分析（单位：{unit}）')
    plt.xlabel('日期')
    plt.ylabel('数量')

    # 设置x轴日期格式
    plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
    plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=7))
    plt.xticks(rotation=45)

    # 添加图例
    plt.legend()

    # 显示图表
    plt.tight_layout()
    # plt.show()


def export_img(file_name, drug_name, drug_specifications):
    # 确保所有父文件夹都存在
    os.makedirs(export_path, exist_ok=True)
    # 导出图片命名
    export_img_file_name = f"{file_name}_{drug_name}_{drug_specifications}"  # 使用下划线连接药物名称和规格
    # 定义非法字符列表
    illegal_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    # 遍历非法字符列表，将每个非法字符替换为下划线
    for char in illegal_chars:
        export_img_file_name = export_img_file_name.replace(char, '_')

    # 使用下划线连接药物名称和规格并替换掉文件名中的非法字符
    export_img_file = os.path.join(export_path, f"{export_img_file_name}.png")  # 使用os.path.join来构造路径

    plt.savefig(export_img_file)


if __name__ == '__main__':
    # 获取所有文件
    all_files = os.listdir(directory_path)
    # 过滤出Excel文件
    excel_files = [file for file in all_files if file.endswith(('.xlsx', '.xls'))]
    # 按文件名中的数字部分排序
    try:
        sorted_excel_files = sorted(excel_files, key=lambda x: int(x.split('.')[0]))
    except ValueError:
        sorted_excel_files = sorted(excel_files, key=lambda x: str(x.split('.')[0]))

    # 遍历所有Excel文件，提取销量数据
    sales_data = []
    for filename in sorted_excel_files:
        file_path = os.path.join(directory_path, filename)
        try:
            result = extract_sales_info(file_path)
            if result:
                sales_data.append(result)
        except Exception as e:
            error_logger.error(f"处理文件 {filename} 时发生错误: {e}")
    app_logger.info(f"提取销量数据，完成！")

    # 分析销量数据并导出结果
    results = []
    for sale_info in sales_data:
        result = analyze_sales_data(sale_info, '2023-04-01', '2023-11-30')
        if result:
            results.append(result)
    app_logger.info(f"分析销量数据，完成！")

    # 将数据列表转换为DataFrame
    df = pd.DataFrame.from_records(results)
    # 导出结果到Excel文件
    os.makedirs(export_path, exist_ok=True)
    export_xls_file = os.path.join(export_path, "销量分析结果.xlsx")
    df.to_excel(export_xls_file, index=False)
    app_logger.info(f"导出结果到: {export_xls_file}")

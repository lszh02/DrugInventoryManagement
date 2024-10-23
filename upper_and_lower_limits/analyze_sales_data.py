import os
import time

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
    sales_df = sales_info.get('销量信息')

    app_logger.info(f"开始分析药品: {basic_info.get('药品名称')} 的销售数据，文件名: {file_name}")

    # 计算近5日日均销量、累计销量及其95百分位数
    sales_df['近5日日均销量'] = sales_df['当日销量'].rolling(window=5, min_periods=1).mean()
    sales_df['近5日累计销量'] = sales_df['当日销量'].rolling(window=5, min_periods=1).sum()
    percentile_95_5 = sales_df['近5日累计销量'].quantile(0.95)

    # 计算近10日日均销量、累计销量及其95百分位数
    sales_df['近10日日均销量'] = sales_df['当日销量'].rolling(window=10, min_periods=1).mean()
    sales_df['近10日累计销量'] = sales_df['当日销量'].rolling(window=10, min_periods=1).sum()
    percentile_95_10 = sales_df['近10日累计销量'].quantile(0.95)

    # 计算sales_df中当日销量的相对标准差
    daily_avg_stock = sales_df['日结库存'].mean()  # 日均库存
    daily_avg_sales = sales_df['当日销量'].mean()  # 日均销量
    relative_std = sales_df['当日销量'].std() / daily_avg_sales

    # 0销量天数占比
    zero_sales_days = sales_df[sales_df['当日销量'] == 0].shape[0]
    zero_sales_days_ratio = zero_sales_days / sales_df.shape[0]

    # 设置库存上下限
    upper_limit, lower_limit, low_value_level = set_the_upper_and_lower_limits(basic_info, percentile_95_5,
                                                                               percentile_95_10, relative_std)

    if not sales_df.empty:
        # 画图
        # draw_a_graph(sales_df, basic_info['药品名称'], basic_info['规格'],
        #              low_value_level=low_value_level, relative_std=round(relative_std, 2),
        #              upper_limit=round(upper_limit, 2), lower_limit=round(lower_limit, 2))
        #
        # # 导出图片
        # export_img(file_name, basic_info['药品名称'], basic_info['规格'])

        return {'文件名': file_name,
                '药品名称': basic_info['药品名称'],
                '规格': basic_info['规格'],
                '单位': basic_info['单位'],
                '厂家': basic_info['厂家'],
                '拟设下限': round(lower_limit, 2),
                '拟设上限': round(upper_limit, 2),
                '10日销售额P95': round(percentile_95_10 * basic_info['购入金额'], 2),
                '销量价值': '极低值' if low_value_level == 1 else '低值' if low_value_level == 2 else None,
                '销量波动': '高波动' if relative_std > 3 else '中波动' if relative_std > 1 else '低波动',
                '库存天数': round(daily_avg_stock / daily_avg_sales, 2),
                '日均销量': round(daily_avg_sales, 2),
                '0销量天数占比': zero_sales_days_ratio,
                '起始日期': sales_df['操作日期'].min(),
                '结束日期': sales_df['操作日期'].max(),
                }


def set_the_upper_and_lower_limits(basic_info, percentile_95_5, percentile_95_10, relative_std):
    # 针对低价值的药品，设置库存上下限与低价值级别：0（非低价值）、1（极低价值）、2（低价值）
    if abs(percentile_95_10 * basic_info['购入金额']) < 500:
        upper_limit = percentile_95_10 * 1.5
        lower_limit = percentile_95_5 * 1.5
        low_value_level = 1
    elif abs(percentile_95_10 * basic_info['购入金额']) < 1000:
        upper_limit = percentile_95_10 * 1.3
        lower_limit = percentile_95_5 * 1.3
        low_value_level = 2
    else:
        low_value_level = 0
        # 针对非低价值,通过波动性设置上下限
        if relative_std > 3:
            upper_limit = percentile_95_10 * 1.3
            lower_limit = percentile_95_5 * 1.3
        elif relative_std > 1:
            upper_limit = percentile_95_10 * 1.2
            lower_limit = percentile_95_5 * 1.2
        else:
            upper_limit = percentile_95_10 * 1.1
            lower_limit = percentile_95_5 * 1.1
    return upper_limit, lower_limit, low_value_level


def draw_a_graph(df, drug_name, unit, **kwargs):
    # 设置matplotlib字体为通用字体
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 绘制柱状图
    plt.figure(figsize=(20, 10))
    plt.bar(df['操作日期'], df['日结库存'], color='lightblue', label='日结库存')

    # 绘制折线图
    plt.plot(df['操作日期'], df['当日销量'], color='red', label='当日销量')
    plt.plot(df['操作日期'], df['近5日累计销量'], color='blue', label='近5日累计销量')
    plt.plot(df['操作日期'], df['近10日累计销量'], color='green', label='近10日累计销量')

    # 添加水平线
    plt.axhline(y=kwargs.get('lower_limit'), color='blue', linestyle='--', label='拟设下限')
    plt.axhline(y=kwargs.get('upper_limit'), color='green', linestyle='--', label='拟设上限')

    # 显示文字：销量价值（用“低值”、“极低值”表示）和波动情况（用“高波动”、“中波动”、“低波动”表示）
    if kwargs.get('low_value_level') == 1:
        plt.text(df['操作日期'].iloc[-1], df['日结库存'].iloc[-1], f'销量价值：极低值', ha='right', va='center')
    elif kwargs.get('low_value_level') == 2:
        plt.text(df['操作日期'].iloc[-1], df['日结库存'].iloc[-1], f'销量价值：低值', ha='right', va='center')
    else:
        if kwargs.get('relative_std') > 3:
            plt.text(df['操作日期'].iloc[-1], df['日结库存'].iloc[-1], f'波动情况：高波动', ha='right', va='center')
        elif kwargs.get('relative_std') > 1:
            plt.text(df['操作日期'].iloc[-1], df['日结库存'].iloc[-1], f'波动情况：中波动', ha='right', va='center')
        else:
            plt.text(df['操作日期'].iloc[-1], df['日结库存'].iloc[-1], f'波动情况：低波动', ha='right', va='center')

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
    app_logger.info(f"已提取销量数据")

    # 分析销量数据并导出结果
    results = []
    for sale_info in sales_data:
        result = analyze_sales_data(sale_info)
        if result:
            results.append(result)
    app_logger.info(f"已分析销量数据")

    # 将数据列表转换为DataFrame
    df = pd.DataFrame.from_records(results)
    # 导出结果到Excel文件
    os.makedirs(export_path, exist_ok=True)
    export_xls_file = os.path.join(export_path, "药品库存上下限3.xlsx")
    df.to_excel(export_xls_file, index=False)
    app_logger.info(f"导出文件: {export_xls_file}")

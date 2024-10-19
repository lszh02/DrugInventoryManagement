import os
import time

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def process_excel(file_path, start_date=None, end_date=None):
    """
    处理Excel文件，计算药品的上下限
    :param file_path: Excel文件路径
    :param start_date: 开始日期
    :param end_date: 结束日期
    :return:
    """

    # 提取Excel文件名
    file_name = os.path.basename(file_path)
    app_logger.info(f"开始处理文件: {file_name}")

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 提取药品基本信息(基于最后一行数据)
    basic_info = df[['药品名称', '规格', '单位', '购入金额', '厂家']].iloc[-1]
    # app_logger.info(f"提取药品基本信息:\n{basic_info}")

    # 选择需要的列
    selected_columns = ['类型', '入出库数量', '库存量', '操作日期']
    df = df[selected_columns].copy()
    # print(df.head())

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

        # 计算近5日日均销量、累计销量及其95百分位数
        merged_df['近5日日均销量'] = merged_df['当日销量'].rolling(window=5, min_periods=1).mean()
        merged_df['近5日累计销量'] = merged_df['当日销量'].rolling(window=5, min_periods=1).sum()
        percentile_95_5 = merged_df['近5日累计销量'].quantile(0.95)

        # 计算近10日日均销量、累计销量及其95百分位数
        merged_df['近10日日均销量'] = merged_df['当日销量'].rolling(window=10, min_periods=1).mean()
        merged_df['近10日累计销量'] = merged_df['当日销量'].rolling(window=10, min_periods=1).sum()
        percentile_95_10 = merged_df['近10日累计销量'].quantile(0.95)

        # 计算merged_df中当日销量的相对标准差
        daily_avg_stock = merged_df['日结库存'].mean()  # 日均库存
        daily_avg_sales = merged_df['当日销量'].mean()  # 日均销量
        relative_std = merged_df['当日销量'].std() / daily_avg_sales

        # 计算用药天数占比
        # merged_df['用药天数占比'] = merged_df['当日销量'].apply(lambda x: 1 if x > 0 else 0).rolling(window=5, min_periods=1).sum()

        # 设置库存上下限
        upper_limit, lower_limit, low_value_level = set_the_upper_and_lower_limits(basic_info, percentile_95_5,
                                                                                   percentile_95_10, relative_std)

        if not merged_df.empty:
            # 画图
            draw_a_graph(merged_df, basic_info['药品名称'], basic_info['规格'],
                         low_value_level=low_value_level, relative_std=round(relative_std, 2),
                         upper_limit=round(upper_limit, 2), lower_limit=round(lower_limit, 2))

            # 导出图片
            export_img(file_name, basic_info['药品名称'], basic_info['规格'])

            return {'文件名': file_name,
                    '药品名称': basic_info['药品名称'],
                    '规格': basic_info['规格'],
                    '单位': basic_info['单位'],
                    '厂家': basic_info['厂家'],
                    '拟设下限': round(lower_limit, 2),
                    '拟设上限': round(upper_limit, 2),
                    '销量价值': '极低值' if low_value_level == 1 else '低值' if low_value_level == 2 else None,
                    '销量波动': '高波动' if relative_std > 3 else '中波动' if relative_std > 1 else '低波动',
                    '库存天数': round(daily_avg_stock / daily_avg_sales, 2),
                    '日均销量': round(daily_avg_sales, 2),
                    '起始日期': merged_df['操作日期'].min(),
                    '结束日期': merged_df['操作日期'].max(),
                    # '用药天数占比': round(merged_df['用药天数占比'].sum(), 2),
                    }

    else:
        app_logger.info(f"{basic_info['药品名称']}_{basic_info['规格']}没有住院摆药记录!")


def set_the_upper_and_lower_limits(basic_info, percentile_95_5, percentile_95_10,
                                   relative_std):
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
    sorted_excel_files = sorted(excel_files, key=lambda x: str(x.split('.')[0]))

    # 遍历所有Excel文件
    results = []
    for filename in sorted_excel_files:
        file_path = os.path.join(directory_path, filename)
        try:
            result = process_excel(file_path)
            if result:
                results.append(result)
        except Exception as e:
            error_logger.error(f"处理文件 {filename} 时发生错误: {e}")

    # 将结果写入日志文件
    app_logger.info(f"信息汇总（含库存上下限）:\n{results}")
    # 将数据列表转换为DataFrame
    df = pd.DataFrame.from_records(results)
    # 将DataFrame写入Excel文件，不包括索引号
    export_xls_file = os.path.join(export_path, f"药品库存上下限.xlsx")
    # 确保所有父文件夹都存在
    os.makedirs(export_path, exist_ok=True)
    df.to_excel(export_xls_file, index=False)
    app_logger.info(f"导出文件: {export_xls_file}")

import os

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def process_excel(file_path, start_date=None, end_date=None):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 提取药品基本信息
    basic_info = df[['药品名称', '规格', '单位', '厂家']].iloc[0]
    app_logger.info(f"提取药品基本信息:\n{basic_info}")

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
        # print(daily_sales.head())

        # 按照操作日期分组，并计算每日的结余库存量
        daily_last_stock = df.groupby('操作日期')['库存量'].last().reset_index()
        daily_last_stock.rename(columns={'库存量': '日结库存'}, inplace=True)
        # print(daily_last_stock.head())

        # 合并日结库存和每日销量的数据
        merged_df = pd.merge(daily_sales, daily_last_stock, on='操作日期', how='outer')
        # print(merged_df.head())

        # 生成一个包含所有日期的序列(从开始日期到结束日期)
        if start_date is None:
            start_date = merged_df['操作日期'].min()
        if end_date is None:
            end_date = merged_df['操作日期'].max()

        all_dates = pd.date_range(start=start_date, end=end_date, freq='D')

        # 将 '操作日期' 列转换为日期时间格式,才能作为键进行合并
        merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])
        merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')
        # print(merged_df.head())

        # 将日期时间格式的 '操作日期' 列转换为日期格式
        merged_df['操作日期'] = merged_df['操作日期'].dt.date

        # 使用前一个有效值填充每日最低库存的缺失值
        merged_df['日结库存'] = merged_df['日结库存'].ffill()

        # 使用0填充每日销量的缺失值
        merged_df['当日销量'] = merged_df['当日销量'].fillna(0)
        # print(merged_df.head(10))

        # 计算近5日累计销量
        merged_df['近5日累计销量'] = merged_df['当日销量'].rolling(window=5, min_periods=1).sum()
        # print(merged_df.head())

        # 计算近7日累计销量
        merged_df['近10日累计销量'] = merged_df['当日销量'].rolling(window=10, min_periods=1).sum()
        # print(merged_df.head(15))

        # 计算merged_df中近5日累计销量的95百分位数
        percentile_95_5 = merged_df['近5日累计销量'].quantile(0.95)
        # print(f"近5日累计销量95百分位: {percentile_95_5}")

        # 计算merged_df中近10日累计销量的95百分位数
        percentile_95_10 = merged_df['近10日累计销量'].quantile(0.95)
        # print(f"近10日累计销量95百分位: {percentile_95_10}")

        # 计算merged_df中当日销量的相对标准差
        daily_avg_sales = merged_df['当日销量'].mean()
        relative_std = merged_df['当日销量'].std() / daily_avg_sales

        if not merged_df.empty:
            # 画图
            draw_a_graph(merged_df, basic_info['药品名称'], basic_info['规格'], percentile_95_5, percentile_95_10, )
            # 导出图片
            export_img(basic_info['药品名称'], basic_info['规格'])

            return {'药品名称': basic_info['药品名称'],
                    '规格': basic_info['规格'],
                    '单位': basic_info['单位'],
                    '厂家': basic_info['厂家'],
                    '日均销量': round(daily_avg_sales, 2),
                    '拟设下限': round(percentile_95_5, 2),
                    '拟设上限': round(percentile_95_10, 2),
                    '相对标准差': round(relative_std, 2)}

    else:
        print(f"{basic_info['药品名称']}_{basic_info['规格']}没有住院摆药记录!")


def draw_a_graph(df, drug_name, unit, percentile_95_5, percentile_95_10):
    # 设置matplotlib字体为通用字体
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False

    # 绘制当日最低库存的柱状图
    plt.figure(figsize=(20, 10))
    plt.bar(df['操作日期'], df['日结库存'], color='lightblue', label='日结库存')

    # 绘制折线图
    plt.plot(df['操作日期'], df['当日销量'], color='red', label='当日销量')
    plt.plot(df['操作日期'], df['近5日累计销量'], color='orange', label='近5日累计销量')
    plt.plot(df['操作日期'], df['近10日累计销量'], color='green', label='近10日累计销量')
    plt.axhline(y=percentile_95_5, color='blue', linestyle='--', label='拟设下限')
    plt.axhline(y=percentile_95_10, color='purple', linestyle='--', label='拟设上限')

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


def export_img(drug_name, drug_specifications):
    # 确保所有父文件夹都存在
    os.makedirs(export_path, exist_ok=True)
    # 导出图片命名
    export_img_file_name = f"{drug_name}_{drug_specifications}"  # 使用下划线连接药物名称和规格
    # 定义非法字符列表
    illegal_chars = ['\\', '/', ':', '*', '?', '"', '<', '>', '|']
    # 遍历非法字符列表，将每个非法字符替换为下划线
    for char in illegal_chars:
        export_img_file_name = export_img_file_name.replace(char, '_')

    # 使用下划线连接药物名称和规格并替换掉文件名中的非法字符
    export_img_file = os.path.join(export_path, f"{export_img_file_name}.png")  # 使用os.path.join来构造路径

    plt.savefig(export_img_file)


if __name__ == '__main__':
    results = []
    for filename in os.listdir(directory_path):
        # 获取所有的Excel文件
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(directory_path, filename)
            result = process_excel(file_path)
            # 如果结果不为空，则添加到结果列表中
            if result:
                results.append(result)
    app_logger.info(f"库存上下限:\n{results}")

    # 将数据列表转换为DataFrame
    df = pd.DataFrame.from_records(results)
    # 将DataFrame写入Excel文件，不包括索引号
    export_xls_file = os.path.join(export_path, f"药品库存上下限.xlsx")
    df.to_excel(export_xls_file, index=False)
    app_logger.info(f"导出文件: {export_xls_file}")

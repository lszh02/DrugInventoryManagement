import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

pd.set_option('expand_frame_repr', False)  # 当列太多时显示不清楚
pd.set_option('display.unicode.east_asian_width', True)  # 设置输出右对齐


def process_excel(file_path):
    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 获取药品名称、规格、单位
    drug_name = df['药品名称'].loc[0]
    drug_specifications = df['规格'].loc[0]
    unit = df['单位'].loc[0]

    # 选择需要的列
    selected_columns = ['类型', '入出库数量', '库存量', '操作日期']
    df = df[selected_columns]

    # 将“操作日期”列转换为日期格式
    df['操作日期'] = pd.to_datetime(df['操作日期']).dt.date

    # 按照操作日期分组，并计算每日的最低库存量
    daily_min_stock = df.groupby('操作日期')['库存量'].min().reset_index()
    daily_min_stock.rename(columns={'库存量': '当日最低库存'}, inplace=True)

    # 筛选出类型为"住院摆药"的记录
    hospitalized_medicine = df[df['类型'] == '住院摆药']
    # 按照操作日期分组，并计算每日的入出库数量总和
    daily_sales = hospitalized_medicine.groupby('操作日期')['入出库数量'].sum().abs().reset_index()
    daily_sales.rename(columns={'入出库数量': '当日销量'}, inplace=True)

    # 合并每日最低库存和每日销量的数据
    merged_df = pd.merge(daily_min_stock, daily_sales, on='操作日期', how='outer')

    # 将 '操作日期' 列转换为日期格式
    merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])

    # 生成一个包含所有日期的序列，也是日期时间格式
    all_dates = pd.date_range(start=merged_df['操作日期'].min(), end=merged_df['操作日期'].max())

    # 直接将 '操作日期' 列作为键进行合并，因为现在两者都是日期时间格式
    merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')

    # 使用前一个有效值填充每日最低库存的缺失值
    merged_df['当日最低库存'].fillna(method='ffill', inplace=True)

    # 使用0填充每日销量的缺失值
    merged_df['当日销量'].fillna(0, inplace=True)

    # 计算近7日、30日的日均销量
    merged_df['近1周的日均销量'] = merged_df['当日销量'].rolling(window=7, min_periods=1).mean()
    # merged_df['近1月的日均销量'] = merged_df['当日销量'].rolling(window=30, min_periods=1).mean()

    # 筛选出当日最低库存小于近1周的日均销量的相关信息
    filtered_df = merged_df[merged_df['当日最低库存'] < merged_df['近1周的日均销量']]

    # 显示筛选后的数据
    print(f'药品名称：{drug_name}，规格：{drug_specifications},单位：{unit}')
    print(filtered_df.head())

    # TODO 将短缺情况输出到Excel中

    if not filtered_df.empty:
        # 设置matplotlib字体为通用字体
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['axes.unicode_minus'] = False

        # 绘制当日最低库存的柱状图
        plt.figure(figsize=(10, 6))
        plt.bar(merged_df['操作日期'], merged_df['当日最低库存'], color='lightblue', label='当日最低库存')

        # 绘制折线图
        plt.plot(merged_df['操作日期'], merged_df['当日销量'], color='red', label='当日销量')
        plt.plot(merged_df['操作日期'], merged_df['近1周的日均销量'], color='orange', label='近1周的日均销量')
        # plt.plot(merged_df['操作日期'], merged_df['近1月的日均销量'], color='blue', label='近1月的日均销量')

        # 设置图表标题和坐标轴标签
        plt.title(f'{drug_name}库存与销量分析')
        plt.xlabel('日期')
        plt.ylabel('数量')

        # 设置x轴日期格式
        plt.gca().xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
        plt.gca().xaxis.set_major_locator(mdates.DayLocator(interval=10))
        plt.xticks(rotation=45)

        # 添加图例
        plt.legend()

        # 显示图表
        plt.tight_layout()
        plt.show()


if __name__ == '__main__':
    from drug_file_path import drug_file_path

    process_excel(drug_file_path)

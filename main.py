# encoding=utf-8
import os

import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

from config import directory_path, export_path, app_logger, error_logger

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
    df = df[selected_columns].copy()

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

    # 将 '操作日期' 列转换为日期时间格式
    merged_df['操作日期'] = pd.to_datetime(merged_df['操作日期'])

    # 生成一个包含所有日期的序列，也是日期时间格式
    all_dates = pd.date_range(start=merged_df['操作日期'].min(), end=merged_df['操作日期'].max())

    # 直接将 '操作日期' 列作为键进行合并，因为现在两者都是日期时间格式
    merged_df = pd.merge(all_dates.to_frame(name='操作日期'), merged_df, on='操作日期', how='left')

    # 将日期时间格式的 '操作日期' 列转换为日期格式
    merged_df['操作日期'] = merged_df['操作日期'].dt.date

    # 使用前一个有效值填充每日最低库存的缺失值
    merged_df['当日最低库存'].fillna(method='ffill', inplace=True)

    # 使用0填充每日销量的缺失值
    merged_df['当日销量'].fillna(0, inplace=True)

    # 计算近7日、30日的日均销量
    merged_df['近7日的日均销量'] = merged_df['当日销量'].rolling(window=7, min_periods=1).mean().round(2)
    # merged_df['近30日的日均销量'] = merged_df['当日销量'].rolling(window=30, min_periods=1).mean().round(2)

    # 筛选出短缺记录
    shortage_df = merged_df[merged_df['当日最低库存'] < merged_df['近7日的日均销量']].copy()

    if not shortage_df.empty:
        # 补充原始数据中的部分信息,以便输出到Excel中进行后续分析
        shortage_df['药品名称'] = drug_name
        shortage_df['规格'] = drug_specifications
        shortage_df['单位'] = unit
        shortage_df['源文件名'] = os.path.basename(file_path)

        # 将短缺记录以追加方式输出到Excel中
        export_records(shortage_df, export_path, 'shortage_records.xlsx')

        # 日志记录
        app_logger.info(f"\n{drug_name} {drug_specifications}短缺记录如下：\n{shortage_df}")

        # 设置matplotlib字体为通用字体
        plt.rcParams['font.sans-serif'] = ['SimHei']
        plt.rcParams['axes.unicode_minus'] = False

        # 绘制当日最低库存的柱状图
        plt.figure(figsize=(16, 7))
        plt.bar(merged_df['操作日期'], merged_df['当日最低库存'], color='lightblue', label='当日最低库存')

        # 绘制折线图
        plt.plot(merged_df['操作日期'], merged_df['当日销量'], color='red', label='当日销量')
        plt.plot(merged_df['操作日期'], merged_df['近7日的日均销量'], color='orange', label='近7日的日均销量')
        # plt.plot(merged_df['操作日期'], merged_df['近30日的日均销量'], color='blue', label='近30日的日均销量')

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

        # 使用plt.savefig()保存图像
        export_img(drug_name, drug_specifications)

        # # 显示图表
        # plt.tight_layout()
        # plt.show()
    else:
        # 无短缺记录时，随机输出5条记录，如果不够5条则取现有条数
        random_df = merged_df.sample(min(5, merged_df.shape[0])).sort_index()

        # 补充原始数据中的部分信息,以便输出到Excel中进行后续分析
        random_df['药品名称'] = drug_name
        random_df['规格'] = drug_specifications
        random_df['单位'] = unit
        random_df['源文件名'] = os.path.basename(file_path)

        # 将随机的非短缺记录以追加方式输出到Excel中
        export_records(random_df, export_path, 'random_non_shortage_records.xlsx')

        # 日志记录
        app_logger.info(f"\n{drug_name} {drug_specifications}无短缺记录，随机输出5条记录：\n{random_df}")


def export_records(df, export_path, export_file_name):
    # 确保所有父文件夹都存在
    os.makedirs(export_path, exist_ok=True)
    # 导出文件命名
    export_file = rf'{export_path}/{export_file_name}'

    try:
        # 检查文件是否存在
        if not os.path.exists(export_file):
            # 如果文件不存在，则创建一个新的Excel文件
            with pd.ExcelWriter(export_file, mode='w', engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        else:
            # 如果文件已存在，则进行追加
            with pd.ExcelWriter(export_file, mode='a', engine='openpyxl', if_sheet_exists='overlay') as writer:
                # 读取已有文件的行数
                last_row = pd.read_excel(export_file, header=None).shape[0]
                # 判断是否保留标题行
                if last_row > 0:
                    # 说明已经有数据了，不需要再保留标题行
                    df.to_excel(writer, index=False, header=False, startrow=last_row + 1)
                else:
                    # 说明没有数据，需要保留标题行
                    df.to_excel(writer, index=False, header=True, startrow=last_row + 1)
    except Exception as e:
        error_logger.error("导出记录失败: {}", e)


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
    for filename in os.listdir(directory_path):
        # 获取所有的Excel文件
        if filename.endswith('.xls') or filename.endswith('.xlsx'):
            file_path = os.path.join(directory_path, filename)
            process_excel(file_path)

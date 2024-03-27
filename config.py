import os
import loguru

directory_path = r'D:\药事\5.降低静配中心药品供应短缺率\消耗记录'
export_path = r'D:\药事\5.降低静配中心药品供应短缺率\汇总记录'

# 设置日志文件路径
app_log_path = os.path.join(export_path, "app.log")
error_log_path = os.path.join(export_path, "errors.log")

# 配置应用日志格式和输出
logger = loguru.logger
logger.add(app_log_path, format="{time} | {level} | {message}")

# 配置错误日志格式和输出
error_logger = loguru.logger
error_logger.add(error_log_path, format="{time} | {level} | {message}")

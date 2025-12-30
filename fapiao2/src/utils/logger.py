import os
from loguru import logger

def setup_logger():
    """配置日志，保存到与src同级的logs文件夹"""
    # 获取当前文件路径（src/utils/logger.py）
    current_path = os.path.abspath(__file__)
    # 计算项目根目录（src的父目录）
    root_dir = os.path.dirname(os.path.dirname(current_path))
    # 定义logs文件夹路径
    logs_dir = os.path.join(root_dir, "logs")
    
    # 确保logs文件夹存在，不存在则创建
    os.makedirs(logs_dir, exist_ok=True)
    
    # 日志文件路径
    log_path = os.path.join(logs_dir, "error.log")
    
    logger.add(
        log_path,
        level="ERROR",
        rotation="1 day",  # 每天轮转
        retention="30 days",  # 保留30天
        encoding="utf-8"
    )
    
    return logger
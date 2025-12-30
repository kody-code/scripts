import os
import time
from loguru import logger
from typing import Optional
from selenium.webdriver.remote.webdriver import WebDriver

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

def capture_screenshot(driver: WebDriver, contract_no: str, root_dir: Optional[str] = None) -> str:
    """
    捕获浏览器截图并以合同号命名保存
    
    Args:
        driver: Selenium浏览器驱动
        contract_no: 合同编号
        root_dir: 根目录，默认为项目根目录
        
    Returns:
        截图保存路径
    """
    try:
        if not root_dir:
            # 计算项目根目录
            current_path = os.path.abspath(__file__)
            root_dir = os.path.dirname(os.path.dirname(current_path))
            
        # 定义截图保存目录（使用传入的root_dir作为基础路径）
        screenshots_dir = root_dir
        os.makedirs(screenshots_dir, exist_ok=True)
        
        # 生成截图文件名（合同号+时间戳避免重复）
        timestamp = time.strftime("%Y%m%d_%H%M%S")
        filename = f"{contract_no}_{timestamp}.png"
        file_path = os.path.join(screenshots_dir, filename)
        
        # 保存截图
        driver.save_screenshot(file_path)
        logger.info(f"已保存错误截图: {file_path}")
        return file_path
    except Exception as e:
        logger.error(f"截图保存失败: {e}")
        return ""
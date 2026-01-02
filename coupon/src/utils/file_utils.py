import os
import re
import shutil
from pathlib import Path
from src.utils.date_utils import get_tomorrow_date
from src.config.constant import backup_path, home_path, template_path


def create_file(file_name: str):
    """创建文件"""
    temp_path = get_template_file_list(template_path)
    if file_name not in get_excel_file_list(home_path):
        if len(temp_path) > 0:
            shutil.copyfile(temp_path[0], home_path + "/" + file_name + get_tomorrow_date() + ".xlsx")
        else:
            print("没有模板文件")
    else:
        print("文件已存在")

def backup_file(file_path: str):
    """备份文件"""
    if not os.path.exists(file_path):
        return
    file_name = get_file_name_from_path(file_path)
    date = get_date_from_file_name(file_name)
    os.makedirs(Path(backup_path + "/" + date), exist_ok=True)
    if os.path.exists(file_path):
        shutil.move(file_path, backup_path + "/" + date + "/" + file_name)


def get_excel_file_list(file_path: str):
    """获取指定目录下的所有excel文件"""
    target_dir = Path(file_path)
    file_list = []
    for excel_file in target_dir.glob("*.xlsx"):
        file_list.append(str(excel_file))
    return file_list

def get_template_file_list(file_path: str):
    """获取指定目录下的所有模板文件"""
    target_dir = Path(file_path)
    file_list = []
    for template_file in target_dir.glob("*.xlsx"):
        file_list.append(str(template_file))
    return file_list

def get_file_name_from_path(file_path: str):
    """从文件路径中提取文件名"""
    return Path(file_path).name

def get_date_from_file_name(file_name: str):
    """从文件名提取 数字.数字 格式的日期/版本字符串"""
    pattern = r'(\d+\.\d+)'
    match_result = re.search(pattern, file_name)
    return match_result.group(1) if match_result else ""
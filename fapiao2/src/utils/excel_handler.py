import os
import pandas as pd

class ExcelHandler:
    @staticmethod
    def file_exists(file_path):
        """检查文件是否存在"""
        return os.path.exists(file_path)
    
    @staticmethod
    def read_excel(file_path):
        """读取Excel文件并返回字典列表"""
        df = pd.read_excel(file_path)
        return df.to_dict('records')
    
    @staticmethod
    def save_error_records(records, file_path):
        """保存错误记录到Excel"""
        df = pd.DataFrame(records)
        df.to_excel(file_path, index=False)
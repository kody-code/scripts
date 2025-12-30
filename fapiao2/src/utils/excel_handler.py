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
    def save_error_records(records, file_path, append=True):
        """保存错误记录到Excel，支持追加模式
        
        Args:
            records: 要保存的记录（可以是单条记录或列表）
            file_path: 保存路径
            append: 是否追加模式，默认为True
        """
        # 确保records是列表格式
        if not isinstance(records, list):
            records = [records]
            
        # 如果是追加模式且文件存在，先读取已有数据
        if append and os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            new_df = pd.DataFrame(records)
            # 合并数据并去重（根据合同编号）
            combined_df = pd.concat([existing_df, new_df], ignore_index=True)
            combined_df.drop_duplicates(subset=['合同编号'], keep='last', inplace=True)
            combined_df.to_excel(file_path, index=False)
        else:
            # 覆盖模式或文件不存在，直接写入
            df = pd.DataFrame(records)
            df.to_excel(file_path, index=False)
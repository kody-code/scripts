from src.utils.file_utils import create_file, backup_file, get_excel_file_list
from src.config.constant import home_path

file_names: list = ["京东优惠券"]

def backup_and_create_file():
    for file_path in get_excel_file_list(home_path):
        backup_file(file_path)
    for file_name in file_names:
        create_file(file_name)

def main():
    backup_and_create_file()

if __name__ == '__main__':
    main()
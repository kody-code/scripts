import tkinter as tk
from gui.main_window import InvoiceApp
from utils.dotenv_loader import load_env
from utils.logger import setup_logger

logger = setup_logger()

try:
    env_vars = load_env()
    logger.info("环境变量：{}".format(env_vars))
    logger.info("加载环境变量成功")
except FileNotFoundError as e:
    logger.error("未找到.env文件")

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
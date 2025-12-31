import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService


class BrowserDriver:
    def __init__(self, logger, driver_path=None):
        self.driver = None
        self.logger = logger
        self.driver_path = driver_path

    def init_driver(self):
        try:
            # 如果提供了自定义驱动路径，则使用该路径
            if self.driver_path:
                driver_path = self.driver_path
                self.logger.info(f"使用自定义Chrome驱动路径: {driver_path}")
            else:
                # 获取当前文件所在目录（src/core）
                current_dir = os.path.dirname(os.path.abspath(__file__))
                # 计算项目根目录（src 的父目录，因为 lib 与 src 同级）
                root_dir = os.path.dirname(
                    os.path.dirname(current_dir)
                )  # 新增：获取 src 的父目录

                if os.name == "nt":
                    # 修正路径：根目录/lib/win/chromedriver.exe
                    driver_path = os.path.join(
                        root_dir, "lib", "win", "chromedriver.exe"
                    )
                else:
                    # 修正路径：根目录/lib/chromedriver-linux64/chromedriver
                    driver_path = os.path.join(
                        root_dir, "lib", "chromedriver-linux64", "chromedriver"
                    )

                self.logger.info(f"根据系统类型选择Chrome驱动路径: {driver_path}")

            service = ChromeService(executable_path=driver_path)
            options = webdriver.ChromeOptions()
            # 可以添加一些选项，如无头模式
            # options.add_argument("--headless=new")
            self.driver = webdriver.Chrome(service=service, options=options)
            self.driver.maximize_window()
            self.logger.info("Chrome浏览器已启动并最大化窗口")
            return True
        except Exception as e:
            self.logger.error(f"初始化浏览器失败: {e}")
            return False

    def quit(self):
        if self.driver:
            try:
                self.driver.quit()
                self.logger.info("浏览器已关闭")
            except Exception as e:
                self.logger.error(f"关闭浏览器失败: {e}")
            self.driver = None

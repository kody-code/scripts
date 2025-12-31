import os
import time
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from src.core.browser_driver import BrowserDriver
from src.utils.excel_handler import ExcelHandler
from src.utils.logger import capture_screenshot


class InvoiceProcessor:
    def __init__(
        self,
        excel_path,
        username,
        password,
        email,
        error_file,
        logger,
        screenshot_dir,
        driver_path=None,
    ):
        self.excel_path = excel_path
        self.username = username
        self.password = password
        self.email = email
        self.error_file = error_file
        self.logger = logger
        self.screenshot_dir = screenshot_dir
        self.driver_path = driver_path
        self.browser = BrowserDriver(logger, driver_path)
        self.all_data = []
        self.total = 0

    def load_data(self):
        """加载Excel数据"""
        if not ExcelHandler.file_exists(self.excel_path):
            raise FileNotFoundError(f"Excel文件不存在: {self.excel_path}")

        self.all_data = ExcelHandler.read_excel(self.excel_path)
        self.logger.info(f"成功加载Excel数据: {self.all_data}")
        self.total = len(self.all_data)

        if self.total == 0:
            raise ValueError("Excel文件中没有数据")

        self.logger.info(f"共加载{self.total}条记录")
        return True

    def login(self) -> bool:
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.browser.driver.get(os.getenv("CRM_URL") or "")
            self.logger.info(f"已访问CRM系统首页: {self.browser.driver.current_url}")

            self.logger.info("开始执行登录操作...")
            # 输入用户名和密码
            username_input = WebDriverWait(self.browser.driver, 10).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, 'input[placeholder="账号"]')
                )
            )
            self.logger.info("找到用户名输入框")

            password_input = self.browser.driver.find_element(
                By.CSS_SELECTOR, 'input[placeholder="密码"]'
            )
            self.logger.info("找到密码输入框")

            username_input.send_keys(self.username)
            self.logger.info("已输入用户名")
            password_input.send_keys(self.password)
            self.logger.info("已输入密码")

            # 点击登录按钮
            login_button = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, "login-submit"))
            )
            self.logger.info("找到登录按钮")
            login_button.click()
            self.logger.info("已点击登录按钮，等待登录完成...")

            time.sleep(2)
            self.logger.success("登录成功")
            return True
        except Exception as e:
            self.logger.error(f"登录失败：{e}")
            return False

    def navigate_to_contract_page(self) -> bool:
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始导航至待开班合同表页面...")
            target_url = os.getenv("HETONG_URL") or ""
            self.browser.driver.get(target_url)
            self.logger.info(f"已直接访问待开班合同表页面：{target_url}")

            WebDriverWait(self.browser.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//div[contains(@class, "el-table")]')
                )
            )
            self.logger.info("待开班合同表页面加载完成")
            time.sleep(1)
            return True
        except Exception as e:
            self.logger.error(f"直接访问待开班合同表失败：{e}")
            return False

    def search_contract(self, target_contract_no: str) -> bool:
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info(f"开始搜索合同编号: {target_contract_no}")

            contract_no_form = WebDriverWait(self.browser.driver, 15).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        '//div[@class="el-form-item el-form-item--small"][label[@class="el-form-item__label" and text()="合同编号"]]',
                    )
                )
            )
            contract_no_input = contract_no_form.find_element(
                By.XPATH, './/input[@class="el-input__inner" and @placeholder="请输入"]'
            )
            contract_no_input.click()
            contract_no_input.clear()
            contract_no_input.send_keys(target_contract_no)
            self.logger.info(f"已输入合同编号：{target_contract_no}")

            search_btn = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//button[contains(@class, "submit-btn") and span[text()=" 搜索 "]]',
                    )
                )
            )
            self.browser.driver.execute_script("arguments[0].click();", search_btn)
            self.logger.info("已点击搜索按钮，等待搜索结果...")
            time.sleep(2)

            try:
                WebDriverWait(self.browser.driver, 10).until(
                    EC.presence_of_element_located(
                        (
                            By.XPATH,
                            '//div[@class="el-table__fixed-body-wrapper"]//tbody',
                        )
                    )
                )
                rows = self.browser.driver.find_elements(
                    By.XPATH, '//div[@class="el-table__fixed-body-wrapper"]//tbody/tr'
                )
                if len(rows) == 0 or (len(rows) == 1 and "暂无数据" in rows[0].text):
                    self.logger.warning(f"合同编号 {target_contract_no} 搜索结果为空")
                    return False
                self.logger.info(f"合同编号 {target_contract_no} 找到{len(rows)}条数据")
                return True
            except Exception as e:
                self.logger.error(f"验证搜索结果失败：{e}")
                return False
        except Exception as e:
            self.logger.error(f"搜索合同失败：{e}")
            return False

    def start_invoice_application(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始执行申请发票操作...")
            fixed_table_rows = WebDriverWait(self.browser.driver, 15).until(
                EC.presence_of_all_elements_located(
                    (
                        By.XPATH,
                        '//div[@class="el-table__fixed"]//div[@class="el-table__fixed-body-wrapper"]//tbody/tr',
                    )
                )
            )

            target_row_index = 0
            if target_row_index >= len(fixed_table_rows):
                raise Exception(
                    f"目标行索引 {target_row_index} 超出实际数据行数 {len(fixed_table_rows)}"
                )

            target_checkbox = fixed_table_rows[target_row_index].find_element(
                By.XPATH,
                './/label[@class="el-checkbox"]//span[@class="el-checkbox__input"]',
            )

            self.browser.driver.execute_script(
                "arguments[0].scrollIntoView(true);", target_checkbox
            )
            time.sleep(0.5)

            class_attr = target_checkbox.get_attribute("class")
            if class_attr is None or "is-checked" not in class_attr:
                try:
                    target_checkbox.click()
                except:
                    self.browser.driver.execute_script(
                        "arguments[0].click();", target_checkbox
                    )
                time.sleep(0.5)

            if "is-checked" in (target_checkbox.get_attribute("class") or ""):
                self.logger.info("✅ 固定列复选框已成功勾选")
            else:
                raise Exception("❌ 勾选失败，复选框仍未选中")

            apply_btn = WebDriverWait(self.browser.driver, 15).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//div[@class="table-tools-btnList"]//button[contains(@class, "table-tools-btn") and span[text()=" 申请发票 "]]',
                    )
                )
            )
            self.logger.info("开始点击「申请发票」按钮...")
            self.browser.driver.execute_script(
                "arguments[0].scrollIntoView(true);", apply_btn
            )
            time.sleep(0.5)
            self.browser.driver.execute_script("arguments[0].click();", apply_btn)
            self.logger.info("已点击「申请发票」按钮")
            time.sleep(1)

            try:
                message = WebDriverWait(self.browser.driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-message"))
                )
                msg_text = message.find_element(
                    By.CLASS_NAME, "el-message__content"
                ).text.strip()
                class_attr = message.get_attribute("class") or ""
                if (
                    "el-message--warning" in class_attr
                    or "el-message--error" in class_attr
                ):
                    self.logger.warning(f"提交被拦截：{msg_text}")
                    return False
                else:
                    self.logger.info(f"提示信息：{msg_text}")
            except TimeoutException:
                self.logger.info("未检测到任何提示，继续流程")
            return True
        except Exception as e:
            self.logger.error(f"点击「申请发票」按钮失败：{e}")
            return False

    def fill_invoice_form(self, content, amount):
        """填写发票表单"""
        try:
            # 选择发票类型
            if not self._select_fapiao_type():
                return False

            # 选择发票抬头
            if not self._select_invoice_type():
                return False

            # 选择抬头类型
            if not self._select_title_type():
                return False

            # 填写发票抬头
            if not self._insert_fapiao_title():
                return False

            # 填写发票内容
            if not self._insert_fapiao_content(str(content)):
                return False

            # 填写发票金额
            if not self._insert_fapiao_amount(str(amount)):
                return False

            # 填写接收邮箱
            if not self._insert_fapiao_email(self.email):
                return False

            return True
        except Exception as e:
            self.logger.error(f"填写发票表单失败：{e}")
            return False

    def submit_invoice(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始提交发票申请...")

            # 点击提交按钮
            ok_btn = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        '//div[@aria-label="发票申请"]//div[@class="el-dialog__footer"]//button[contains(@class,"el-button--primary")]',
                    )
                )
            )
            self.browser.driver.execute_script("arguments[0].click();", ok_btn)
            self.logger.info("已点击提交按钮")
            time.sleep(2)
            return True
        except Exception as e:
            self.logger.error(f"提交发票申请失败：{e}")
            return False

    def _select_fapiao_type(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始选择发票类型...")
            WebDriverWait(self.browser.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "el-dialog__title"))
            )
            invoice_group_select = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "label[for='invoiceGroup'] + div .el-select")
                )
            )
            invoice_group_select.click()
            self.logger.info("已打开发票类型下拉框")

            WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='增值税普通发票']",
                    )
                )
            ).click()
            self.logger.info("已选择发票类型：增值税普通发票")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"选择发票类型失败：{e}")
            return False

    def _select_invoice_type(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始选择发票抬头...")
            invoice_type_select = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "label[for='invoiceType'] + div .el-select")
                )
            )
            invoice_type_select.click()
            self.logger.info("已打开发票抬头下拉框")

            WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='电子票']",
                    )
                )
            ).click()
            self.logger.info("已选择发票抬头：电子票")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"选择发票抬头失败：{e}")
            return False

    def _select_title_type(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始选择抬头类型...")
            title_type_select = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "label[for='invoiceUpHeadType'] + div .el-select")
                )
            )
            title_type_select.click()
            self.logger.info("已打开抬头类型下拉框")

            WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='个人']",
                    )
                )
            ).click()
            self.logger.info("已选择抬头类型：个人")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"选择抬头类型失败：{e}")
            return False

    def _insert_fapiao_title(self):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info("开始填写发票抬头...")
            invoice_title_input = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "label[for='invoiceUpHead'] + div .el-input__inner",
                    )
                )
            )
            invoice_title_input.clear()
            invoice_title_input.send_keys("个人")
            self.logger.info("已填写发票抬头：个人")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"填写发票抬头失败：{e}")
            return False

    def _insert_fapiao_content(self, content: str):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info(f"开始填写发票内容：{content}")
            invoice_content_select = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "label[for='invoiceContext'] + div .el-select")
                )
            )
            invoice_content_select.click()
            self.logger.info("已打开发票内容下拉框")

            WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        f"//li[contains(@class, 'el-select-dropdown__item')]/span[text()='{content}']",
                    )
                )
            ).click()
            self.logger.info(f"已选择发票内容：{content}")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"填写发票内容失败：{e}")
            return False

    def _insert_fapiao_amount(self, amount: str):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info(f"开始填写发票金额：{amount}")
            invoice_amount_input = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.CSS_SELECTOR, "label[for='billMoney'] + div .el-input__inner")
                )
            )
            invoice_amount_input.clear()
            invoice_amount_input.send_keys(amount)
            self.logger.info(f"已填写发票金额：{amount}")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"填写发票金额失败：{e}")
            return False

    def _insert_fapiao_email(self, email: str):
        try:
            if not self.browser.driver:
                self.logger.error("浏览器驱动未初始化")
                return False

            self.logger.info(f"开始填写接收邮箱：{email}")
            email_input = WebDriverWait(self.browser.driver, 10).until(
                EC.element_to_be_clickable(
                    (
                        By.CSS_SELECTOR,
                        "label[for='invoiceEmail'] + div .el-input__inner",
                    )
                )
            )
            email_input.clear()
            email_input.send_keys(email)
            self.logger.info(f"已填写接收邮箱：{email}")
            time.sleep(0.5)
            return True
        except Exception as e:
            self.logger.error(f"填写接收邮箱失败：{e}")
            return False

    def process(self, progress_callback, stop_check, error_callback):
        """处理发票申请的主流程"""
        try:
            # 加载数据
            self.load_data()

            if not self.load_data():
                error_callback("加载数据失败")
                return

            # 初始化浏览器
            if not self.browser.init_driver():
                error_callback("浏览器初始化失败")
                return

            # 登录系统
            if not self.login():
                error_callback("登录失败")
                if self.browser.driver:
                    capture_screenshot(
                        self.browser.driver, "login_failed", self.screenshot_dir
                    )
                return

            # 导航到合同页面
            if not self.navigate_to_contract_page():
                error_callback("导航到合同页面失败")
                if self.browser.driver:
                    capture_screenshot(
                        self.browser.driver, "navigate_failed", self.screenshot_dir
                    )
                return

            # 处理每个合同
            for i, record in enumerate(self.all_data):
                if not stop_check():
                    break

                # 更新进度
                progress = (i + 1) / self.total * 100
                progress_callback(progress, f"处理中: {i+1}/{self.total}")

                # 获取合同信息
                contract_no = record.get("合同编号")
                invoice_content = record.get("开票项目")
                amount = record.get("开票金额")

                if not contract_no:
                    self.logger.warning("跳过缺少合同编号的记录")
                    ExcelHandler.save_error_records(
                        {**record, "错误原因": "缺少合同编号"}, self.error_file
                    )
                    continue

                self.logger.info(
                    f"\n===== 开始处理第{i+1}条记录: 合同编号 {contract_no} ====="
                )

                # 搜索合同
                if not self.search_contract(contract_no):
                    self.logger.warning(f"合同 {contract_no} 未找到，添加到错误记录")
                    if self.browser.driver:
                        screenshot_path = capture_screenshot(
                            self.browser.driver, contract_no, self.screenshot_dir
                        )
                    ExcelHandler.save_error_records(
                        {
                            **record,
                            "错误原因": "合同未找到",
                            "截图路径": screenshot_path,
                        },
                        self.error_file,
                    )
                    continue

                # 申请发票
                if not self.start_invoice_application():
                    self.logger.warning(f"合同 {contract_no} 申请发票失败")
                    if self.browser.driver:
                        screenshot_path = capture_screenshot(
                            self.browser.driver, contract_no, self.screenshot_dir
                        )
                    ExcelHandler.save_error_records(
                        {
                            **record,
                            "错误原因": "申请发票失败",
                            "截图路径": screenshot_path,
                        },
                        self.error_file,
                    )
                    continue

                # 填写发票表单
                if not self.fill_invoice_form(invoice_content, amount):
                    self.logger.warning(f"合同 {contract_no} 填写发票表单失败")
                    if self.browser.driver:
                        screenshot_path = capture_screenshot(
                            self.browser.driver, contract_no, self.screenshot_dir
                        )
                    ExcelHandler.save_error_records(
                        {
                            **record,
                            "错误原因": "填写发票表单失败",
                            "截图路径": screenshot_path,
                        },
                        self.error_file,
                    )
                    continue

                # 提交申请
                if not self.submit_invoice():
                    self.logger.warning(f"合同 {contract_no} 提交申请失败")
                    if self.browser.driver:
                        screenshot_path = capture_screenshot(
                            self.browser.driver, contract_no, self.screenshot_dir
                        )
                    ExcelHandler.save_error_records(
                        {
                            **record,
                            "错误原因": "提交申请失败",
                            "截图路径": screenshot_path,
                        },
                        self.error_file,
                    )
                    continue

                self.logger.success(f"合同 {contract_no} 处理成功")
                time.sleep(2)  # 处理间隔

            # 保存错误记录
            if os.path.exists(self.error_file) and os.path.getsize(self.error_file) > 0:
                self.logger.warning(f"有错误记录已保存至: {self.error_file}")

        finally:
            self.browser.quit()

    def stop(self):
        """停止处理并清理资源"""
        self.browser.quit()

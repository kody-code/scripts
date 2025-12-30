import os
import time
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from loguru import logger as log
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
import threading  # 新增线程支持

# 日志配置
log.add(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.log"),
    level="ERROR",
    rotation="1 day",
    retention="30 days",
    encoding="utf-8"
)

class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("发票申请系统")
        self.root.geometry("700x500")
        self.root.resizable(True, True)
        
        # 初始化变量
        self.excel_path = tk.StringVar()
        self.error_file = os.path.join(os.path.expanduser("~"), "Desktop", "error_records.xlsx")
        self.all_data = []
        self.driver = None
        self.is_running = False
        
        # 创建界面
        self.create_widgets()
        
    def create_widgets(self):
        # 创建标签页
        tab_control = ttk.Notebook(self.root)
        
        # 主界面标签页
        main_tab = ttk.Frame(tab_control)
        tab_control.add(main_tab, text="主界面")
        
        # 日志标签页
        log_tab = ttk.Frame(tab_control)
        tab_control.add(log_tab, text="运行日志")
        
        tab_control.pack(expand=1, fill="both")
        
        # 主界面内容
        frame = ttk.LabelFrame(main_tab, text="配置")
        frame.pack(padx=10, pady=10, fill="x")
        
        # Excel文件选择
        ttk.Label(frame, text="Excel文件路径:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ttk.Entry(frame, textvariable=self.excel_path, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="浏览", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)
        
        # 用户信息
        ttk.Label(frame, text="用户名:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.username_var = tk.StringVar(value=os.getenv("USER_NAME") or "")
        ttk.Entry(frame, textvariable=self.username_var).grid(row=1, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(frame, text="密码:").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.password_var = tk.StringVar(value=os.getenv("PASSWORD") or "")
        ttk.Entry(frame, textvariable=self.password_var, show="*").grid(row=2, column=1, padx=5, pady=5, sticky="w")
        
        ttk.Label(frame, text="接收邮箱:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.email_var = tk.StringVar(value="fapiao@cuour.org")
        ttk.Entry(frame, textvariable=self.email_var).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        
        # 按钮区域
        btn_frame = ttk.Frame(main_tab)
        btn_frame.pack(padx=10, pady=10)
        
        self.start_btn = ttk.Button(btn_frame, text="开始处理", command=self.start_processing)
        self.start_btn.pack(side="left", padx=5)
        
        self.stop_btn = ttk.Button(btn_frame, text="停止处理", command=self.stop_processing, state="disabled")
        self.stop_btn.pack(side="left", padx=5)
        
        # 进度条
        self.progress_var = tk.DoubleVar()
        progress_frame = ttk.LabelFrame(main_tab, text="进度")
        progress_frame.pack(padx=10, pady=10, fill="x")
        ttk.Progressbar(progress_frame, variable=self.progress_var, length=100).pack(padx=5, pady=5, fill="x")
        self.progress_label = ttk.Label(progress_frame, text="等待开始...")
        self.progress_label.pack(padx=5, pady=5, anchor="w")
        
        # 日志区域
        log_frame = ttk.LabelFrame(log_tab, text="运行日志")
        log_frame.pack(padx=10, pady=10, fill="both", expand=True)
        
        self.log_text = tk.Text(log_frame, wrap="word", state="disabled")
        self.log_text.pack(padx=5, pady=5, fill="both", expand=True, side="left")
        
        scrollbar = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side="right", fill="y")
        self.log_text.config(yscrollcommand=scrollbar.set)
        
        # 重定向日志到文本框
        import logging
        class TextHandler(logging.Handler):
            def __init__(self, text_widget):
                logging.Handler.__init__(self)
                self.text_widget = text_widget
                
            def emit(self, record):
                msg = self.format(record) + "\n"
                self.text_widget.configure(state="normal")
                self.text_widget.insert("end", msg)
                self.text_widget.see("end")
                self.text_widget.configure(state="disabled")
        
        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        log.add(text_handler)
    
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")],
            title="选择发票数据Excel文件"
        )
        if filename:
            self.excel_path.set(filename)
    
    def start_processing(self):
        if not self.excel_path.get():
            messagebox.showerror("错误", "请选择Excel文件")
            return
            
        if not self.username_var.get() or not self.password_var.get():
            messagebox.showerror("错误", "请输入用户名和密码")
            return
            
        if not self.email_var.get():
            messagebox.showerror("错误", "请输入接收邮箱")
            return
            
        self.is_running = True
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        
        # 使用线程执行耗时操作，避免界面无响应
        threading.Thread(target=self.process_invoices, daemon=True).start()
    
    def stop_processing(self):
        self.is_running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.progress_label.config(text="已停止")
        
        # 关闭浏览器
        if self.driver:
            try:
                self.driver.quit()
                log.info("浏览器已关闭")
            except Exception as e:
                log.error(f"关闭浏览器失败: {e}")
        self.driver = None
    
    def init_driver(self):
        try:
            # 获取当前文件所在目录
            current_dir = os.path.dirname(os.path.abspath(__file__))
            if os.name == 'nt':
                driver_path = os.path.join(current_dir, "lib", "win", "chromedriver.exe")
            else:
                driver_path = os.path.join(current_dir, "lib","chromedriver-linux64", "chromedriver")
            
            log.info(f"根据系统类型选择Chrome驱动路径: {driver_path}")
            service = ChromeService(executable_path=driver_path)
            options = webdriver.ChromeOptions()
            # 可以添加一些选项，如无头模式
            # options.add_argument("--headless=new")
            driver = webdriver.Chrome(service=service, options=options)
            driver.maximize_window()
            log.info("Chrome浏览器已启动并最大化窗口")
            
            driver.get(os.getenv("CRM_URL") or "")
            log.info(f"已访问CRM系统首页: {driver.current_url}")
            return driver
        except Exception as e:
            log.error(f"初始化浏览器失败: {e}")
            return None
    
    def login(self, user_name: str, password: str) -> bool:
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始执行登录操作...")
            # 输入用户名和密码
            username_input = WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="账号"]'))
            )
            log.info("找到用户名输入框")
            password_input = self.driver.find_element(By.CSS_SELECTOR, 'input[placeholder="密码"]')
            log.info("找到密码输入框")

            username_input.send_keys(user_name)
            log.info("已输入用户名")
            password_input.send_keys(password)
            log.info("已输入密码")

            # 点击登录按钮
            login_button = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CLASS_NAME, 'login-submit'))
            )
            log.info("找到登录按钮")
            login_button.click()
            log.info("已点击登录按钮，等待登录完成...")

            time.sleep(2)
            log.success("登录成功")
            return True
        except Exception as e:
            log.error(f"登录失败：{e}")
            return False
    
    def choose_hetong(self) -> bool:
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始导航至待开班合同表页面...")
            target_url = os.getenv("HETONG_URL") or ""
            self.driver.get(target_url)
            log.info(f"已直接访问待开班合同表页面：{target_url}")
            
            WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "el-table")]'))
            )
            log.info("待开班合同表页面加载完成")
            time.sleep(1)
            return True
        except Exception as e:
            log.error(f"直接访问待开班合同表失败：{e}")
            return False
    
    def search_hetong(self, target_contract_no: str) -> bool:
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            
            log.info(f"开始搜索合同编号: {target_contract_no}")

            contract_no_form = WebDriverWait(self.driver, 15).until(
                EC.presence_of_element_located(
                    (By.XPATH, '//div[@class="el-form-item el-form-item--small"][label[@class="el-form-item__label" and text()="合同编号"]]')
                )
            )
            contract_no_input = contract_no_form.find_element(
                By.XPATH, './/input[@class="el-input__inner" and @placeholder="请输入"]'
            )
            contract_no_input.click()
            contract_no_input.clear()
            contract_no_input.send_keys(target_contract_no)
            log.info(f"已输入合同编号：{target_contract_no}")

            search_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable(
                    (By.XPATH, '//button[contains(@class, "submit-btn") and span[text()=" 搜索 "]]')
                )
            )
            self.driver.execute_script("arguments[0].click();", search_btn)
            log.info("已点击搜索按钮，等待搜索结果...")
            time.sleep(2)

            try:
                WebDriverWait(self.driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@class="el-table__fixed-body-wrapper"]//tbody'))
                )
                rows = self.driver.find_elements(By.XPATH, '//div[@class="el-table__fixed-body-wrapper"]//tbody/tr')
                if len(rows) == 0 or (len(rows) == 1 and "暂无数据" in rows[0].text):
                    log.warning(f"合同编号 {target_contract_no} 搜索结果为空")
                    return False
                log.info(f"合同编号 {target_contract_no} 找到{len(rows)}条数据")
                return True
            except Exception as e:
                log.error(f"验证搜索结果失败：{e}")
                return False
        except Exception as e:
            log.error(f"搜索合同失败：{e}")
            return False
    
    def start_fapiao(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始执行申请发票操作...")
            fixed_table_rows = WebDriverWait(self.driver, 15).until(
                EC.presence_of_all_elements_located(
                    (By.XPATH, '//div[@class="el-table__fixed"]//div[@class="el-table__fixed-body-wrapper"]//tbody/tr')
                )
            )

            target_row_index = 0
            if target_row_index >= len(fixed_table_rows):
                raise Exception(f"目标行索引 {target_row_index} 超出实际数据行数 {len(fixed_table_rows)}")

            target_checkbox = fixed_table_rows[target_row_index].find_element(
                By.XPATH, './/label[@class="el-checkbox"]//span[@class="el-checkbox__input"]'
            )

            self.driver.execute_script("arguments[0].scrollIntoView(true);", target_checkbox)
            time.sleep(0.5)

            class_attr = target_checkbox.get_attribute("class")
            if class_attr is None or "is-checked" not in class_attr:
                try:
                    target_checkbox.click()
                except:
                    self.driver.execute_script("arguments[0].click();", target_checkbox)
                time.sleep(0.5)

            if "is-checked" in (target_checkbox.get_attribute("class") or ""):
                log.info("✅ 固定列复选框已成功勾选")
            else:
                raise Exception("❌ 勾选失败，复选框仍未选中")
                
            apply_btn = WebDriverWait(self.driver, 15).until(
                EC.element_to_be_clickable(
                    (By.XPATH,
                     '//div[@class="table-tools-btnList"]//button[contains(@class, "table-tools-btn") and span[text()=" 申请发票 "]]')
                )
            )
            log.info("开始点击「申请发票」按钮...")
            self.driver.execute_script("arguments[0].scrollIntoView(true);", apply_btn)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", apply_btn)
            log.info("已点击「申请发票」按钮")
            time.sleep(1)
            
            try:
                message = WebDriverWait(self.driver, 3).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-message"))
                )
                msg_text = message.find_element(By.CLASS_NAME, "el-message__content").text.strip()
                class_attr = message.get_attribute("class") or ""
                if "el-message--warning" in class_attr or "el-message--error" in class_attr:
                    log.warning(f"提交被拦截：{msg_text}")
                    return False
                else:
                    log.info(f"提示信息：{msg_text}")
            except TimeoutException:
                log.info("未检测到任何提示，继续流程")
            return True
        except Exception as e:
            log.error(f"点击「申请发票」按钮失败：{e}")
            return False
    
    def select_fapiao_type(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始选择发票类型...")
            WebDriverWait(self.driver, 10).until(
                EC.presence_of_element_located((By.CLASS_NAME, "el-dialog__title"))
            )
            invoice_group_select = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceGroup'] + div .el-select"))
            )
            invoice_group_select.click()
            log.info("已打开发票类型下拉框")

            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='增值税普通发票']"))
            ).click()
            log.info("已选择发票类型：增值税普通发票")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"选择发票类型失败：{e}")
            return False
    
    def select_invoice_type(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始选择发票抬头...")
            invoice_type_select = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceType'] + div .el-select"))
            )
            invoice_type_select.click()
            log.info("已打开发票抬头下拉框")

            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='电子票']"))
            ).click()
            log.info("已选择发票抬头：电子票")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"选择发票抬头失败：{e}")
            return False
    
    def select_title_type(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始选择抬头类型...")
            title_type_select = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceUpHeadType'] + div .el-select"))
            )
            title_type_select.click()
            log.info("已打开抬头类型下拉框")

            WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='个人']"))
            ).click()
            log.info("已选择抬头类型：个人")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"选择抬头类型失败：{e}")
            return False
    
    def insert_fapiao_title(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始填写发票抬头...")
            invoice_title_input = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceUpHead'] + div .el-input__inner"))
            )
            invoice_title_input.clear()
            invoice_title_input.send_keys("个人")
            log.info("已填写发票抬头：个人")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"填写发票抬头失败：{e}")
            return False
    
    def insert_fapiao_content(self, content: str):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info(f"开始填写发票内容：{content}")
            invoice_content_select = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceContext'] + div .el-select"))
            )
            invoice_content_select.click()
            log.info("已打开发票内容下拉框")

            WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@class, 'el-select-dropdown__item')]/span[text()='{content}']"))
                ).click()
            log.info(f"已选择发票内容：{content}")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"填写发票内容失败：{e}")
            return False
        
    def insert_fapiao_amount(self, amount: str):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info(f"开始填写发票金额：{amount}")
            invoice_amount_input = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='billMoney'] + div .el-input__inner"))
                )
            invoice_amount_input.clear()
            invoice_amount_input.send_keys(amount)
            log.info(f"已填写发票金额：{amount}")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"填写发票金额失败：{e}")
            return False
        
    def insert_fapiao_email(self, email: str):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info(f"开始填写接收邮箱：{email}")
            email_input = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceEmail'] + div .el-input__inner"))
            )
            email_input.clear()
            email_input.send_keys(email)
            log.info(f"已填写接收邮箱：{email}")
            time.sleep(0.5)
            return True
        except Exception as e:
            log.error(f"填写接收邮箱失败：{e}")
            return False
    
    def submit_fapiao_application(self):
        try:
            if not self.driver:
                log.error("浏览器驱动未初始化，请先调用init_driver()")
                return False
            log.info("开始提交发票申请...")
            
            # 点击提交按钮
            ok_btn = WebDriverWait(self.driver, 10).until(
                EC.element_to_be_clickable((By.XPATH,
                 '//div[@aria-label="发票申请"]//div[@class="el-dialog__footer"]//button[contains(@class,"el-button--primary")]')))
            self.driver.execute_script("arguments[0].click();", ok_btn)
            log.info("已点击提交按钮")
            time.sleep(2)
            return True
        except Exception as e:
            log.error(f"提交发票申请失败：{e}")
            return False
    
    def process_invoices(self):
        """处理发票申请的主流程"""
        error_records = []
        try:
            # 加载Excel数据
            if not os.path.exists(self.excel_path.get()):
                log.error(f"Excel文件不存在: {self.excel_path.get()}")
                self.root.after(0, lambda: messagebox.showerror("错误", "Excel文件不存在"))
                self.root.after(0, self.stop_processing)
                return
                
            self.all_data = pd.read_excel(self.excel_path.get()).to_dict('records')
            log.info(self.all_data)
            total = len(self.all_data)
            if total == 0:
                log.error("Excel文件中没有数据")
                self.root.after(0, lambda: messagebox.showerror("错误", "Excel文件中没有数据"))
                self.root.after(0, self.stop_processing)
                return
                
            log.info(f"成功加载Excel数据，共{total}条记录")
            
            # 初始化浏览器
            self.driver = self.init_driver()
            if not self.driver:
                self.root.after(0, lambda: messagebox.showerror("错误", "浏览器初始化失败"))
                self.root.after(0, self.stop_processing)
                return
                
            # 登录系统
            if not self.login(self.username_var.get(), self.password_var.get()):
                self.root.after(0, lambda: messagebox.showerror("错误", "登录失败"))
                self.root.after(0, self.stop_processing)
                return
                
            # 导航到合同页面
            if not self.choose_hetong():
                self.root.after(0, lambda: messagebox.showerror("错误", "导航到合同页面失败"))
                self.root.after(0, self.stop_processing)
                return
            
            # 处理每个合同
            for i, record in enumerate(self.all_data):
                if not self.is_running:
                    break
                    
                # 更新进度（通过主线程）
                progress = (i + 1) / total * 100
                self.root.after(0, lambda p=progress, t=f"处理中: {i+1}/{total}": self.update_progress(p, t))
                
                # 获取合同信息
                contract_no = record.get('合同编号')
                invoice_content = record.get('开票项目')
                amount = record.get('开票金额')
                
                if not contract_no:
                    log.warning("跳过缺少合同编号的记录")
                    error_records.append({**record, "错误原因": "缺少合同编号"})
                    continue
                
                log.info(f"\n===== 开始处理第{i+1}条记录: 合同编号 {contract_no} =====")
                
                # 搜索合同
                if not self.search_hetong(contract_no):
                    log.warning(f"合同 {contract_no} 未找到，添加到错误记录")
                    error_records.append({**record, "错误原因": "合同未找到"})
                    continue
                
                # 申请发票
                if not self.start_fapiao():
                    log.warning(f"合同 {contract_no} 申请发票失败")
                    error_records.append({**record, "错误原因": "申请发票失败"})
                    continue
                
                # 选择发票类型
                if not self.select_fapiao_type():
                    log.warning(f"合同 {contract_no} 选择发票类型失败")
                    error_records.append({**record, "错误原因": "选择发票类型失败"})
                    continue
                
                # 选择发票抬头
                if not self.select_invoice_type():
                    log.warning(f"合同 {contract_no} 选择发票抬头失败")
                    error_records.append({**record, "错误原因": "选择发票抬头失败"})
                    continue
                
                # 选择抬头类型
                if not self.select_title_type():
                    log.warning(f"合同 {contract_no} 选择抬头类型失败")
                    error_records.append({**record, "错误原因": "选择抬头类型失败"})
                    continue
                
                # 填写发票抬头
                if not self.insert_fapiao_title():
                    log.warning(f"合同 {contract_no} 填写发票抬头失败")
                    error_records.append({**record, "错误原因": "填写发票抬头失败"})
                    continue
                
                # 填写发票内容
                if not self.insert_fapiao_content(str(invoice_content)):
                    log.warning(f"合同 {contract_no} 填写发票内容失败")
                    error_records.append({**record, "错误原因": "填写发票内容失败"})
                    continue

                # 填写发票金额
                if not self.insert_fapiao_amount(str(amount)):
                    log.warning(f"合同 {contract_no} 填写发票金额失败")
                    error_records.append({**record, "错误原因": "填写发票金额失败"})
                    continue

                # 填写接收邮箱
                if not self.insert_fapiao_email(self.email_var.get()):
                    log.warning(f"合同 {contract_no} 填写接收邮箱失败")
                    error_records.append({**record, "错误原因": "填写接收邮箱失败"})
                    continue
                
                # 提交申请
                if not self.submit_fapiao_application():
                    log.warning(f"合同 {contract_no} 提交申请失败")
                    error_records.append({**record, "错误原因": "提交申请失败"})
                    continue
                
                log.success(f"合同 {contract_no} 处理成功")
                time.sleep(2)  # 处理间隔
                
            # 保存错误记录
            if error_records:
                df = pd.DataFrame(error_records)
                df.to_excel(self.error_file, index=False)
                log.warning(f"共{len(error_records)}条记录处理失败，已保存至: {self.error_file}")
                self.root.after(0, lambda: messagebox.showwarning("提示", f"部分记录处理失败，已保存至桌面: error_records.xlsx"))
            
            # 处理完成
            log.success("所有记录处理完毕")
            self.root.after(0, lambda: self.update_progress(100, "处理完成"))
            self.root.after(0, lambda: messagebox.showinfo("提示", "所有记录处理完毕"))
            
        except Exception as e:
            log.error(f"处理过程出错: {e}")
            self.root.after(0, lambda: self.update_progress(0, f"处理出错: {str(e)}"))
            self.root.after(0, lambda: messagebox.showerror("错误", f"处理过程出错: {str(e)}"))
        finally:
            self.root.after(0, self.stop_processing)
    
    def update_progress(self, value, text):
        """更新进度条和进度文本"""
        self.progress_var.set(value)
        self.progress_label.config(text=text)

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceApp(root)
    root.mainloop()
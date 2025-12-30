import os
import time
import pandas as pd
from datetime import datetime
from loguru import logger as log
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common import TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service as ChromeService
from read_excel import read_excel

from dotenv import find_dotenv, load_dotenv

load_dotenv(find_dotenv('.env'))

log.add(
    os.path.join(os.path.dirname(os.path.abspath(__file__)), "error.log"),
    level="ERROR",  # 只记录ERROR及以上级别
    rotation="1 day",
    retention="30 days",
    encoding="utf-8"
)


# 获取当前文件所在目录
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.dirname(current_dir)
if os.name == 'nt':
    # 拼接驱动路径
    driver_path = os.path.join(parent_dir, "lib", "win", "chromedriver.exe")
else:
    driver_path = os.path.join(parent_dir, "lib","chromedriver-linux64", "chromedriver")

log.info(f"根据系统类型选择Chrome驱动路径: {driver_path}")
service = ChromeService(executable_path=driver_path)
options = webdriver.ChromeOptions()
driver = webdriver.Chrome(service=service, options=options)
driver.maximize_window()
log.info("Chrome浏览器已启动并最大化窗口")

driver.get(os.getenv("CRM_URL") or "")
log.info(f"已访问CRM系统首页: {driver.current_url}")

EMAIL = os.getenv("EMAIL") or ""
log.info(f"加载发票接收邮箱: {EMAIL}")

def login(user_name: str, password: str) -> bool:
    try:
        log.info("开始执行登录操作...")
        # 输入用户名和密码
        username_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[placeholder="账号"]'))
        )
        log.info("找到用户名输入框")
        password_input = driver.find_element(By.CSS_SELECTOR, 'input[placeholder="密码"]')
        log.info("找到密码输入框")

        username_input.send_keys(user_name)
        log.info("已输入用户名")
        password_input.send_keys(password)
        log.info("已输入密码")

        # 点击登录按钮
        login_button = WebDriverWait(driver, 10).until(
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

def choose_hetong() -> bool:
    try:
        log.info("开始导航至待开班合同表页面...")
        # 直接访问待开班合同表的URL（无需点击菜单）
        target_url = os.getenv("HETONG_URL") or ""
        driver.get(target_url)
        log.info(f"已直接访问待开班合同表页面：{target_url}")
        
        # 等待页面加载完成（根据页面特征调整等待条件）
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, '//div[contains(@class, "el-table")]'))  # 等待表格加载
        )
        log.info("待开班合同表页面加载完成")
        time.sleep(1)
        return True
    except Exception as e:
        log.error(f"直接访问待开班合同表失败：{e}")
        return False
    
def search_hetong(target_contract_no: str) -> bool:
    try:
        log.info(f"开始搜索合同编号: {target_contract_no}")

        # 定位"合同编号"输入框
        contract_no_form = WebDriverWait(driver, 15).until(
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

        # 点击搜索按钮（适配submit-btn类名和文本）
        search_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, '//button[contains(@class, "submit-btn") and span[text()=" 搜索 "]]')
            )
        )
        driver.execute_script("arguments[0].click();", search_btn)
        log.info("已点击搜索按钮，等待搜索结果...")
        time.sleep(2)

        # 验证搜索结果
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, '//div[@class="el-table__fixed-body-wrapper"]//tbody'))
            )
            rows = driver.find_elements(By.XPATH, '//div[@class="el-table__fixed-body-wrapper"]//tbody/tr')
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
    
def start_fapiao():
    try:
        log.info("开始执行申请发票操作...")
        fixed_table_rows = WebDriverWait(driver, 15).until(
            EC.presence_of_all_elements_located(
                (By.XPATH, '//div[@class="el-table__fixed"]//div[@class="el-table__fixed-body-wrapper"]//tbody/tr')
            )
        )
        

        # 选择第1条数据（可修改index选择其他行）
        target_row_index = 0
        if target_row_index >= len(fixed_table_rows):
            raise Exception(f"目标行索引 {target_row_index} 超出实际数据行数 {len(fixed_table_rows)}")

        # 定位固定列中的复选框（可见且可点击）
        target_checkbox = fixed_table_rows[target_row_index].find_element(
            By.XPATH, './/label[@class="el-checkbox"]//span[@class="el-checkbox__input"]'
        )

        # 确保复选框可见（滚动到视图内）
        driver.execute_script("arguments[0].scrollIntoView(true);", target_checkbox)
        time.sleep(0.5)

        class_attr = target_checkbox.get_attribute("class")
        # 检查是否已勾选，未勾选则点击（用JS点击兜底）
        if class_attr is None or "is-checked" not in class_attr:
            try:
                target_checkbox.click()
            except:
                driver.execute_script("arguments[0].click();", target_checkbox)
            time.sleep(0.5)  # 等待状态更新

        # 验证勾选状态
        if "is-checked" in (target_checkbox.get_attribute("class") or ""):
            log.info("✅ 固定列复选框已成功勾选")
        else:
            raise Exception("❌ 勾选失败，复选框仍未选中")
            
        apply_btn = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//div[@class="table-tools-btnList"]//button[contains(@class, "table-tools-btn") and span[text()=" 申请发票 "]]')
            )
        )
        log.info("开始点击「申请发票」按钮...")
        driver.execute_script("arguments[0].scrollIntoView(true);", apply_btn)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", apply_btn)
        log.info("已点击「申请发票」按钮")
        time.sleep(1)
        
        try:
            message = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.el-message"))
            )
            msg_text = message.find_element(By.CLASS_NAME, "el-message__content").text.strip()
            class_attr = message.get_attribute("class") or ""
            if "el-message--warning" in class_attr or "el-message--error" in class_attr:
                log.warning(f"提交被拦截：{msg_text}")
                return False  # ✅ 关键：检测到警告/错误，返回 False
            else:
                log.info(f"提示信息：{msg_text}")
        except TimeoutException:
            log.info("未检测到任何提示，继续流程")
        return True
    except Exception as e:
        log.error(f"点击「申请发票」按钮失败：{e}")
        return False
    
def select_fapiao_type():
    try:
        log.info("开始选择发票类型...")
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CLASS_NAME, "el-dialog__title"))
        )
        invoice_group_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceGroup'] + div .el-select"))
        )
        invoice_group_select.click()
        log.info("已打开发票类型下拉框")

        # 选择"增值税普通发票"选项
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='增值税普通发票']"))
        ).click()
        log.info("已选择发票类型：增值税普通发票")
        # 等待选择生效
        time.sleep(0.5)
    except Exception as e:
        log.error(f"选择发票类型失败：{e}")


def select_invoice_type():
    try:
        log.info("开始选择发票抬头...")
        # 点击发票类型下拉框
        invoice_type_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceType'] + div .el-select"))
        )
        invoice_type_select.click()
        log.info("已打开发票抬头下拉框")

        # 选择"电子票"选项
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='电子票']"))
        ).click()
        log.info("已选择发票抬头：电子票")

        # 停留观察结果
        time.sleep(0.5)
    except Exception as e:
        log.error(f"选择发票抬头失败：{e}")


def select_title_type():
    try:
        log.info("开始选择抬头类型...")
        # 点击抬头类型下拉框
        title_type_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceUpHeadType'] + div .el-select"))
        )
        title_type_select.click()
        log.info("已打开抬头类型下拉框")

        # 选择"个人"选项
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//li[contains(@class, 'el-select-dropdown__item')]/span[text()='个人']"))
        ).click()
        log.info("已选择抬头类型：个人")

        # 停留观察结果
        time.sleep(0.5)
    except Exception as e:
        log.error(f"选择抬头类型失败：{e}")

def insert_fapiao_title():
    try:
        log.info("开始填写发票抬头...")
        invoice_title_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceUpHead'] + div .el-input__inner"))
        )
        invoice_title_input.clear()
        invoice_title_input.send_keys("个人")
        log.info("已填写发票抬头：个人")
        time.sleep(0.5)
    except Exception as e:
        log.error(f"填写发票抬头失败：{e}")

def insert_fapiao_content(content: str):
    try:
        log.info(f"开始填写发票内容：{content}")
        invoice_content_select = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceContext'] + div .el-select"))
        )
        invoice_content_select.click()
        log.info("已打开发票内容下拉框")

        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH, f"//li[contains(@class, 'el-select-dropdown__item')]/span[text()='{content}']"))
        ).click()
        log.info(f"已选择发票内容：{content}")
        time.sleep(0.5)
    except Exception as e:
        log.error(f"填写发票内容失败：{e}")

def insert_fapiao_amount(amount: str):
    try:
        log.info(f"开始填写发票金额：{amount}")
        invoice_amount_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='billMoney'] + div .el-input__inner"))
        )
        invoice_amount_input.clear()
        invoice_amount_input.send_keys(amount)
        log.info(f"已填写发票金额：{amount}")
        time.sleep(0.5)
    except Exception as e:
        log.error(f"填写发票金额失败：{e}")

def insert_fapiao_email():
    try:
        log.info("开始填写发票邮箱...")
        email_input = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "label[for='invoiceEmail'] + div .el-input__inner"))
        )
        email_input.clear()
        email_input.send_keys(EMAIL)
        log.info(f"已填写发票接收邮箱：{EMAIL}")
        time.sleep(0.5)
    except Exception as e:
        log.error(f"填写发票邮箱失败：{e}")

def submit_fapiao_info():
    try:
        log.info("开始提交发票信息...")
        ok_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable(
                (By.XPATH,
                 '//div[@aria-label="发票申请"]//div[@class="el-dialog__footer"]//button[contains(@class,"el-button--primary")]'))
        )
        driver.execute_script("arguments[0].click();", ok_btn)
        log.info("已点击提交按钮，等待提交完成...")
        time.sleep(5)
    except Exception as e:
        log.error(f"提交发票信息失败：{e}")

def insert_fapiao_info(content: str, amount: str):
    try:
        log.info("开始填写发票信息...")
        select_fapiao_type()
        select_invoice_type()
        select_title_type()
        insert_fapiao_title()
        insert_fapiao_content(content)
        insert_fapiao_amount(amount)
        insert_fapiao_email()
        log.info("发票信息填写完成")
        submit_fapiao_info()
    except Exception as e:
        log.error(f"填写发票信息失败：{e}")

def write_to_excel(data, excel_path, sheet_name="错误记录"):
    try:
        new_df = pd.DataFrame(data)

        if os.path.exists(excel_path):
            existing_df = pd.read_excel(excel_path, sheet_name=sheet_name, engine="openpyxl")
            combined_df = pd.concat([existing_df, new_df], ignore_index=True).drop_duplicates(
                subset=["合同号"], keep='first'
            )
        else:
            combined_df = new_df

        with pd.ExcelWriter(excel_path, engine="openpyxl", mode='w') as writer:
            combined_df.to_excel(writer, index=False, sheet_name=sheet_name)

        log.info(f"已将错误数据写入Excel文件: {excel_path}")
    except Exception as e:
        log.error(f"写入Excel文件失败：{e}")
    
if __name__ == "__main__":
    log.info("===== 开始执行发票申请流程 =====")

    error_data = []
    error_file = os.path.join(os.path.expanduser("~"), "Desktop", "error_records.xlsx")

    excel_path = "C:\\Users\\PC\\Desktop\\发票.xlsx"
    log.info(f"读取Excel文件: {excel_path}")
    all_data = read_excel(excel_path)
    log.info(f"Excel文件读取完成，共{len(all_data)}条记录")
    
    user_name = os.getenv("USER_NAME") or "default_user"
    password = os.getenv("PASSWORD") or "default_password"
    log.info(f"加载用户信息: {user_name}")

    try:
        if not login(user_name=user_name, password=password):
            log.error("登录失败，终止流程")
            exit(1)
        if not (choose_hetong()):
            log.error("无法进入合同列表页面，终止流程")
            exit(1)
            
        for index, row in enumerate(all_data, 1):
            contract_no, amount, content = row[0], row[1], row[2]
            log.info(f"\n===== 开始处理第{index}条记录 - 合同编号: {contract_no} =====")

            try:
                if not search_hetong(contract_no):
                    log.error(f"合同 {contract_no} 搜索失败，未找到记录")
                    error_data.append({
                        "合同号": contract_no,
                        "金额": amount,
                        "发票内容": content,
                        "错误时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "错误原因": "搜索合同失败，未找到记录"
                        })
                    continue

                if not start_fapiao():
                    log.error(f"合同 {contract_no} 点击申请发票失败")
                    error_data.append({
                        "合同号": contract_no,
                        "金额": amount,
                        "发票内容": content,
                        "错误时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "错误原因": "点击申请发票失败"
                        })
                    continue

                insert_fapiao_info(content=content, amount=amount)
                log.success(f"合同 {contract_no} 发票信息填写完成")
                time.sleep(2)  # 等待窗口关闭，酌情调整
            except Exception as e:
                error_data.append({
                    "合同号": contract_no,
                    "金额": amount,
                    "发票内容": content,
                    "错误时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "错误原因": str(e)
                    })
                log.error(f"合同 {contract_no} 处理失败，金额：{amount}，内容：{content}，错误原因：{str(e)}", exc_info=True)
                continue
        
        log.success("===== 所有合同记录处理完毕 =====")
    finally:
        log.info("关闭浏览器")
        driver.quit()
    
    if error_data:
        write_to_excel(error_data, error_file)
    log.info("===== 发票申请流程结束 =====")
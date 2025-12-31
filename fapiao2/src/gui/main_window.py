import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
from src.core.invoice_processor import InvoiceProcessor
from src.utils.logger import setup_logger


class InvoiceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("发票申请系统")
        self.root.geometry("700x500")
        self.root.resizable(True, True)

        # 初始化变量
        self.excel_path = tk.StringVar()
        self.driver_path = tk.StringVar()
        self.error_file = os.path.join(
            os.path.expanduser("~"), "Desktop", "error_records.xlsx"
        )
        self.screenshot_dir = tk.StringVar(
            value=os.path.join(os.path.expanduser("~"), "Desktop", "screenshots")
        )
        self.processor = None
        self.is_running = False

        # 初始化日志
        self.logger = setup_logger()

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

        # 主界面内容（省略部分重复代码）
        frame = ttk.LabelFrame(main_tab, text="配置")
        frame.pack(padx=10, pady=10, fill="x")

        # Excel文件选择
        ttk.Label(frame, text="Excel文件路径:").grid(
            row=0, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(frame, textvariable=self.excel_path, width=50).grid(
            row=0, column=1, padx=5, pady=5
        )
        ttk.Button(frame, text="浏览", command=self.browse_excel).grid(
            row=0, column=2, padx=5, pady=5
        )

        # 驱动文件选择
        ttk.Label(frame, text="Chrome驱动路径:").grid(
            row=1, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(frame, textvariable=self.driver_path, width=50).grid(
            row=1, column=1, padx=5, pady=5
        )
        ttk.Button(frame, text="浏览", command=self.browse_driver).grid(
            row=1, column=2, padx=5, pady=5
        )

        # 用户信息（省略部分重复代码）
        ttk.Label(frame, text="用户名:").grid(
            row=2, column=0, padx=5, pady=5, sticky="w"
        )
        self.username_var = tk.StringVar(value=os.getenv("USER_NAME") or "")
        ttk.Entry(frame, textvariable=self.username_var).grid(
            row=2, column=1, padx=5, pady=5, sticky="w"
        )

        ttk.Label(frame, text="密码:").grid(row=3, column=0, padx=5, pady=5, sticky="w")
        self.password_var = tk.StringVar(value=os.getenv("PASSWORD") or "")
        ttk.Entry(frame, textvariable=self.password_var, show="*").grid(
            row=3, column=1, padx=5, pady=5, sticky="w"
        )

        ttk.Label(frame, text="接收邮箱:").grid(
            row=4, column=0, padx=5, pady=5, sticky="w"
        )
        self.email_var = tk.StringVar(value="fapiao@cuour.org")
        ttk.Entry(frame, textvariable=self.email_var).grid(
            row=4, column=1, padx=5, pady=5, sticky="w"
        )

        ttk.Label(frame, text="截图路径:").grid(
            row=5, column=0, padx=5, pady=5, sticky="w"
        )
        ttk.Entry(frame, textvariable=self.screenshot_dir, width=50).grid(
            row=5, column=1, padx=5, pady=5
        )
        ttk.Button(frame, text="浏览", command=self.browse_screenshot_dir).grid(
            row=5, column=2, padx=5, pady=5
        )

        # 按钮区域（省略部分重复代码）
        btn_frame = ttk.Frame(main_tab)
        btn_frame.pack(padx=10, pady=10)

        self.start_btn = ttk.Button(
            btn_frame, text="开始处理", command=self.start_processing
        )
        self.start_btn.pack(side="left", padx=5)

        self.stop_btn = ttk.Button(
            btn_frame, text="停止处理", command=self.stop_processing, state="disabled"
        )
        self.stop_btn.pack(side="left", padx=5)

        # 进度条（省略部分重复代码）
        self.progress_var = tk.DoubleVar()
        progress_frame = ttk.LabelFrame(main_tab, text="进度")
        progress_frame.pack(padx=10, pady=10, fill="x")
        ttk.Progressbar(progress_frame, variable=self.progress_var, length=100).pack(
            padx=5, pady=5, fill="x"
        )
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

        # # 重定向日志到文本框（使用自定义sink）
        # text_sink = TextSink(self.log_text)
        # # 为日志器添加sink处理器
        # self.logger.addHandler(logging.StreamHandler(text_sink))

    # 以下方法与原代码相同，省略部分重复代码
    def browse_excel(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx;*.xls")], title="选择发票数据Excel文件"
        )
        if filename:
            self.excel_path.set(filename)

    def browse_screenshot_dir(self):
        dirname = filedialog.askdirectory(
            title="选择截图保存目录", initialdir=self.screenshot_dir.get()
        )
        if dirname:
            self.screenshot_dir.set(dirname)

    def browse_driver(self):
        filename = filedialog.askopenfilename(
            filetypes=[("Executable files", "*.exe"), ("All files", "*")],
            title="选择Chrome驱动文件(chromedriver)",
        )
        if filename:
            self.driver_path.set(filename)

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

        try:
            # 初始化处理器
            self.processor = InvoiceProcessor(
                excel_path=self.excel_path.get(),
                username=self.username_var.get(),
                password=self.password_var.get(),
                email=self.email_var.get(),
                error_file=self.error_file,
                logger=self.logger,
                screenshot_dir=self.screenshot_dir.get(),
                driver_path=self.driver_path.get()
                or None,  # 如果驱动路径为空，则传递None
            )

            # 检查处理器是否初始化成功
            if not self.processor:
                raise Exception("处理器初始化失败")

            # 检查数据加载是否成功
            if not self.processor.load_data():
                raise Exception("无法加载Excel数据")

        except Exception as e:
            self.logger.error(f"初始化处理器失败: {e}")
            messagebox.showerror("错误", f"初始化处理器失败: {str(e)}")
            self.is_running = False
            self.start_btn.config(state="normal")
            self.stop_btn.config(state="disabled")
            self.processor = None
            return

        # 使用线程执行耗时操作
        threading.Thread(target=self.process_invoices, daemon=True).start()

    def stop_processing(self):
        self.is_running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.progress_label.config(text="已停止")

        # 停止处理器
        if self.processor:
            self.processor.stop()

    def process_invoices(self):
        try:
            # 新增检查：确保处理器已正确初始化
            if not self.processor:
                error_msg = "处理器未正确初始化"
                self.logger.error(error_msg)
                self.root.after(0, lambda: self.update_progress(0, error_msg))
                self.root.after(0, lambda: messagebox.showerror("错误", error_msg))
                return

            # 处理发票
            self.processor.process(
                progress_callback=self.update_progress,
                stop_check=self.is_running_check,
                error_callback=self.show_error,
            )

            # 处理完成
            self.logger.success("所有记录处理完毕")
            self.root.after(0, lambda: self.update_progress(100, "处理完成"))
            self.root.after(0, lambda: messagebox.showinfo("提示", "所有记录处理完毕"))

        except Exception as e:
            self.logger.error(f"处理过程出错: {e}")
            self.root.after(0, lambda: self.update_progress(0, f"处理出错: {str(e)}"))
            self.root.after(
                0, lambda: messagebox.showerror("错误", f"处理过程出错: {str(e)}")
            )
        finally:
            self.root.after(0, self.stop_processing)

    def is_running_check(self):
        return self.is_running

    def update_progress(self, value, text):
        """更新进度条和进度文本"""
        self.progress_var.set(value)
        self.progress_label.config(text=text)

    def show_error(self, message):
        self.root.after(0, lambda: messagebox.showerror("错误", message))

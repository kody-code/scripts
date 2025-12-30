# fapiao 发票申请自动化脚本

用于自动处理发票申请流程的脚本，通过图形化界面选择Excel数据文件，自动登录CRM系统完成发票申请操作，并记录错误数据。


## 功能说明
- 读取Excel中的发票申请数据（合同号、金额等）
- 提供图形化界面（GUI），支持文件选择、账号配置和流程监控
- 自动登录CRM系统，批量提交发票申请
- 异常数据自动保存至桌面的`error_records.xlsx`
- 实时显示运行日志，支持进度条监控处理状态


## 目录结构
fapiao/
├── .gitignore      # 忽略临时文件、依赖缓存等
├── doc/            # 示例Excel文件和说明文档
│   └── 申请发票.xlsx  # 发票数据格式示例
└── src/            # 源代码目录
    ├── gui_main.py    # 图形化界面及主逻辑
    ├── read_excel.py  # Excel数据读取模块
    └── lib/        # 依赖资源（如chromedriver）
        ├── win/    # Windows系统chromedriver
        └── chromedriver-linux64/  # Linux系统chromedriver


## 环境要求
- Python 3.8+
- 依赖库：`openpyxl`、`tkinter`、`selenium`
- Chrome浏览器（需与`lib`目录中的chromedriver版本匹配）


## 使用步骤
1. 安装依赖
   pip install openpyxl selenium

2. 运行脚本
   cd src
   python gui_main.py

3. 操作流程：
   - 点击"浏览"选择包含发票数据的Excel文件（格式参考`doc/申请发票.xlsx`）
   - 输入CRM系统的用户名、密码（若已配置环境变量可自动填充）
   - 确认接收邮箱（默认fapiao@cuour.org）
   - 点击"开始处理"启动自动化流程
   - 在"运行日志"标签页查看实时进度


## 数据格式要求
- Excel文件需包含"Sheet1"工作表
- 数据从第2行开始（第1行为表头），至少包含以下字段：
  - 第1列：合同号（用于CRM系统搜索）
  - 第2列：金额
  - 第3列：发票内容

## 文件说明

1. `gui_main.py`
   - 功能：程序主入口，实现图形化界面和自动化流程控制
   - 核心类：`InvoiceApp`（封装界面组件和业务逻辑）
   - 界面组成：
     - 主界面：Excel文件选择、用户名/密码/邮箱输入、开始/停止按钮、进度条
     - 运行日志：实时显示操作过程，支持滚动查看
   - 核心功能：
     - 初始化Chrome浏览器驱动（根据系统自动选择`lib`目录下的驱动）
     - 自动登录CRM系统（通过CSS选择器定位账号密码输入框）
     - 导航至待开班合同表页面
     - 搜索合同编号并处理发票申请
     - 多线程处理避免界面卡顿，支持中途停止操作

2. `read_excel.py`
   - 功能：读取Excel文件中的发票申请数据
   - 核心函数：`read_excel(file_path)`
     - 读取范围："Sheet1"工作表，从第2行到最后一行
     - 读取逻辑：按行读取所有列数据，返回二维列表（每行对应一条申请记录）
     - 注意：若Excel中无"Sheet1"会直接退出程序


## 代码细节说明
- 日志系统：通过自定义`TextHandler`将日志输出到界面文本框，格式为"时间 - 级别 - 消息"
- 浏览器控制：使用`selenium`库，支持Chrome浏览器，可通过注释开启无头模式
- 元素定位：主要使用CSS选择器和XPath定位CRM系统页面元素
- 进度更新：处理过程中实时更新进度条和进度标签文字


## 开发备注
- 若CRM系统页面更新，需同步修改元素定位选择器（CSS/XPath）
- 新增数据字段时，需同时修改`read_excel.py`的读取逻辑和`gui_main.py`的处理逻辑
- 浏览器驱动需与本地Chrome版本匹配，更新驱动可替换`lib`目录下的对应文件
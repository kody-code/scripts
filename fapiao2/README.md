# 发票申请系统

一个基于Selenium的自动化发票申请系统，用于批量处理发票申请任务。

## 功能特点

- 批量处理发票申请
- Excel数据导入
- 自动化登录和操作
- 手动上传Chrome驱动功能（适用于打包后使用）
- 错误记录保存
- 截图保存功能

## 安装要求

- Python 3.7+
- Chrome浏览器

## 安装步骤

1. 克隆项目：
```bash
git clone https://github.com/kody-code/scripts.git
cd scripts/fapiao2
```

2. 安装依赖：
```bash
uv sync
```

## 配置

### 环境变量配置

在项目根目录创建 `.env` 文件，内容如下：

```env
CRM_URL=你的CRM系统地址
HETONG_URL=合同页面地址
USER_NAME=你的用户名
PASSWORD=你的密码
```

### 驱动配置

系统支持两种方式使用Chrome驱动：

#### 方式一：自动查找驱动（推荐）

将Chrome驱动放在指定目录：
- Windows: `项目根目录/lib/win/chromedriver.exe`
- Linux/Mac: `项目根目录/lib/chromedriver-linux64/chromedriver`

#### 方式二：手动上传驱动（打包后使用）

在GUI界面中，可以通过"Chrome驱动路径"字段手动选择驱动文件，这种方式特别适合打包后的应用程序。

## 使用方法

### 直接运行

```bash
python -m src.main
```

### GUI界面操作

1. 运行程序启动GUI界面
2. 填写必要信息：
   - Excel文件路径：包含发票数据的Excel文件
   - Chrome驱动路径：（可选）手动指定Chrome驱动路径
   - 用户名和密码：登录凭证
   - 接收邮箱：发票接收邮箱
   - 截图路径：错误截图保存路径
3. 点击"开始处理"按钮

### Excel文件格式

Excel文件应包含以下列：
- 合同编号
- 开票项目
- 开票金额

## 手动上传驱动功能说明

为了方便打包后的使用，系统增加了手动上传Chrome驱动的功能：

1. 在主界面中找到"Chrome驱动路径"字段
2. 点击"浏览"按钮
3. 选择对应的Chrome驱动文件（chromedriver）
4. 系统将使用选定的驱动文件启动浏览器

此功能解决了在不同环境下驱动路径不一致的问题，特别是在打包分发应用程序时非常有用。

## 项目结构

```
src/
├── core/
│   ├── browser_driver.py      # 浏览器驱动管理（支持手动指定驱动）
│   └── invoice_processor.py   # 发票处理逻辑
├── gui/
│   └── main_window.py         # GUI界面（包含驱动上传功能）
├── utils/
│   ├── dotenv_loader.py       # 环境变量加载
│   ├── excel_handler.py       # Excel文件处理
│   └── logger.py              # 日志处理
└── main.py                    # 程序入口
```

## 常见问题

### 驱动问题
如果遇到驱动找不到的错误，请检查：
1. 驱动版本是否与Chrome浏览器版本兼容
2. 驱动文件权限是否正确
3. 使用手动上传功能指定驱动路径

### 登录问题
确认环境变量中的用户名和密码正确无误。

## 日志

系统会将错误日志保存到 `logs/error.log` 文件中。

## 错误处理

- 系统会自动保存处理失败的记录到Excel文件
- 错误截图会保存到指定目录
- 详细的错误信息会在日志中记录
# 自用脚本仓库

本仓库用于个人日常使用的各类自动化脚本，每个文件夹对应一个独立功能的脚本，方便管理和复用。

## 仓库结构
```
.
├── .gitignore          # Git忽略文件配置
├── README.md           # 仓库说明文档
├── fapiao/             # 发票申请自动化脚本（旧版）
│   ├── .gitignore      # 脚本专属忽略配置
│   ├── README.md       # 项目说明文档
│   ├── src/            # 脚本源代码目录
│   │   ├── gui_main.py # GUI主程序入口
│   │   ├── main.py     # 主程序
│   │   └── read_excel.py # Excel读取工具
│   └── doc/            # 脚本相关文档和示例文件
└── fapiao2/            # 发票申请自动化脚本（新版）
    ├── .gitignore      # 脚本专属忽略配置
    ├── README.md       # 项目说明文档
    ├── pyproject.toml  # 项目依赖配置文件
    ├── src/            # 脚本源代码目录
    │   ├── __init__.py
    │   ├── core/       # 核心功能模块
    │   │   └── browser_driver.py # 浏览器驱动管理
    │   ├── gui/        # 图形界面模块
    │   │   └── main_window.py   # 主窗口界面
    │   ├── utils/      # 工具模块
    │   │   ├── dotenv_loader.py  # 环境变量加载器
    │   │   ├── excel_handler.py  # Excel处理器
    │   │   └── logger.py         # 日志工具
    │   └── main.py     # 主程序入口
```

## 脚本说明

| 脚本名称 | 功能描述 | 核心文件 | 状态 |
|----------|----------|----------|------|
| `fapiao` | 自动化处理发票申请流程，通过图形化界面选择Excel数据文件，自动登录CRM系统完成发票申请操作，并记录错误数据（旧版） | `fapiao/src/gui_main.py`（GUI主程序入口） | ⚠️ 维护中 |
| `fapiao2` | 自动化处理发票申请流程，通过图形化界面选择Excel数据文件，自动登录CRM系统完成发票申请操作，并记录错误数据（新版） | `fapiao2/src/main.py`（主程序入口） | ✅ 推荐使用 |

## 环境要求
- Python 3.7+
- pip 包管理器

## 安装与使用

### 通用使用说明
1. **克隆仓库到本地**
   ```bash
   git clone https://github.com/kody-code/scripts.git
   cd scripts
   ```

2. **进入对应脚本目录，根据脚本内的说明文档安装依赖并运行**
   ```bash
   # 进入发票申请脚本目录（推荐使用新版）
   cd fapiao2/

   # 使用uv
   uv sync
   ```

3. **配置环境变量**
   ```bash
   # 设置账号密码等敏感信息
   export USER_NAME="your_username"
   export PASSWORD="your_password"
   ```

4. **运行脚本**
   ```bash
   # 运行新版发票申请脚本
   cd src/
   python main.py
   ```

## 开发规范

### 项目结构标准
每个脚本目录应包含以下结构：
- `src/` - 源代码目录
- `doc/` - 文档和使用示例
- `tests/` - 单元测试（如适用）
- `config/` - 配置文件（如适用）
- `requirements.txt` 或 `pyproject.toml` - 依赖包列表
- `README.md` - 脚本详细说明

### 代码规范
- 所有敏感信息通过环境变量管理
- 代码需包含必要的注释和文档字符串
- 编写单元测试确保代码质量
- 遵循 Python PEP8 编码规范

## 注意事项
- 所有脚本均为个人自用，可能依赖特定环境配置（如系统版本、第三方工具）
- 敏感信息（如账号密码）通过环境变量（USER_NAME、PASSWORD）管理，避免硬编码
- 脚本功能可能随个人需求迭代更新，使用前请查看对应目录的最新说明
- 使用自动化脚本时请注意遵守相关网站的使用条款
- 推荐使用 `fapiao2` 项目，它是 `fapiao` 的改进版本

## 贡献指南
如果你希望扩展此仓库的功能：
1. Fork 仓库
2. 创建功能分支
3. 提交你的代码
4. 发起 Pull Request

## 许可证
此项目仅供个人学习和参考使用
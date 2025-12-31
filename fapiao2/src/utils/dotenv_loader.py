import os
from dotenv import load_dotenv


def load_env():
    """加载与src同级的.env文件（修正路径计算逻辑）"""
    # 获取当前文件的绝对路径（src/utils/dotenv_loader.py）
    current_file = os.path.abspath(__file__)
    # 第一步：获取当前文件所在目录 → src/utils
    utils_dir = os.path.dirname(current_file)
    # 第二步：获取src目录 → src
    src_dir = os.path.dirname(utils_dir)
    # 第三步：获取项目根目录（src的父目录）→ fapiao2
    project_root = os.path.dirname(src_dir)

    # 拼接正确的.env文件路径（项目根目录下）
    env_path = os.path.join(project_root, ".env")
    print(f"尝试加载.env文件路径：{env_path}")  # 调试用，可删除

    # 检查.env文件是否存在
    if not os.path.exists(env_path):
        raise FileNotFoundError(f"未找到.env文件！路径：{env_path}")

    # 加载.env文件（override=True 覆盖系统环境变量）
    load_dotenv(dotenv_path=env_path, override=True)
    print(f"成功加载.env文件：{env_path}")

    # 可选：返回加载的环境变量（方便调试）
    return {
        key: os.getenv(key) for key in os.environ.keys() if not key.startswith("PYTHON")
    }

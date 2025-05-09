from typing import Dict, Any

DEFAULT_CONFIG = {
    "check_interval": 24,  # 检查更新的间隔（小时）
    "auto_download": False,  # 是否自动下载更新
    "auto_install": False,  # 是否自动安装更新
    "github_repo": "yourusername/modbusanalyzer",  # GitHub仓库地址
    "download_dir": "~/Downloads/ModbusAnalyzer"  # 下载目录
}

def get_default_config() -> Dict[str, Any]:
    """Get default configuration"""
    return DEFAULT_CONFIG.copy() 
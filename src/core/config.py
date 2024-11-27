"""配置管理模块"""

import os
from pathlib import Path
from typing import Dict, Any, Optional
import yaml
from .exceptions import ConfigurationError

class ConfigManager:
    """配置管理器"""
    _instance = None
    _config: Dict[str, Any] = {}

    def __new__(cls):
        if cls._instance is None:
            cls._instance = super(ConfigManager, cls).__new__(cls)
        return cls._instance

    def __init__(self):
        if not self._config:
            self.load_config()

    def load_config(self, config_path: Optional[str] = None) -> None:
        """
        加载配置文件
        
        Args:
            config_path: 配置文件路径，如果为None则使用默认路径
        
        Raises:
            ConfigurationError: 配置文件加载失败时抛出
        """
        try:
            if config_path is None:
                # 获取当前文件所在目录的父目录的父目录（项目根目录）
                root_dir = Path(__file__).parent.parent.parent
                config_path = os.path.join(root_dir, 'config', 'config.yaml')

            with open(config_path, 'r', encoding='utf-8') as f:
                self._config = yaml.safe_load(f)
                
            # 设置常用配置属性
            self.excel_config = self._config.get('excel', {})
            self.feishu_config = self._config.get('feishu', {})
            self.styles_config = self._config.get('styles', {})
            self.logging_config = self._config.get('logging', {})
            self.output_config = self._config.get('output', {})
        except Exception as e:
            raise ConfigurationError(f"加载配置文件失败: {str(e)}")

    def get(self, key: str, default: Any = None) -> Any:
        """
        获取配置项
        
        Args:
            key: 配置项键名，支持点号分隔的多级键名
            default: 默认值
        
        Returns:
            配置项值
        """
        try:
            value = self._config
            for k in key.split('.'):
                value = value[k]
            return value
        except (KeyError, TypeError):
            return default

    def reload(self) -> None:
        """重新加载配置"""
        self._config = {}
        self.load_config()

# 全局配置管理器实例
config = ConfigManager()

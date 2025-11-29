# -*- coding: utf-8 -*-

import json
from typing import Any


class ConfigManager:
    """
    配置管理器，负责读取和管理配置文件。
    """

    def __init__(self, config_path: str = 'config/config.json'):
        """
        初始化配置管理器。

        :param config_path: 配置文件的路径。
        """
        self.config_path = config_path
        self.config = self.load_config()

    def load_config(self) -> dict:
        """
        加载配置文件。

        :return: 配置字典。
        """
        try:
            with open(self.config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            raise Exception(f"配置文件 '{self.config_path}' 未找到。")
        except json.JSONDecodeError:
            raise Exception(f"配置文件 '{self.config_path}' 格式错误。")

    def get_config_value(self, key: str, default: Any = None) -> Any:
        """
        获取配置项的值。

        :param key: 配置项的键，支持多级访问，例如 'app.maximumThread'。
        :param default: 如果键不存在，则返回的默认值。
        :return: 配置项的值。
        """
        keys = key.split('.')
        value = self.config
        try:
            for k in keys:
                value = value[k]
            return value
        except (KeyError, TypeError):
            return default

    def reload_config(self):
        """
        重新加载配置文件。
        """
        self.config = self.load_config()

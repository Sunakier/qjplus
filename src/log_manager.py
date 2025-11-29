# -*- coding: utf-8 -*-

import logging
from logging.handlers import TimedRotatingFileHandler
import sys
from .config_manager import ConfigManager

import os
import time


def setup_logging(config: ConfigManager) -> str:
    """
    配置全局日志记录器。
    每次运行都会创建一个新的、带时间戳的日志文件，并返回运行ID。
    :return: 本次运行的唯一ID (例如 'app_20251001_210004')
    """
    log_level = config.get_config_value(
        'logging_settings.log_level', 'INFO').upper()
    log_to_console = config.get_config_value(
        'logging_settings.log_to_console', True)
    log_to_file = config.get_config_value('logging_settings.log_to_file', True)
    log_dir = config.get_config_value('logging_settings.log_dir', 'logs')

    # 确保日志文件目录存在
    if log_dir and not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 获取根logger
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # 清除所有现有的处理器，以避免重复记录
    if root_logger.hasHandlers():
        root_logger.handlers.clear()

    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    if log_to_console:
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setFormatter(formatter)
        root_logger.addHandler(console_handler)

    run_id = f"app_{time.strftime('%Y%m%d_%H%M%S')}"
    log_file_path = None
    if log_to_file:
        log_file_path = os.path.join(log_dir, f"{run_id}.log")

        file_handler = logging.FileHandler(log_file_path, encoding='utf-8')
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)

    logging.info(f"日志系统初始化完成。日志文件: {log_file_path if log_file_path else '无'}")
    return run_id

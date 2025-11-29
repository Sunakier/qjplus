# -*- coding: utf-8 -*-

import logging
import concurrent.futures
from typing import Callable, Any
from .config_manager import ConfigManager
from .interrupt_manager import interrupt_manager

logger = logging.getLogger(__name__)


class TaskManager:
    """
    任务执行管理器，负责任务的排队、分配和执行。
    """

    def __init__(self, config: ConfigManager):
        self.config = config
        self.login_threads = self.config.get_config_value(
            'app_settings.login_threads', 1)
        self.processing_threads = self.config.get_config_value(
            'app_settings.processing_threads', 1)

        logger.debug(f"登录线程池大小设置为: {self.login_threads}")
        logger.debug(f"处理线程池大小设置为: {self.processing_threads}")

        self.login_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=self.login_threads)
        self.processing_executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=self.processing_threads)

    def submit_login_task(self, fn: Callable, *args, **kwargs) -> concurrent.futures.Future:
        """
        提交一个登录任务到登录线程池。
        """
        return self.login_executor.submit(fn, *args, **kwargs)

    def submit_processing_task(self, fn: Callable, *args, **kwargs) -> concurrent.futures.Future:
        """
        提交一个后续处理任务到任务线程池。
        """
        return self.processing_executor.submit(fn, *args, **kwargs)

    def shutdown(self):
        """
        停止所有工作线程并关闭线程池。
        """
        logger.info("正在关闭任务管理器...")
        interrupt_manager.request_shutdown()
        # 设置 wait=False，让主程序可以继续执行，而不是在这里阻塞
        self.login_executor.shutdown(wait=False)
        self.processing_executor.shutdown(wait=False)
        logger.info("任务管理器已关闭。")

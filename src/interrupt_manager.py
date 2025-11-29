# -*- coding: utf-8 -*-

import threading
import time


class InterruptedException(Exception):
    """
    自定义异常，表示任务被用户或系统中断。
    """
    pass


class InterruptManager:
    """
    一个全局中断管理器，用于协调程序的优雅关闭。
    这是一个单例模式的实现，确保整个应用只有一个中断信号源。
    """
    _instance = None
    _lock = threading.Lock()

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            with cls._lock:
                if not cls._instance:
                    cls._instance = super(InterruptManager, cls).__new__(cls)
        return cls._instance

    def __init__(self):
        # 防止重复初始化
        if not hasattr(self, '_initialized'):
            self._shutdown_event = threading.Event()
            self._initialized = True

    def request_shutdown(self):
        """
        请求关闭所有任务。
        """
        self._shutdown_event.set()

    def is_shutdown(self) -> bool:
        """
        检查是否已请求关闭。
        """
        return self._shutdown_event.is_set()

    def check_interrupted(self):
        """
        如果已请求关闭，则抛出 InterruptedException。
        """
        if self.is_shutdown():
            raise InterruptedException("任务被中断。")

    def wait(self, seconds: float) -> bool:
        """
        替代 time.sleep()，可以被中断。
        :param seconds: 等待的秒数。
        :return: 如果等待完成返回 True，如果被中断则返回 False。
        """
        return not self._shutdown_event.wait(seconds)


# 全局单例
interrupt_manager = InterruptManager()

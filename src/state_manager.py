# -*- coding: utf-8 -*-

import threading
import logging

logger = logging.getLogger(__name__)


class StateManager:
    """
    统一管理程序运行状态和计数器，确保线程安全。
    """

    def __init__(self):
        # 计数器
        self.total_users = 0
        self.login_success_count = 0
        self.login_failure_count = 0
        self.tasks_completed_count = 0
        self.tasks_failed_count = 0

        # 线程锁，用于保护计数器的原子性
        self._lock = threading.Lock()

        # 事件，用于线程间通信
        self.all_logins_done = threading.Event()
        self.all_tasks_done = threading.Event()

    def set_total_users(self, count: int):
        """
        设置需要处理的总用户数。
        """
        with self._lock:
            self.total_users = count
            logger.info(f"已设置需要处理的总用户数为: {count}")

    def _check_login_completion(self):
        """
        检查所有登录任务是否已完成。
        """
        if (self.login_success_count + self.login_failure_count) == self.total_users:
            logger.info("所有登录任务已完成。")
            self.all_logins_done.set()

    def _check_task_completion(self):
        """
        检查所有后续任务是否已完成。
        """
        tasks_processed = self.tasks_completed_count + self.tasks_failed_count
        log_msg = (
            f"检查任务完成状态: "
            f"已处理 {tasks_processed} / 登录成功 {self.login_success_count}"
        )
        logger.debug(log_msg)
        # 只有在所有登录任务都完成后，才开始检查后续任务是否完成
        all_logins_processed = (self.login_success_count + self.login_failure_count) == self.total_users
        all_tasks_processed = tasks_processed == self.login_success_count

        if all_logins_processed and all_tasks_processed:
            logger.info("所有登录任务和后续处理任务均已完成，触发 all_tasks_done 事件。")
            self.all_tasks_done.set()

    def record_login_success(self):
        """
        记录一次成功的登录。
        """
        with self._lock:
            self.login_success_count += 1
            logger.debug(f"登录成功数: {self.login_success_count}")
            self._check_login_completion()

    def record_login_failure(self):
        """
        记录一次失败的登录。
        """
        with self._lock:
            self.login_failure_count += 1
            logger.debug(f"登录失败数: {self.login_failure_count}")
            self._check_login_completion()

    def record_task_completed(self):
        """
        记录一次成功的后续任务。
        """
        with self._lock:
            self.tasks_completed_count += 1
            logger.debug(f"任务成功数: {self.tasks_completed_count}")
            self._check_task_completion()

    def record_task_failed(self):
        """
        记录一次失败的后续任务。
        """
        with self._lock:
            self.tasks_failed_count += 1
            logger.debug(f"任务失败数: {self.tasks_failed_count}")
            self._check_task_completion()

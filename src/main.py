# -*- coding: utf-8 -*-

import sys
import os
# 添加项目根目录到Python路径中
project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, project_root)

import requests
import datetime
from typing import Dict, Any, List
import logging
from src.config_manager import ConfigManager
from src.log_manager import setup_logging
from src.data_processor import DataProcessor
from src.browser_manager import BrowserManager
from src.task_manager import TaskManager
from src.login_manager import LoginManager
from src.report_generator import ReportGenerator
from src.utils import retry, TaskStatusCode
from src.user_data import UserData
from src.course_completer import CourseCompleter
from src.state_manager import StateManager
from src.progress_manager import ProgressManager


class MainController:
    def __init__(self, run_id: str):
        self.run_id = run_id
        self.config = ConfigManager(config_path='config/config.json')
        self.data_processor = DataProcessor(
            data_folder=self.config.get_config_value('app_settings.user_data_files'))
        self.browser_manager = BrowserManager(self.config)
        self.task_manager = TaskManager(self.config)
        self.login_manager = LoginManager()
        self.report_generator = ReportGenerator(run_id=self.run_id)
        self.course_completer = CourseCompleter(self.config)
        self.state_manager = StateManager()
        self.progress_manager = ProgressManager()  # 初始化进度管理器
        self.users: List[UserData] = []

        # 动态应用重试逻辑
        self._apply_retry_logic()

    def _login_task(self, user: UserData):
        import threading
        thread_id = threading.get_ident()
        logging.debug(f"[线程: {thread_id}] 开始处理用户 '{user.username}' 的登录任务。")
        user.login_start_time = datetime.datetime.now()
        browser = self.browser_manager.get_browser()
        if not browser:
            user.status_code = TaskStatusCode.LOGIN_FAILED_OTHER
            raise Exception("无法获取浏览器实例。")

        try:
            self.login_manager.login(browser.page, user)
            if user.status_code != TaskStatusCode.LOGIN_SUCCESS:
                raise Exception(f"登录失败，状态码: {user.status_code.name}")
        except Exception as e:
            # 保留由 login_manager 设置的精确状态码
            if user.status_code in [TaskStatusCode.LOGIN_SUCCESS, TaskStatusCode.PENDING]:
                user.status_code = TaskStatusCode.LOGIN_FAILED_OTHER
            logging.debug(f"用户 {user.username} 登录尝试失败: {e}")
            raise  # 重新抛出异常以触发重试
        finally:
            user.login_end_time = datetime.datetime.now()
            self.browser_manager.release_browser(browser)
            logging.debug(f"[线程: {thread_id}] 用户 '{user.username}' 的登录任务处理完毕。")

    def _processing_task(self, user: UserData):
        user.task_start_time = datetime.datetime.now()
        try:
            logging.info(f"开始为用户 {user.username} 处理课程完成任务...")
            newly_completed, total_completed = self.course_completer.complete_courses(user)
            user.newly_completed_courses = newly_completed
            user.total_completed_courses = total_completed
            if newly_completed > 0:
                user.status_code = TaskStatusCode.TASK_SUCCESS
            else:
                user.status_code = TaskStatusCode.TASK_COMPLETED_NO_NEW_COURSES
        except Exception as e:
            user.status_code = TaskStatusCode.TASK_FAILED_OTHER
            logging.warning(f"用户 {user.username} 的课程完成任务失败: {e}")
            raise  # 重新抛出异常以触发重试
        finally:
            user.task_end_time = datetime.datetime.now()

    def _login_task_wrapper(self, user: UserData):
        """包装登录任务，处理重试后的最终结果统计。"""
        try:
            self._login_task(user)
            self.state_manager.record_login_success()
            # 登录成功后，提交后续处理任务
            self.task_manager.submit_processing_task(self._processing_task_wrapper, user)
        except Exception as e:
            logging.error(f"用户 {user.username} 登录任务在所有重试后最终失败: {e}")
            self.state_manager.record_login_failure()

    def _processing_task_wrapper(self, user: UserData):
        """包装处理任务，处理重试后的最终结果统计。"""
        try:
            self._processing_task(user)
            self.state_manager.record_task_completed()
            # 如果任务成功完成，标记用户为已完成
            if user.status_code in [TaskStatusCode.TASK_SUCCESS, TaskStatusCode.TASK_COMPLETED_NO_NEW_COURSES]:
                self.progress_manager.mark_user_completed(user)
        except Exception as e:
            logging.error(f"用户 {user.username} 处理任务在所有重试后最终失败: {e}")
            self.state_manager.record_task_failed()

    def _apply_retry_logic(self):
        """从配置中读取重试设置并动态应用到任务方法上"""
        login_retry_settings = self.config.get_config_value('retry_settings.login_task', {})
        processing_retry_settings = self.config.get_config_value('retry_settings.processing_task', {})

        if login_retry_settings.get('max_retries', 0) > 0:
            self._login_task = retry(
                max_retries=login_retry_settings['max_retries'],
                delay=login_retry_settings.get('delay_seconds', 1)
            )(self._login_task)

        if processing_retry_settings.get('max_retries', 0) > 0:
            self._processing_task = retry(
                max_retries=processing_retry_settings['max_retries'],
                delay=processing_retry_settings.get('delay_seconds', 1)
            )(self._processing_task)

    def run(self):
        logging.info("自动化任务开始...")

        users_raw_data = self.data_processor.read_user_data()
        if not users_raw_data:
            logging.warning("没有找到任何用户数据，程序退出。")
            return

        self.users = [UserData.from_dict(data) for data in users_raw_data]

        # 使用进度管理器筛选出未完成的用户
        self.users = self.progress_manager.get_uncompleted_users(self.users)

        if not self.users:
            logging.info("所有用户任务已完成，无需重复执行。")
            return

        logging.info(f"本次需要处理 {len(self.users)} 个用户（总共有 {len(users_raw_data)} 个用户）")
        self.state_manager.set_total_users(len(self.users))

        for user in self.users:
            #logging.debug(f"提交用户 {user.username} 的登录任务。")
            self.task_manager.submit_login_task(self._login_task_wrapper, user)

        # 等待所有登录任务完成
        logging.info("主线程开始等待 all_logins_done 事件...")
        self.state_manager.all_logins_done.wait()
        logging.info("主线程接收到 all_logins_done 事件，关闭浏览器实例...")
        self.browser_manager.close_all_browsers()

        # 等待所有后续任务完成
        logging.info("主线程开始等待 all_tasks_done 事件...")
        self.state_manager.all_tasks_done.wait()
        logging.info("主线程接收到 all_tasks_done 事件，关闭任务管理器...")
        self.task_manager.shutdown()

        # 标记所有已完成的用户
        completed_users = [user for user in self.users if user.status_code in [TaskStatusCode.TASK_SUCCESS, TaskStatusCode.TASK_COMPLETED_NO_NEW_COURSES]]
        if completed_users:
            self.progress_manager.batch_mark_users_completed(completed_users)
            logging.info(f"已标记 {len(completed_users)} 个用户为已完成状态")

        logging.info("准备生成报告...")
        self.report_generator.export_report(self.users)

        logging.info("自动化任务结束。")


if __name__ == "__main__":
    # 在启动时配置日志
    config = ConfigManager(config_path='config/config.json')
    run_id = setup_logging(config)

    controller = MainController(run_id=run_id)
    controller.run()

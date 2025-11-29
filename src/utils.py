# -*- coding: utf-8 -*-

import time
import logging
from functools import wraps
from typing import Optional, Union, Any
from .interrupt_manager import interrupt_manager, InterruptedException
from enum import Enum, auto

# 获取与此模块同名的logger
logger = logging.getLogger(__name__)


def retry(max_retries: int = 3, delay: int | float = 1):
    """
    一个装饰器，用于在函数执行失败时自动重试。

    :param max_retries: 最大重试次数。
    :param delay: 每次重试之间的延迟（秒）。
    """
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_retries):
                # 在每次尝试前检查中断信号
                interrupt_manager.check_interrupted()
                try:
                    return func(*args, **kwargs)
                except (KeyboardInterrupt, InterruptedException):
                    logger.info("任务被用户中断，停止重试。")
                    raise InterruptedException("任务在重试过程中被中断。")
                except Exception as e:
                    last_exception = e
                    logger.warning(
                        f"函数 {func.__name__} 第 {attempt + 1} 次执行失败: {e}。将在 {delay} 秒后重试...")
                    # 使用可中断的 sleep
                    if not interrupt_manager.wait(delay):
                        logger.info("等待重试时被中断。")
                        raise InterruptedException("任务在重试等待期间被中断。")

            logger.error(f"函数 {func.__name__} 在 {max_retries} 次重试后仍然失败。")

            if last_exception:
                raise last_exception
            # 如果循环完成且没有异常（理论上不应该发生），则引发一个通用异常
            raise Exception(f"函数 {func.__name__} 重试失败，但没有捕获到异常。")
        return wrapper
    return decorator


def convert_grade(grade_input: Any) -> Optional[str]:
    """
    将输入的年级转换为标准年级格式。

    :param grade_input: 年级输入，可以是字符串或数字。
    :return: 标准年级格式的字符串，如果无法转换则返回None。
    """
    if not grade_input:
        return None

    try:
        grade_str = str(grade_input).strip()
        if grade_str.isdigit():
            grade_num = int(grade_str)
            grades = {
                5: "五年级", 6: "六年级", 7: "初一", 8: "初二", 9: "初三",
                10: "高一", 11: "高二", 12: "中职一", 13: "中职二"
            }
            return grades.get(grade_num)

        if "中职" in grade_str or "中专" in grade_str:
            if "一" in grade_str or "1" in grade_str:
                return "中职一"
            if "二" in grade_str or "2" in grade_str:
                return "中职二"
        elif "高" in grade_str:
            if "一" in grade_str or "1" in grade_str:
                return "高一"
            if "二" in grade_str or "2" in grade_str:
                return "高二"
        elif "初" in grade_str:
            if "一" in grade_str or "1" in grade_str or "七" in grade_str or "7" in grade_str:
                return "初一"
            if "二" in grade_str or "2" in grade_str or "八" in grade_str or "8" in grade_str:
                return "初二"
            if "三" in grade_str or "3" in grade_str or "九" in grade_str or "9" in grade_str:
                return "初三"
        elif "小" in grade_str or "年级" in grade_str:
            if "五" in grade_str or "5" in grade_str:
                return "五年级"
            if "六" in grade_str or "6" in grade_str:
                return "六年级"
    except:
        return None

    return None


class TaskStatusCode(Enum):
    """
    定义任务执行过程中的所有可能状态码。
    """
    PENDING = auto()
    LOGIN_SUCCESS = auto()
    LOGIN_FAILED_CREDENTIALS = auto()
    LOGIN_FAILED_SLIDER = auto()
    LOGIN_FAILED_OTHER = auto()
    TASK_SUCCESS = auto()
    TASK_COMPLETED_NO_NEW_COURSES = auto()
    TASK_FAILED_OTHER = auto()

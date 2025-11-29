# -*- coding: utf-8 -*-

import datetime
import os
from typing import List, Dict, Any
import logging
import pandas as pd
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text
from .user_data import UserData
from .utils import TaskStatusCode

logger = logging.getLogger(__name__)


class ReportGenerator:
    """
    执行报告生成器，负责根据最终的用户数据列表生成多种格式的报告。
    """

    STATUS_CODE_MAPPING = {
        TaskStatusCode.PENDING: "未开始",
        TaskStatusCode.LOGIN_SUCCESS: "登录成功",
        TaskStatusCode.LOGIN_FAILED_CREDENTIALS: "登录失败-账号密码错误",
        TaskStatusCode.LOGIN_FAILED_SLIDER: "登录失败-滑块验证失败",
        TaskStatusCode.LOGIN_FAILED_OTHER: "登录失败-其他原因",
        TaskStatusCode.TASK_SUCCESS: "完成任务",
        TaskStatusCode.TASK_COMPLETED_NO_NEW_COURSES: "完成任务-无新课程",
        TaskStatusCode.TASK_FAILED_OTHER: "执行失败-其他原因",
    }

    def __init__(self, run_id: str, report_dir: str = "output"):
        self.run_id = run_id
        self.report_dir = report_dir
        self.start_time = datetime.datetime.now()
        self._ensure_report_dir_exists()

    def _ensure_report_dir_exists(self):
        if not os.path.exists(self.report_dir):
            os.makedirs(self.report_dir)
            logger.info(f"创建报告目录: {self.report_dir}")

    def _interpret_status(self, user: UserData) -> str:
        return self.STATUS_CODE_MAPPING.get(user.status_code, "未知状态")

    def _calculate_duration(self, user: UserData) -> float:
        if user.login_start_time and user.task_end_time:
            return (user.task_end_time - user.login_start_time).total_seconds()
        if user.login_start_time and user.login_end_time:
            return (user.login_end_time - user.login_start_time).total_seconds()
        return 0.0

    def export_text_report(self, users: List[UserData]):
        console = Console(record=True, width=120)
        end_time = datetime.datetime.now()

        console.print(Panel(Text("自 动 化 执 行 报 告", justify="center", style="bold blue"), border_style="blue"))

        # 任务统计
        duration = end_time - self.start_time
        total_users = len(users)
        successful_logins = sum(1 for u in users if u.status_code in [TaskStatusCode.LOGIN_SUCCESS, TaskStatusCode.TASK_SUCCESS, TaskStatusCode.TASK_COMPLETED_NO_NEW_COURSES])
        total_completed_courses = sum(u.total_completed_courses for u in users)

        summary_table = Table(show_header=False, box=None, padding=(0, 2))
        summary_table.add_column(style="cyan")
        summary_table.add_column(style="magenta")
        summary_table.add_row("开始时间:", self.start_time.strftime('%Y-%m-%d %H:%M:%S'))
        summary_table.add_row("结束时间:", end_time.strftime('%Y-%m-%d %H:%M:%S'))
        summary_table.add_row("总耗时:", str(duration))
        summary_table.add_row("-" * 20, "-" * 20)
        summary_table.add_row("总人数:", str(total_users))
        summary_table.add_row("成功人数:", str(successful_logins))
        summary_table.add_row("失败人数:", str(total_users - successful_logins))
        summary_table.add_row("总完成课程数:", str(total_completed_courses))
        console.print(Panel(summary_table, title="[bold cyan]任务统计[/bold cyan]", border_style="cyan"))

        # 详情
        details_table = Table(title="执行详情", header_style="bold green", border_style="green")
        headers = ["姓名", "账号", "结果", "完成课程数", "开始时间", "完整用时（秒）"]
        for header in headers:
            details_table.add_column(header)
        
        for user in users:
            start_time_str = user.login_start_time.strftime('%Y-%m-%d %H:%M:%S') if user.login_start_time else "N/A"
            duration = f"{self._calculate_duration(user):.2f}"
            result_text = self._interpret_status(user)
            completed_courses_str = f"{user.newly_completed_courses}(本次)/{user.total_completed_courses}(总共)"
            details_table.add_row(user.name, user.username, result_text, completed_courses_str, start_time_str, duration)
        
        console.print(details_table)

        report_path = os.path.join(self.report_dir, f"{self.run_id}.txt")
        with open(report_path, "w", encoding="utf-8") as f:
            f.write(console.export_text())
        logger.info(f"文本报告已生成: {report_path}")

    def export_excel_report(self, users: List[UserData]):
        if not users:
            logger.warning("没有用户数据可供生成 Excel 报告。")
            return

        report_data = []
        for user in users:
            report_data.append({
                "姓名": user.name,
                "账号": user.username,
                "结果": self._interpret_status(user),
                "完成课程数": f"{user.newly_completed_courses}(本次)/{user.total_completed_courses}(总共)",
                "开始时间": user.login_start_time.strftime('%Y-%m-%d %H:%M:%S') if user.login_start_time else None,
                "完整用时（秒）": self._calculate_duration(user)
            })

        df = pd.DataFrame(report_data)
        report_path = os.path.join(self.report_dir, f"{self.run_id}.xlsx")
        
        try:
            with pd.ExcelWriter(report_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='执行详情')
            logger.info(f"Excel 报告已生成: {report_path}")
        except Exception as e:
            logger.error(f"生成 Excel 报告失败: {e}")

    def export_report(self, users: List[UserData]):
        self.export_text_report(users)
        self.export_excel_report(users)

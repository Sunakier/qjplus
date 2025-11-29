# -*- coding: utf-8 -*-

from dataclasses import dataclass, field
from typing import Dict, Any, Optional
import datetime
from .utils import convert_grade, TaskStatusCode


@dataclass
class UserData:
    """
    用于存储和传递用户数据的自定义数据类型。
    """
    username: str
    password: str
    name: str
    grade: Optional[str] = None  # 改为 Optional，因为将在 post_init 中设置

    # 登录成功后才会填充的字段
    reqtoken: Optional[str] = None
    sid: Optional[str] = None
    cookies: Optional[Dict[str, Any]] = field(default_factory=dict)
    login_time: Optional[datetime.datetime] = None

    # 任务执行过程中的状态字段
    status_code: TaskStatusCode = TaskStatusCode.PENDING
    newly_completed_courses: int = 0
    total_completed_courses: int = 0
    login_start_time: Optional[datetime.datetime] = None
    login_end_time: Optional[datetime.datetime] = None
    task_start_time: Optional[datetime.datetime] = None
    task_end_time: Optional[datetime.datetime] = None
    # 原始数据，以备不时之需
    raw_data: Dict[str, Any] = field(default_factory=dict)

    def __post_init__(self):
        """
        初始化后，自动转换年级并填充 raw_data。
        """
        # 转换年级
        self.grade = convert_grade(self.raw_data.get('年级'))

        # 确保 raw_data 包含所有基本信息
        if not self.raw_data:
            self.raw_data = {
                '账号': self.username,
                '密码': self.password,
                '姓名': self.name,
                '年级': self.raw_data.get('年级')  # 使用原始年级
            }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'UserData':
        """
        从字典创建 UserData 实例。
        年级转换将在 __post_init__ 中自动完成。
        """
        return cls(
            username=data.get('账号') or '',
            password=data.get('密码') or '',
            name=data.get('姓名') or '',
            raw_data=data
        )

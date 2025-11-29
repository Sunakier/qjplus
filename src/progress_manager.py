# -*- coding: utf-8 -*-

import json
import os
import threading
from typing import Dict, Set
from src.user_data import UserData
import logging

logger = logging.getLogger(__name__)


class ProgressManager:
    """
    进度管理器，用于记录和恢复用户任务进度。
    """
    
    def __init__(self, progress_file_path: str = "cache/user_progress.json"):
        self.progress_file_path = progress_file_path
        self.completed_users_file_path = "cache/completed_users.json"
        self.completed_users: Set[str] = set()
        self.lock = threading.Lock()  # 确保线程安全
        self._ensure_cache_dir_exists()
        self._load_completed_users()

    def _ensure_cache_dir_exists(self):
        """确保缓存目录存在"""
        # 确保两个文件的目录都存在
        progress_dir = os.path.dirname(self.progress_file_path)
        completed_dir = os.path.dirname(self.completed_users_file_path)

        for cache_dir in [progress_dir, completed_dir]:
            if cache_dir and not os.path.exists(cache_dir):
                os.makedirs(cache_dir, exist_ok=True)

    def _load_completed_users(self):
        """从文件加载已完成的用户列表"""
        try:
            if os.path.exists(self.completed_users_file_path):
                with open(self.completed_users_file_path, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 合并新旧已完成用户列表，而不是覆盖
                    old_completed_users = set(data.get('completed_users', []))
                    # 将旧的完成用户添加到当前集合中
                    self.completed_users.update(old_completed_users)
                logger.info(f"已加载 {len(old_completed_users)} 个历史已完成的用户，当前总计: {len(self.completed_users)} 个")
        except Exception as e:
            logger.error(f"加载已完成用户列表失败: {e}")
            self.completed_users = set()

    def _save_completed_users(self):
        """保存已完成的用户列表到文件"""
        try:
            data = {
                'completed_users': list(self.completed_users)
            }
            with open(self.completed_users_file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"保存已完成用户列表失败: {e}")

    def is_user_completed(self, user: UserData) -> bool:
        """检查用户是否已完成所有任务"""
        user_key = f"{user.name}-{user.username}"
        with self.lock:
            return user_key in self.completed_users

    def mark_user_completed(self, user: UserData):
        """标记用户已完成所有任务"""
        user_key = f"{user.name}-{user.username}"
        with self.lock:
            self.completed_users.add(user_key)
            self._save_completed_users()
        logger.info(f"用户 {user_key} 已标记为完成")

    def get_uncompleted_users(self, users: list) -> list:
        """从用户列表中筛选出未完成的用户"""
        uncompleted_users = []
        for user in users:
            if not self.is_user_completed(user):
                uncompleted_users.append(user)
        logger.info(f"总共 {len(users)} 个用户，{len(uncompleted_users)} 个未完成，{len(users) - len(uncompleted_users)} 个已完成")
        return uncompleted_users
    
    def batch_mark_users_completed(self, users: list):
        """批量标记用户完成"""
        with self.lock:
            for user in users:
                user_key = f"{user.name}-{user.username}"
                self.completed_users.add(user_key)
            self._save_completed_users()
        logger.info(f"批量标记 {len(users)} 个用户为完成状态")
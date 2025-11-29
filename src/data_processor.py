# -*- coding: utf-8 -*-

import os
import pandas as pd
from typing import List, Dict, Any, Optional
import logging

logger = logging.getLogger(__name__)


class DataProcessor:
    """
    数据处理模块，负责读取和处理用户数据文件。
    """

    def __init__(self, data_folder: str):
        """
        初始化数据处理器。

        :param data_folder: 存放用户数据文件的文件夹路径。
        """
        self.data_folder = data_folder

    def read_user_data(self) -> List[Dict[str, Any]]:
        """
        从指定文件夹中读取所有支持的表格文件，并提取用户数据。

        :return: 用户数据列表，每个用户是一个字典。
        """
        all_users = []
        supported_formats = ('.xls', '.xlsx', '.csv')
        processed_files = 0

        if not os.path.isdir(self.data_folder):
            raise FileNotFoundError(f"数据文件夹 '{self.data_folder}' 不存在。")

        for filename in os.listdir(self.data_folder):
            if filename.endswith(supported_formats):
                file_path = os.path.join(self.data_folder, filename)
                processed_files += 1
                try:
                    if filename.endswith('.csv'):
                        df = pd.read_csv(file_path)
                        logger.debug(
                            f"处理文件: {filename}, 表: CSV, 处理数量: {len(df)}")
                    else:
                        df = pd.read_excel(file_path, sheet_name=None)
                        if isinstance(df, dict):
                            # 处理多sheet文件
                            sheet_names = list(df.keys())
                            combined_df = pd.concat(
                                df.values(), ignore_index=True)
                            df = combined_df
                            logger.debug(
                                f"处理文件: {filename}, 表: {sheet_names}, 处理数量: {len(df)}")
                        else:
                            logger.debug(
                                f"处理文件: {filename}, 表: 单表, 处理数量: {len(df)}")

                    # 识别表头
                    df.columns = df.columns.str.strip()

                    # 提取数据
                    # 假设列名为 '姓名', '账号', '密码', '年级'
                    # 这里可以根据实际情况调整
                    df = df.rename(columns={
                        '学生姓名': '姓名',
                        '账户': '账号',
                        '帐号': '账号',
                    })

                    required_columns = {'姓名', '账号', '密码', '年级'}
                    if not required_columns.issubset(df.columns):
                        logger.warning(f"文件 {filename} 的列名不完全匹配，将尝试进行部分匹配。")

                    users = df.to_dict('records')
                    all_users.extend(users)

                except Exception as e:
                    logger.error(f"处理文件 '{filename}' 时出错: {e}")

        # 数据清洗和去重
        unique_users = list(
            {v['账号']: v for v in all_users if v.get('账号')}.values())

        # 记录处理结果信息
        logger.info(f"处理文件数量: {processed_files}, 去重后总数量: {len(unique_users)}")

        # 年级转换
        return unique_users

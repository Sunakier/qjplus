# -*- coding: utf-8 -*-

import threading
import time
from queue import Queue, Empty
from typing import Optional
import logging
from DrissionPage import ChromiumPage, ChromiumOptions
from DrissionPage.errors import PageDisconnectedError, BrowserConnectError, CDPError, ContextLostError
from .config_manager import ConfigManager

logger = logging.getLogger(__name__)


class BrowserInstance:
    def __init__(self, page: ChromiumPage, instance_id: str, max_usage: int):
        self.page = page
        self.id = instance_id
        self.usage_count = 0
        self.max_usage = max_usage
        self.last_in_pool_time = time.time()  # 记录最后在池中的时间

    def is_expired(self) -> bool:
        return self.usage_count >= self.max_usage

    def increment_usage(self):
        self.usage_count += 1

    def update_last_in_pool_time(self):
        """更新最后在池中的时间"""
        self.last_in_pool_time = time.time()

    def is_healthy(self) -> bool:
        """检查浏览器实例是否健康（可以正常访问）"""
        for attempt in range(3):  # 重试3次
            try:
                # 尝试访问页面状态来判断是否健康
                if self.page:
                    # 简单的健康检查：尝试获取页面URL
                    _ = self.page.url
                    return True
                return False
            except (PageDisconnectedError, BrowserConnectError, CDPError, ContextLostError) as e:
                # 这些异常明确表示浏览器连接已断开或出现严重问题
                logger.warning(f"检测到浏览器连接失败，浏览器实例 {self.id} 需要重建: {e}")
                return False  # 立即返回False，不需要重试
            except Exception as e:
                if attempt < 2:  # 如果不是最后一次尝试
                    time.sleep(0.5)  # 等待一段时间再重试
                    continue
                else:  # 最后一次尝试也失败了
                    logger.debug(f"浏览器实例 {self.id} 健康检查失败: {e}")
                    return False

    def close(self):
        if self.page:
            try:
                self.page.quit()
            except Exception as e:
                logger.warning(f"关闭浏览器实例 {self.id} 时出错: {e}")

    def safe_execute(self, operation, *args, **kwargs):
        """
        安全执行浏览器操作，如果操作失败则尝试重建浏览器
        :param operation: 要执行的操作函数
        :param args: 操作函数的位置参数
        :param kwargs: 操作函数的关键字参数
        :return: 操作结果或None
        """
        try:
            # 直接执行操作
            return operation(*args, **kwargs)
        except (PageDisconnectedError, BrowserConnectError, CDPError, ContextLostError) as e:
            # 这些异常明确表示浏览器连接已断开或出现严重问题
            logger.warning(f"浏览器实例 {self.id} 执行操作时检测到连接失败: {e}")
            # 标记为不健康
            return None
        except Exception as e:
            error_msg = str(e)
            # 检查错误消息是否包含浏览器连接失败的关键字
            if ("浏览器连接失败" in error_msg or
                "connection" in error_msg.lower() or
                "timeout" in error_msg.lower() or
                "disconnected" in error_msg.lower()):

                logger.warning(f"浏览器实例 {self.id} 执行操作时检测到连接失败: {e}")
                # 标记为不健康
                return None
            else:
                # 其他错误，直接抛出
                raise e


class BrowserManager:
    """
    浏览器实例管理器，负责浏览器实例的创建、销毁和生命周期管理。
    """

    def __init__(self, config: ConfigManager):
        self.config = config
        self.browser_pool = Queue()
        self.max_browsers = self.config.get_config_value(
            'app_settings.max_browsers', 3)
        logger.debug(f"浏览器池大小设置为: {self.max_browsers}")
        self._initialize_pool()

    def _create_browser_instance(self, instance_id: str) -> Optional[BrowserInstance]:
        browser_path = self.config.get_config_value(
            'browser_settings.browser_path')
        max_usage = self.config.get_config_value(
            'browser_settings.max_usage_per_instance', 10)

        # 从配置中读取启动超时时间，默认为60秒
        start_timeout = self.config.get_config_value(
            'browser_settings.browser_start_timeout', 60)

        logger.info(f"尝试启动浏览器实例 {instance_id}")

        start_time = time.time()  # 在 try 之外初始化 start_time
        try:
            # 设置启动超时
            co = ChromiumOptions().set_browser_path(browser_path).auto_port()
            page = ChromiumPage(addr_or_opts=co)

            # 简单的健康检查，确保浏览器已准备好
            page.get('about:blank')  # 确保页面可以访问

            logger.info(f"浏览器实例 {instance_id} 启动成功")
            return BrowserInstance(page, instance_id, max_usage)
        except Exception as e:
            elapsed = time.time() - start_time
            if elapsed >= start_timeout:
                logger.error(f"创建浏览器实例 {instance_id} 超时 ({start_timeout}秒): {e}")
            else:
                logger.error(f"创建浏览器实例 {instance_id} 失败: {e}")
            return None

    def _initialize_pool(self):
        for i in range(self.max_browsers):
            instance = self._create_browser_instance(f"browser_{i}")
            if instance:
                self.browser_pool.put(instance)
        logger.info(f"浏览器实例池初始化完成，共 {self.browser_pool.qsize()} 个实例。")

    def get_browser(self) -> Optional[BrowserInstance]:
        """
        从池中获取一个可用的浏览器实例。
        """
        max_retries = self.max_browsers  # 最大重试次数，避免无限循环
        retry_count = 0

        while retry_count <= max_retries:
            try:
                thread_id = threading.get_ident()
                logger.debug(f"[线程: {thread_id}] 正在等待获取浏览器... 池中剩余: {self.browser_pool.qsize()}")
                instance = self.browser_pool.get(block=True)  # 改为无限期阻塞等待
                logger.debug(f"[线程: {thread_id}] 成功获取到浏览器 {instance.id}。池中剩余: {self.browser_pool.qsize()}")

                # 检查实例是否过期或不健康，如果是，则销毁并创建一个新的
                needs_replacement = False
                reason = ""

                if instance.is_expired():
                    needs_replacement = True
                    reason = "已达到最大使用次数"
                elif not instance.is_healthy():
                    needs_replacement = True
                    reason = "健康检查失败"

                if needs_replacement:
                    logger.debug(f"浏览器实例 {instance.id} {reason}，将进行销毁和重建。")
                    instance.close()
                    new_instance = self._create_browser_instance(instance.id)
                    if new_instance:
                        # 将新实例放入池中，并重新获取另一个实例
                        self.browser_pool.put(new_instance)
                        retry_count += 1  # 增加重试计数
                        continue  # 继续循环以获取另一个实例
                    else:
                        logger.error(f"无法创建新的浏览器实例 {instance.id}")
                        return None

                return instance
            except Empty:
                # 理论上，在 block=True 的情况下，这里不应该被触发，但保留以防万一
                logger.error("获取浏览器实例时队列为空，这不应该发生。")
                return None

        logger.error(f"获取健康浏览器实例失败，已达到最大重试次数 {max_retries}")
        return None

    def release_browser(self, instance: BrowserInstance, force_recreate=False):
        """
        将浏览器实例释放回池中。
        :param instance: 要释放的浏览器实例
        :param force_recreate: 是否强制重新创建浏览器实例
        """
        instance.increment_usage()
        thread_id = threading.get_ident()

        # 如果需要强制重建或实例过期或不健康
        needs_replacement = force_recreate

        if not needs_replacement:  # 如果不需要强制重建，再检查其他条件
            if instance.is_expired():
                needs_replacement = True
                reason = "已达到最大使用次数"
            elif not instance.is_healthy():
                needs_replacement = True
                reason = "健康检查失败"

        if needs_replacement:
            if not force_recreate:  # 如果不是强制重建，记录原因
                logger.debug(f"浏览器实例 {instance.id} {reason}，将进行销毁和重建。")
            else:
                logger.debug(f"浏览器实例 {instance.id} 启动异常，将进行销毁和重建。")

            instance.close()
            new_instance = self._create_browser_instance(instance.id)
            if new_instance:
                new_instance.update_last_in_pool_time()  # 更新最后在池中的时间
                self.browser_pool.put(new_instance)
        else:
            instance.update_last_in_pool_time()  # 更新最后在池中的时间
            self.browser_pool.put(instance)

        logger.debug(f"[线程: {thread_id}] 释放浏览器 {instance.id} 回池中。池中剩余: {self.browser_pool.qsize()}")

    def handle_disconnected_instance(self, instance: BrowserInstance):
        """
        处理已断开连接的浏览器实例，立即重建
        :param instance: 已断开连接的浏览器实例
        """
        logger.warning(f"处理已断开连接的浏览器实例 {instance.id}")

        # 立即关闭原实例
        instance.close()

        # 创建新实例并放入池中
        new_instance = self._create_browser_instance(instance.id)
        if new_instance:
            new_instance.update_last_in_pool_time()
            self.browser_pool.put(new_instance)
            logger.info(f"浏览器实例 {instance.id} 已成功重建并放回池中")
        else:
            logger.error(f"无法重建浏览器实例 {instance.id}")

    def close_all_browsers(self):
        """
        关闭池中所有的浏览器实例。
        """
        logger.info("正在关闭所有浏览器实例...")
        while not self.browser_pool.empty():
            try:
                instance = self.browser_pool.get_nowait()
                instance.close()
            except Empty:
                break
        logger.info("所有浏览器实例已关闭")

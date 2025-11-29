# -*- coding: utf-8 -*-

import time
import logging
from DrissionPage import ChromiumPage
import datetime
from DrissionPage.common import Actions
from .utils import retry, TaskStatusCode
from typing import Dict, Any, Tuple, Optional
from .user_data import UserData
import random
logger = logging.getLogger(__name__)


class LoginManager:
    """
    登录管理器，负责处理用户登录流程。
    """

    def __init__(self):
        pass

    @retry(max_retries=3, delay=0.5)
    def login(self, page: ChromiumPage, user: UserData):
        """
        执行登录操作，并直接更新 user 对象的状态。
        """
        try:
            page.clear_cache()
            logger.debug(f"用户 {user.username} - 浏览器缓存和cookies已清除。")
        except Exception as e:
            logger.warning(f"用户 {user.username} - 清除浏览器缓存和cookies时发生错误: {e}")

        if not user.username or not user.password:
            logger.error(f"用户 {user.name} 的账号或密码为空。")
            user.status_code = TaskStatusCode.LOGIN_FAILED_CREDENTIALS
            return

        logger.info(f"开始为用户 {user.name} {user.username} 执行登录...")
        page.get("https://www.2-class.com/")
        page.ele('tag:a@text():登录').click()

        user_input = page.ele('#account', timeout=10)
        pwd_input = page.ele('#password', timeout=10)
        user_input.input(user.username)
        pwd_input.input(user.password)

        login_button = page.ele('t:button@class:submit-btn', timeout=10)
        if login_button:
            Actions(page).move_to(login_button, duration=0.5).left(
                30).hold().wait(0.01, 0.15).release()
        else:
            logger.error("未能定位到登录按钮，请检查选择器。")
            user.status_code = TaskStatusCode.LOGIN_FAILED_OTHER
            return

        # 等待页面开始跳转/刷新
        time.sleep(0.5)  # 短暂等待登录操作生效

        # 等待页面加载完成，尝试多种方法检测页面状态
        page_refreshed = False
        try:
            # 尝试等待URL变化（页面跳转）
            page.wait.url_change(timeout=10)
            page_refreshed = True
        except Exception:
            # 如果URL没有变化，可能是页面内部刷新，继续等待内容变化
            logger.debug(f"用户 {user.username} - 登录后URL未发生明显变化，继续检测内容变化")

        # 等待一段时间让页面内容加载
        time.sleep(1)

        # 等待页面稳定，最多等待10秒
        stability_check_interval = 0.25
        stability_check_count = 0
        max_stability_checks = 40  # 10秒

        while stability_check_count < max_stability_checks:
            try:
                # 尝试访问页面内容以检查是否稳定
                current_html = page.html
                page_refreshed = True
                break  # 如果可以正常访问页面内容，则跳出循环
            except Exception as e:
                logger.debug(f"用户 {user.username} 页面仍在加载或刷新: {e}")
                time.sleep(stability_check_interval)
                stability_check_count += 1

        if not page_refreshed:
            logger.warning(f"用户 {user.username} 页面状态检测超时")

        for _ in range(40):
            try:
                # 检查页面是否有内容，防止页面刷新导致的访问错误
                if "我的课程" in page.html:
                    try:
                        uinfo = page.run_js("return window.__DATA__")
                        user.reqtoken = uinfo['reqtoken']
                        user.sid = page.cookies().as_dict().get('sid', '')
                        user.cookies = page.cookies().as_dict()
                        user.login_time = datetime.datetime.now()
                        user.status_code = TaskStatusCode.LOGIN_SUCCESS
                        logger.info(f"用户 {user.username} 登录成功")
                        return
                    except Exception as e:
                        logger.debug(f"用户 {user.username} 尝试获取登录信息时出错: {e}")
                        time.sleep(0.25)  # 等待JS加载
                elif "用户名或密码错误" in page.html:  # 检查是否有密码错误提示
                    logger.warning(f"用户 {user.username} 登录失败: 账号或密码错误")
                    user.status_code = TaskStatusCode.LOGIN_FAILED_CREDENTIALS
                    return
            except Exception as e:
                # 页面可能在刷新中，捕获异常并继续等待
                if "页面被刷新" in str(e):
                    logger.debug(f"用户 {user.username} 页面正在刷新: {e}")
                else:
                    logger.debug(f"用户 {user.username} 检查页面内容时遇到异常 (可能是页面刷新): {e}")
                time.sleep(0.25)
                continue

            try:
                loginVeriBox_div = page.ele(
                    't:div@class:loginVeriBox', timeout=0.1)
                if loginVeriBox_div:  # 检测到认证框
                    from DrissionPage.common import Settings

                    click_ver_div = page.ele(
                        '#nc_1_refresh2', timeout=0.1)  # 点击框
                    if click_ver_div:
                        logger.info(f"用户 {user.username} 检测到点击验证，尝试处理...")
                        click_ver_div_size = click_ver_div.rect.size  # 元素大小
                        logger.debug(
                            "用户 {user.username} 点击认证框 {click_ver_div_size[0]}x{click_ver_div_size[1]}")
                        page.actions.move_to(click_ver_div, offset_x=click_ver_div_size[0], offset_y=click_ver_div_size[1], duration=0.5).left(
                            random.randint(1, 5))
                        page.actions.hold(click_ver_div).wait(
                            0.01, 0.1).release()
                        time.sleep(0.5)
                        if "验证通过" not in page.html:
                            logger.warning(f"用户 {user.username} 滑块验证失败。")
                            user.status_code = TaskStatusCode.LOGIN_FAILED_SLIDER

                    slider_button = page.ele('#nc_1_n1z', timeout=0.1)  # 滑块按钮
                    if slider_button:
                        logger.info(f"用户 {user.username} 检测到滑块验证，尝试处理...")
                        slider_bg_map = page.ele(
                            '#nc_1__scale_text', timeout=0.1)  # 底图用于获取移动位置
                        if not slider_bg_map:
                            logger.debug("用户 {user.username} 主元素定位失败，启用备用定位1")
                            slider_bg_map = page.ele(
                                't:div@class:slidetounlock', timeout=0.1)  # 备用定位
                            if not slider_bg_map:
                                logger.debug(
                                    "用户 {user.username} 备用元素定位1 定位失败, 使用盲打模式 394x34")
                                slider_bg_map_size = [394, 34]
                            else:
                                slider_bg_map_size = slider_bg_map.rect.size
                        else:
                            slider_bg_map_size = slider_bg_map.rect.size  # 元素大小

                        slider_button_rect = slider_button.rect
                        start_point = (slider_button_rect.location[0] + random.randint(0, int(
                            slider_button_rect.size[0])), slider_button_rect.location[1] + random.randint(0, int(slider_button_rect.size[1])))
                        forward_times = random.randint(4, 8)
                        avg_x = int(slider_bg_map_size[0] / forward_times) # 平均每次移动x大致
                        move_point_list = []
                        moved_x = 0
                        moved_y = 0
                        msg = f"{start_point[0]}x{start_point[1]} -> "
                        for i in range(forward_times):
                            if i + 1 == forward_times:
                                move_x_this = slider_bg_map_size[0] - moved_x
                            else:
                                move_x_this = avg_x + random.randint(0, 10)-5
                            if moved_y > slider_button_rect.size[1] / 2:
                                move_y_this = slider_button_rect.size[1] / 2 + random.randint(
                                    0, 15)
                            else:
                                move_y_this = slider_button_rect.size[1] / 2 - random.randint(
                                    0, 15)
                            move_point_list.append((move_x_this, move_y_this))
                            moved_x += move_x_this
                            moved_y += move_x_this
                            msg += f" {move_x_this}x{move_y_this}"
                        logger.debug(f"用户 {user.username} 生成滑块轨迹成功 {msg}")
                        page.actions.move_to(start_point, duration=0.25)
                        page.actions.hold()
                        for i in range(forward_times):
                            page.actions.move(
                                move_point_list[i][0], move_point_list[i][1])
                        page.actions.release()
                        logger.debug(f"用户 {user.username} 滑块轨迹已经释放 {msg}")
                        time.sleep(0.15)
            except Exception as e:
                logger.debug(
                    f"用户 {user.username} 在处理滑块验证时发生异常。", exc_info=True)

            else:
                time.sleep(0.25)

        logger.warning(f"用户 {user.username} 登录超时或遇到未知问题")
        user.status_code = TaskStatusCode.LOGIN_FAILED_OTHER

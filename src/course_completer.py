# -*- coding: utf-8 -*-

import logging
import requests
import json
from .user_data import UserData
from .config_manager import ConfigManager

logger = logging.getLogger(__name__)


class CourseCompleter:
    """
    负责完成课程任务。
    """

    def __init__(self, config: ConfigManager):
        self.config = config

    def complete_courses(self, user: UserData) -> tuple[int, int]:
        """
        为指定用户完成课程。
        返回一个元组 (本次成功完成课程数, 总完成课程数)。
        """
        session = requests.Session()
        if user.cookies:
            for name, value in user.cookies.items():
                session.cookies.set(name, value)
        
        ua = self.config.get_config_value('network_settings.user_agent')
        session.headers.update({"User-Agent": ua})

        grade = user.grade
        if not grade:
            logger.error(f"用户 {user.username} 年级信息缺失，无法进行刷课。")
            return 0, 0

        # 1. 获取课程列表和已完成课程
        url = f"https://www.2-class.com/api/course/getHomepageCourseList?grade={grade}&pageSize=24&pageNo=1"
        finished_list = []
        try:
            response = session.get(url, timeout=10)
            response.raise_for_status()
            data = response.json().get('data', {})
            course_list = data.get('list', [])
            
            finished_list = [
                str(course['id']) for course in course_list
                if course.get('type') == 'course' and str(course.get('isFinish')) == '1'
            ]
            logger.info(f"用户 {user.username} 已完成 {len(finished_list)} 门课程。")

        except requests.RequestException as e:
            logger.error(f"为用户 {user.username} 获取课程列表失败: {e}")
            return 0, len(finished_list)

        # 2. 判断是否需要刷课
        min_finished = self.config.get_config_value('task_settings.min_finished_courses', 2)
        if len(finished_list) >= min_finished:
            logger.info(f"用户 {user.username} 已完成所有必需的课程，无需操作。")
            return 0, len(finished_list)

        # 3. 执行刷课
        courses_to_finish_map = self.config.get_config_value('task_settings.courses_to_finish', {})
        courses_for_grade = courses_to_finish_map.get(grade)
        if not courses_for_grade:
            logger.info(f"年级 {grade} 的待完成课程列表为空，将自动选择前两个未完成课程。")
            unfinished_courses = [
                str(course['id']) for course in course_list
                if str(course.get('isFinish')) != '1'
            ]
            courses_for_grade = unfinished_courses[:2]
        
        logger.debug(f"待完成课程列表 (年级: {grade}): {courses_for_grade}")
        
        completed_count = 0
        for course_id in courses_for_grade:
            if str(course_id) in finished_list:
                logger.debug(f"课程 {course_id} 已完成，跳过。")
                continue

            try:
                self._complete_single_course(session, user, course_id)
                completed_count += 1
                if len(finished_list) + completed_count >= min_finished:
                    break
            except Exception as e:
                logger.error(f"为用户 {user.username} 完成课程 {course_id} 时失败: {e}")
        
        total_completed = len(finished_list) + completed_count
        logger.info(f"用户 {user.username} 完成了 {completed_count}(本次) / {total_completed}(总共) 门课程。")
        return completed_count, total_completed

    def _complete_single_course(self, session: requests.Session, user: UserData, course_id: str):
        """
        完成单门课程的逻辑。
        """
        # 获取试卷
        test_paper_url = f"https://www.2-class.com/api/exam/getTestPaperList?courseId={course_id}"
        response = session.get(test_paper_url)
        response.raise_for_status()
        test_paper_list = response.json()['data']['testPaperList']

        logger.debug(f"获取到课程 {course_id} 的试卷列表: {test_paper_list}")
 
        # 构造答案
        exam_commit_data = {
            "courseId": str(course_id),
            "reqtoken": user.reqtoken,
            "exam": "course",
            "examCommitReqDataList": [
                {"examId": i + 1, "answer": paper['answer']}
                for i, paper in enumerate(test_paper_list)
            ]
        }

        # 提交答案
        commit_url = "https://www.2-class.com/api/exam/commit"
        headers = dict(session.headers)
        headers['content-type'] = "application/json;charset=UTF-8"
        response = session.post(commit_url, data=json.dumps(exam_commit_data), headers=headers)
        response.raise_for_status()

        if not response.json().get('data', False):
             raise Exception("提交考试后返回失败状态。")
        
        logger.debug(f"用户 {user.username} 课程 {course_id} 提交成功。")

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.common.exceptions import *
import requests
import pyexcel
import json
import time
import threading
import subprocess
import socket
import random
from _thread import start_new_thread as st

browsers = []
students = []
checkThreads = []
ResReady_Browers = False
ResReady_UserData = False


class Config:
    def __init__(self):
        with open('./config.json', encoding="utf8") as f:
            f = json.loads(f.read())

        self.browerPath = f['browersEnv']["browerPath"]
        self.webDriver = f['browersEnv']['webDriver']
        self.cachesPath = f['browersEnv']['cachesPath']
        self.proxy = f['proxy']['proxy']
        self.onlyProxyLogin = f['proxy']['onlyProxyLogin']
        self.maximumBrowers = f['app']['maximumBrowers']
        self.maximumThread = f['app']['maximumThread']
        self.studentList = f['app']['studentList']
        self.startDelay = f['browersEnv']['startDelay']
        self.needFinishNum = f['app']['needFinishCourses']['needFinishNum']
        self.willFinishListIfLess = f['app']['needFinishCourses']['willFinishListIfLess']
        self.ua = f['proxy']['ua']


Config = Config()


# 文本年级to标准年级, 规避各种问题  # 5 6 7 8 9 10(高一) 11(高二) 12(中职一) 13(中职二)
def gradeToNum(grade):
    try:
        if 5 <= int(grade) <= 13:
            grades = ["五年级", "六年级", "初一", "初二", "初三", "高一", "高二", "中职一", "中职二"]
            return grades[int(grade) - 5]
    except:
        pass
    grade = str(grade)
    if grade.find("中") != -1 or grade.find("职") != -1 or grade.find("业") != -1:
        if grade.find("一") != -1 or grade.find("1") != -1:
            return "中职一"
        if grade.find("二") != -1 or grade.find("2") != -1:
            return "中职二"
    if grade.find("高") != -1:
        if grade.find("一") != -1 or grade.find("1") != -1:
            return "高一"
        if grade.find("二") != -1 or grade.find("2") != -1:
            return "高二"
    if grade.find("初") != -1:
        if grade.find("一") != -1 or grade.find("1") != -1:
            return "初一"
        if grade.find("二") != -1 or grade.find("2") != -1:
            return "初二"
        if grade.find("三") != -1 or grade.find("3") != -1:
            return "初三"
    if grade.find("小") != -1:
        if grade.find("五") != -1 or grade.find("5") != -1:
            return "五年级"
        if grade.find("六") != -1 or grade.find("6") != -1:
            return "六年级"
    if grade.find("年") != -1 or grade.find("级") != -1:
        if grade.find("五") != -1 or grade.find("5") != -1:
            return "五年级"
        if grade.find("六") != -1 or grade.find("6") != -1:
            return "六年级"
        if grade.find("七") != -1 or grade.find("7") != -1:
            return "初一"
        if grade.find("八") != -1 or grade.find("8") != -1:
            return "初二"
        if grade.find("九") != -1 or grade.find("9") != -1:
            return "初三"
    return None


def getFreePort():
    while True:
        try:
            sock = socket.socket()
            sock.bind(('', 0))
            sock.close()
            ip, port = sock.getsockname()
            return port
        except OSError:
            return random.randint(10000, 65535)


class BrowerX(threading.Thread):
    def __init__(self, bid="AnonymousID"):
        super().__init__()
        self.xbid = bid
        self.bid = bid + "_" + str(random.randint(10000, 99999))
        self.port = None
        self.browser = None
        self.process = None

    def run(self):
        global browsers
        for _ in range(5):
            try:
                self.port = getFreePort()
                print("[INFO] 尝试启动浏览器 ID: " + str(self.bid) + " 运行于端口:", self.port)
                cmd = [Config.browerPath,
                       "--disk-cache-dir=" + Config.cachesPath + "/disk_" + str(self.xbid),
                       "--user-data-dir=" + Config.cachesPath + "/userdata_" + str(self.xbid),
                       "--remote-debugging-port=" + str(self.port),
                       "--window-position=0,0",
                       "--no-first-run"]
                self.process = subprocess.Popen(cmd)
                break
            except Exception as error:
                print("[ERROR] 启动浏览器出现错误, 启动进程出现错误, 请等待自动重试 :( L" + str(error.__traceback__.tb_lineno), ")", error)
                time.sleep(3)

        for _ in range(5):
            try:
                time.sleep(3)
                chrome_options = Options()
                chrome_service = Service(Config.webDriver)
                chrome_options.binary_location = Config.browerPath
                chrome_options.add_experimental_option("debuggerAddress", "127.0.0.1:" + str(self.port))
                self.browser = webdriver.Chrome(options=chrome_options, service=chrome_service)
                break
            except Exception as error:
                print("[ERROR] 启动浏览器出现错误, 连接浏览器出现错误, 可能是尚未完全启动, 请等待自动重试:( L" + str(error.__traceback__.tb_lineno), ")", error)
        print("[INFO] 启动浏览器完成 ID: " + str(self.bid))
        browsers.append(self)

    def stop(self):
        if self.process:
            self.process.terminate()
            print("[INFO] 停止浏览器 ID:", self.bid, "成功从端口:", self.port)

    def waitElementAppeard(self, by: str, value: str):  # 等待元素出现
        while True:
            try:
                svv = self.browser.find_element(by=by, value=value)

                return svv
            except NoSuchElementException:
                time.sleep(0.05)

    def jsClick(self, locat):
        try:
            self.browser.execute_script("arguments[0].click();", locat)
            return True
        finally:
            return False

    def login(self, user: str, pwd: str):
        print("[INFO] 开始登录:", user, "使用浏览器:", self.bid)
        self.browser.execute_cdp_cmd('Network.clearBrowserCookies', {})
        time.sleep(0.05)
        self.browser.get("https://www.2-class.com/")
        ess1 = self.waitElementAppeard(By.XPATH, value="""//*[@id="app"]/div/div[1]/div/header/div/div[2]/div[2]/div/span/a""")
        self.jsClick(ess1)
        ess1 = self.waitElementAppeard(by=By.XPATH, value="""//*[@id="account"]""")  # 用户框
        ess2 = self.waitElementAppeard(by=By.XPATH, value="""//*[@id="password"]""")  # 密码框
        ess1.clear()
        ess2.clear()
        ess1.send_keys(user)
        ess2.send_keys(pwd)
        ess1 = self.waitElementAppeard(By.XPATH, value="""/html/body/div[2]/div/div[2]/div/div[1]/div/form/div/div/div/button""")  # 登
        self.jsClick(ess1)
        print("[INFO] 进入查询环节登录成功:", user, "使用浏览器:", self.bid)
        nk = False
        while True:  # TODO: 支持错误重登
            try:
                uinfo = self.browser.execute_script("return window.__DATA__")  # 此处可以获取基础信息
                # print("[DEBUG] LOGIN1", user, pwd, self.bid, json.dumps(uinfo))
                reqtoken = uinfo['reqtoken']
                userInformation = str(uinfo['userInfo']['id']) + "|" + uinfo['userInfo']['department']['schoolName'] + "(" + str(uinfo['userInfo']['department']['schoolId']) + ")|" + uinfo['userInfo']['department'][
                    'gradeName'] + "(" + str(uinfo['userInfo']['department']['gradeId']) + ")|" + uinfo['userInfo']['department']['className']
                sid = self.browser.get_cookie("sid")['value']
                cookies = self.browser.get_cookies()
                print("[INFO] 登录成功:", user, "使用浏览器:", self.bid, "ReqToken:", reqtoken, "sid:", sid, "用户信息:" + userInformation)
                return 0, reqtoken, sid, cookies
            except:
                pass
            # except Exception as error:
            # print("[DEBUG] GET ERROR BS2 BID:" + self.bid + " ERROR(LN" + str(error.__traceback__.tb_lineno) + "):\n" + str(error) + "\n" + json.dumps(uinfo))
            try:
                ele = self.browser.find_element(By.XPATH, """/html/body/div[2]/div/div[2]/div/div[1]/div/form/div/div/div/div[1]/span[1]""")
                print("[DEBUG]", ele.text, "BID:", self.bid)
                if ele.text.find("用户名或密码错误") != -1:
                    print("[ERROR] 用户名或密码错误, 用户:", user, "BID:", self.bid)
                    return 1, "", ""
            except:
                pass
            try:
                self.browser.find_element(By.XPATH, """/html/body/div[2]/div/div[2]/div/div[1]/div""")  # UI
                nk = True
            except:
                if nk is False:
                    time.sleep(0.5)
                    continue
            #   else: # 登录成功
            #       break
            try:
                ele = self.browser.find_element(By.XPATH, """/html/body/div[2]/div/div[2]/div/div[1]/div/form/div/div/div/button""")  # 滑块
                if ele.is_enabled():
                    try:
                        ele = self.browser.find_element(By.XPATH, """//*[@id="`nc_1_refresh2`"]""")  # 失败框
                        print("[DEBUG] 发现点击认证 BID:", self.bid)
                        ele.click()
                        time.sleep(0.4)
                    except:
                        pass
                    try:
                        ele = self.browser.find_element(By.XPATH, """//*[@id="nc_1_n1z"]""")  # 滑块按钮
                        print("[DEBUG] 发现滑块认证 BID:", self.bid)
                        ActionChains(self.browser).drag_and_drop_by_offset(ele, 759 - 405, 0).perform()
                        time.sleep(0.5)
                    except:
                        pass
                else:
                    time.sleep(0.2)
            except:
                time.sleep(0.1)


class UserX(threading.Thread):  # 用户多线程执行函数
    def __init__(self, bid="AnonymousID"):
        super().__init__()
        self.bid = bid + "_" + str(random.randint(10000, 99999))

    def run(self):
        # print("[DEBUG] 执行线程已启动, 进入竞争流程")
        global students
        global ResReady_UserData
        while not ResReady_UserData:
            time.sleep(1)
        while True:
            if len(students) > 0:  # 线程竞争用户
                try:
                    student = students.pop(0)
                except IndexError:
                    time.sleep(0.1)
                    continue
            else:
                time.sleep(0.1)
                continue

            grade = gradeToNum(student[3])
            username = student[1]
            password = student[2]
            # print("[DEBUG] 竞得用户成功," + str(grade) + "|" + str(username) + "|" + str(password) + ", BID:" + str(self.bid))
            if username is None or password is None or username == "" or password == "" or username == "账号" or username == "账户" or username == "帐号" or password == "密码" or grade == "年级":
                # print("[ERROR] 用户名密码不能为空 BID:", self.bid)
                continue
            if grade is not None:
                user = doUser(name=student[0], username=student[1], password=password, grade=grade, bid=self.bid)
                user.login()
                user.doMain()
            else:
                print("[ERROR] 年级填写有误 BID:", self.bid, "用户:", username)  # TODO: 使用登录时获取到的默认年级


class doUser:  # 主功能函数
    def __init__(self, username: str, password: str, grade: str, bid: str, name="Anonymous"):
        self.sid = None
        self.reqtoken = None
        self.name = str(name)
        self.user = str(username)
        self.password = str(password)
        self.grade = grade
        self.bid = bid
        self.cookies = None

    def login(self):
        while True:
            global browsers
            if len(browsers) > 0:  # 线程竞争浏览器
                try:
                    browser = browsers.pop(0)
                except IndexError:
                    # print("[DEBUG] R8X 无法竞得浏览器 BID:" + str(self.bid))
                    time.sleep(0.01)
                    continue
            else:
                # print("[DEBUG] R9X 无法竞得浏览器 BID:" + str(self.bid))
                time.sleep(0.1)
                continue

            # print("[DEBUG] D6X 竞得浏览器 BID:" + str(self.bid))
            _, reqtoken, sid, cookies = browser.login(user=self.user, pwd=self.password)
            self.sid = sid
            self.reqtoken = reqtoken
            self.cookies = cookies
            browsers.append(browser)
            return _

    def doMain(self):
        url = "https://www.2-class.com/api/course/getHomepageCourseList?grade=" + self.grade + "&pageSize=24&pageNo=1"
        session = requests.session()
        headers = {
            "User-Agent": Config.ua
        }
        for cookie in self.cookies:  # fxxk
            session.cookies.set(cookie['name'], cookie['value'])
        session.headers = headers
        rs = session.get(url=url)
        # print("[DEBUG]",rs.text)
        r = json.loads(rs.text)['data']
        finished_list = []
        for i in range(len(r['list'])):
            if r['list'][i]["type"] == "course":
                if not r['list'][i]["isFinish"] is None:
                    if str(r['list'][i]["isFinish"]) == "1":
                        finished_list.append(str(r['list'][i]["id"]))

        # print("[DEBUG] " + name + " 完成课程数: " + str(finishcount))
        if len(finished_list) < Config.needFinishNum:
            print("[INFO] " + self.name, self.user, "完成数不足, 开始刷课 BID:", self.bid)  # TODO: 增加二次校验
            for examR in Config.willFinishListIfLess[self.grade]:
                examR = str(examR)
                if str(examR) in finished_list:  # 已经完
                    print("[DEBUG] RSD0 已完成 跳过 " + self.name + " " + self.user + " " + self.bid + " ID:" + str(examR))
                    continue
                qheader = session.headers.copy()
                qheader['content-type'] = "application/json;charset=UTF-8"
                # rs = session.post(url="https://2-class.com/api/course/addCoursePlayPV", data={"courseId": int(examR), "reqtoken": self.reqtoken}, headers=qheader)
                # print("[DEBUG] RSD1 提交观看 " + self.name + " " + self.user + " " + self.bid + " " + rs.text)
                rs = session.get(url="https://www.2-class.com/api/exam/getTestPaperList?courseId=" + str(examR))
                print("[DEBUG] RSD2 获取内容 " + self.name + " " + self.user + " " + self.bid + " " + rs.text)
                testPaperList = json.loads(rs.text)['data']['testPaperList']
                data = {"courseId": str(examR), "reqtoken": self.reqtoken, "exam": "course"}
                examCommitReqDataList = []
                for i in range(len(testPaperList)):
                    examCommitReqDataList.append({
                        "examId": i + 1,
                        "answer": testPaperList[i]['answer']
                    })
                data['examCommitReqDataList'] = examCommitReqDataList
                print(json.dumps(data))
                print(session.cookies.items())
                print(session.headers.items())
                rs = session.post(url="https://www.2-class.com/api/exam/commit", data=json.dumps(data), headers=qheader)
                if json.loads(rs.text)['data']:
                    ypass = "是"
                else:
                    ypass = "否"
                print("[DEBUG] RSD3 提交考试 " + self.name + " " + self.user + " " + self.bid + " 通过:" + ypass + " " + rs.text)
                finished_list.append(examR)
            print("[INFO] " + self.name + " " + self.user + " 已完成课程, R1, BID: " + self.bid)

        else:
            print("[INFO] " + self.name + " " + self.user + " 已完成课程, 完成数: " + str(len(finished_list)) + " BID: " + str(self.bid))


def getMTime():
    return int(time.time() * 1000)


def readSheet():
    global students
    global ResReady_UserData
    print("[INFO] 执行读取账号数据")
    ctime_1 = time.time()
    book = pyexcel.get_book(file_name=Config.studentList)
    for sheet in book:  # 遍历Book中的所有Sheet
        students = students + sheet.get_array()
    students = [tuple(student) for student in students]  # 先转元组然后去重
    # students = list(set(students)) # 使用set会改变排序
    students = list({}.fromkeys(students).keys())
    ctime_2 = time.time()
    ResReady_UserData = True
    print("[INFO] 读取数据完成, 共 " + str(len(book)) + " 个表 " + str(len(students)) + " 个账号, 共耗时 " + str(1000 * (ctime_2 - ctime_1)) + " ms")


def startBrowser():
    global ResReady_Browers
    ctime_1 = time.time()
    print("[INFO] 执行浏览器启动流程, 共需要启动: " + str(Config.maximumBrowers) + " 个浏览器")
    for i in range(Config.maximumBrowers):  # 启动浏览器
        browser = BrowerX("BRO_" + str(i))
        # print("[DEBUG] TEST 已启动 X N " + str(i))
        browser.start()
        time.sleep(Config.startDelay / 1000)
    ctime_2 = time.time()
    while len(browsers) < Config.maximumBrowers:
        time.sleep(0.02)
    ResReady_Browers = True
    print("[INFO] 启动浏览器完成, 共 " + str(Config.maximumBrowers) + " 个浏览器, 共耗时 " + str(1000 * (ctime_2 - ctime_1)) + " ms")


def startUser():
    print("[INFO] 执行线程启动流程, 共需要启动: " + str(Config.maximumThread) + " 线程")
    for i in range(Config.maximumThread):
        # print("[DEBUG] 启动执行线程A1, " + str(i))
        user = UserX("USER_" + str(i))
        user.start()
    print("[INFO] 执行线程启动流程完成, 共 " + str(Config.maximumThread) + " 个线程")


def init():
    global students
    global ResReady_UserData
    global ResReady_Browers
    print("[INFO] 初始化开始")
    st(readSheet, ())  # 读数据
    time.sleep(0.001)
    st(startBrowser, ())  # 启动浏览器
    time.sleep(0.001)
    st(startUser, ())  # 启动调用线程
    while not ResReady_UserData or not ResReady_Browers:  # 等待启动
        time.sleep(0.1)
    print("[INFO] 初始化完成")


def exitBrowser():
    global browsers
    h = len(browsers)
    print("[INFO] 执行浏览器退出流程, 共需要退出:", h, "个")
    for i in range(h):
        brower = browsers.pop(0)
        brower.stop()


if __name__ == "__main__":
    init()
    # time.sleep(10)
    while True:
        a = input()
        if a == "e":
            exitBrowser()
            time.sleep(10)
            exit()
        if a == "p":
            print(browsers, students)

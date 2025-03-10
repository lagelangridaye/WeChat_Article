# -*- coding: utf-8 -*-
import WeChat
from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
import sys
import os
import re
from time import sleep, localtime, time, strftime
import undetected_chromedriver as uc
from PyQt5.QtWidgets import QApplication
from bs4 import BeautifulSoup
import requests
import json
import urllib.parse
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from math import ceil
import threading
import inspect
import ctypes
import random
from goto import with_goto
import configparser
import pyautogui
# import pdfkit

'''
conf.ini
    [resume]
    rootpath = ''
    pagenum = 0
    linkbuf_cnt = 0
    download_cnt = 0
'''
# 设置 递归调用深度 为 一百万
sys.setrecursionlimit(1000000)

# https://github.com/wnma3mz/wechat_articles_spider/blob/master/docs/使用的微信公众号接口.md
# title_buf = []
# link_buf = []
pro_continue = 0
class MyMainWindow(WeChat.Ui_MainWindow):
    def __init__(self):
        self.sess = requests.Session()
        self.headers = {
            'Host': 'mp.weixin.qq.com',
            'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:65.0) Gecko/20100101 Firefox/65.0',
        }
        self.browser_path = r'Chrome/BitBrowser.exe'
        self.driver_path = r'Chrome/chromedriver.exe'

        self.initpath = os.getcwd()  # 获取当前工作目录（Current Working Directory）的路径
        self.rootpath = os.getcwd() + r"/spider/"  # 全局变量，存放路径
        self.time_gap = 5  # 全局变量，每页爬取等待时间
        self.timeStart = 1999  # 全局变量，起始时间

        self.year_now = localtime(time()).tm_year  # 当前年份，用于比对时间
        self.timeEnd = self.year_now + 1  # 全局变量，结束时间
        self.thread_list = []  # 存储线程对象
        self.label_debug_string = ""  # 存储调试信息
        self.label_debug_cnt = 0  # 记录调试信息的行数
        self.total_articles = 0  # 记录当前已抓取的文章数量
        self.keyWord = ""  # 存储用户输入的关键词
        self.keyword_search_mode = 0  # 标识当前是否处于关键词搜索模式。0 表示普通模式，1 表示关键词搜索模式。
        self.keyWord_2 = ""  # 存储用户输入的第二个关键词，会在程序运行过程中用于进一步筛选文章。
        self.freq_control = 0  # 用于控制请求频率。0 表示正常频率，1 表示需要降低频率以避免触发反爬机制。
        self.download_cnt = 0  # 记录已下载的文章数量
        self.linkbuf_cnt = 0  # 记录已抓取的文章链接数量
        self.download_end = 0  # 标识下载任务是否完成 0:未完成 1:已完成
        # 检查和初始化程序的配置文件conf.ini，
        # 返回值：如果配置文件存在且读取成功，返回 1，表示可以继续之前的爬取任务；
        # 如果配置文件不存在且已创建，返回 0，表示这是一个新的爬取任务
        # isresume = 0不需要断点续爬，需要从零开始
        # isresume = 1需要断点续爬，需要从上次爬取的位置继续
        self.isresume = self.Check_Config()
        self.url_json_init()  # 初始化和管理 url.json 文件
        self.title_buf = []  # 存储标题
        self.link_buf = []  # 存储链接
        self.wechat_uin = None  # 微信公众号的 UIN
        self.wechat_key = None  # 微信公众号的 Key

    def vari_init(self):
        # 初始化路径相关变量
        self.rootpath = os.getcwd() + r"/spider/"  # 设置爬取结果的存储路径

        # 初始化线程相关变量
        self.thread_list = []  # 清空线程列表

        # 初始化调试信息相关变量
        self.label_debug_string = ""  # 清空调试信息字符串
        self.label_debug_cnt = 0  # 重置调试信息计数器

        # 初始化文章相关变量
        self.total_articles = 0  # 重置已抓取文章总数
        self.keyWord = ""  # 清空主关键词
        self.keyword_search_mode = 0  # 重置关键词搜索模式标志
        self.keyWord_2 = ""  # 清空备用关键词

        # 初始化状态标志
        self.freq_control = 0  # 重置频率控制标志
        self.download_cnt = 0  # 重置已下载文章计数
        self.linkbuf_cnt = 0  # 重置已抓取链接计数
        self.download_end = 0  # 重置下载完成标志

        # 清空缓存
        self.title_buf.clear()  # 清空标题缓存
        self.link_buf.clear()  # 清空链接缓存

        # 初始化进度条
        self.progressBar.setMaximum(100)  # 设置进度条最大值为100
        self.progressBar.setValue(0)  # 重置进度条值为0

    def Label_Debug(self, string):
        if self.label_debug_cnt == 12:  # 当调试信息计数器达到12时，清空调试信息
            self.label_debug_string = ""
            self.label_notes.setText(self.label_debug_string)
            self.label_debug_cnt = 0
        self.label_debug_string += "\r\n" + string  # 将新的调试信息追加到调试字符串中
        self.label_notes.setText(self.label_debug_string)  # 更新UI中的调试信息显示
        self.label_debug_cnt += 1  # 调试信息计数器加1

    def Label_Debug_Clear(self):
        self.label_debug_string = ""
        self.label_notes.setText(self.label_debug_string)
        self.label_notes.clear()
        self.label_debug_cnt = 0

    def setupUi(self, MainWindow):
        super(MyMainWindow, self).setupUi(MainWindow)
        try:
            if os.path.exists(os.getcwd()+r'/login.json'):
                with open(os.getcwd()+r'/login.json', 'r', encoding='utf-8') as p:
                    login_dict = json.load(p)
                    print("登陆文件读取成功")
                    self.Label_Debug("登陆文件读取成功")
                    self.LineEdit_target.setText(login_dict['target'])  # 公众号的英文名称
                    self.LineEdit_user.setText(login_dict['user'])  # 自己公众号的账号
                    self.LineEdit_pwd.setText(login_dict['pwd'])  # 自己公众号的密码
                    self.LineEdit_timegap.setText(str(login_dict['timegap']))  # 每页爬取等待时间"
                    self.lineEdit_timeEnd.setText(str(self.year_now+1))  # 结束时间为当前年
                    self.lineEdit_timeStart.setText("1999")  # 开始时间为1999
                    QApplication.processEvents()  # 刷新文本操作
            
            image_url = "http://xfxuezhang.cn/web/share/donate/yf.png"
            response = requests.get(image_url)
            if response.status_code == 200:
                self.label_yf.setAlignment(Qt.AlignCenter)
                pixmap = QPixmap()
                pixmap.loadFromData(response.content)
                # 缩放图片以适应标签的大小
                scaled_pixmap = pixmap.scaled(self.label_yf.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
                self.label_yf.setPixmap(scaled_pixmap)
                print('image download ok')
            else:
                self.label_yf.setText("image url not found.")
        

        except Exception as e:
            print(e)

    def Start_Run(self):
        self.total_articles = 0
        Process_thread = threading.Thread(target=self.Process, daemon=True)
        Process_thread.start()
        self.thread_list.append(Process_thread)

    def Stop_Run(self):
        try:
            self.stop_thread(self.thread_list.pop())
            self.stop_thread(self.thread_list.pop())
            self.vari_init()  # 变量复位
            self.Label_Debug("终止成功!")
            print("终止成功!")
        except Exception as e:
            self.Label_Debug("终止失败!")
            print(e)

    def Start_Run_2(self):
        try:
            os.makedirs(self.rootpath)
        except:
            pass
        self.keyword_search_mode = 1
        self.total_articles = 0
        Process_thread = threading.Thread(target=self.Process, daemon=True)
        Process_thread.start()
        self.thread_list.append(Process_thread)

    def Stop_Run_2(self):
        try:
            self.keyword_search_mode = 0
            self.stop_thread(self.thread_list.pop())
            self.stop_thread(self.thread_list.pop())
            self.vari_init()  # 变量复位
            self.Label_Debug("终止成功!")
            print("终止成功!")
        except Exception as e:
            self.Label_Debug("终止失败!")
            print(e)

    def Change_IP(self):
        tar_url = r'https://www.douban.com'  # 目标测试URL
        http_s = '111.26.9.26:80'  # 代理IP地址和端口

        # 根据目标URL的协议类型设置代理
        if (tar_url.split(':')[0] == 'https'):
            proxies = {'https': http_s}
        else:
            proxies = {'http': http_s}

        try:
            # 使用代理访问目标URL
            html = self.sess.get(tar_url, proxies=proxies, timeout=(30, 60))
            print("* 代理有效√ *")
            print(html)
        except Exception as e:
            print("* 代理无效× *")
            print(e)
        pass

    def Check_Config(self):
        # 配置解析器的初始化操作，用于解析配置文件（通常是 .ini 格式的文件）
        self.conf = configparser.ConfigParser()
        self.cfgpath = os.path.join(os.getcwd(), "conf.ini")
        if os.path.exists(self.cfgpath):
            print("[Yes] conf.ini")
            try:
                self.conf.read(self.cfgpath, encoding="utf8")  # 读ini文件
            except:
                self.conf.read(self.cfgpath)  # 读ini文件
            resume = self.conf.items('resume')
            self.rootpath       = resume[0][1]
            self.pagenum        = int(resume[1][1])
            self.linkbuf_cnt    = int(resume[2][1])
            self.download_cnt   = int(resume[3][1])
            self.total_articles = int(resume[4][1])
            print(self.rootpath, self.pagenum, self.linkbuf_cnt, self.download_cnt, self.total_articles)
            return 1
        else:
            print("[NO] conf.ini")
            # 以写入模式（'w'）打开配置文件路径 self.cfgpath，并指定使用 UTF-8 编码。如果文件不存在，则会创建一个新文件。
            f = open(self.cfgpath, 'w', encoding="utf-8")
            f.close()
            self.conf.add_section("resume")
            self.conf.set("resume", "rootpath", os.getcwd())
            self.conf.set("resume", "pagenum", "0")
            self.conf.set("resume", "linkbuf_cnt", "0")
            self.conf.set("resume", "download_cnt", "0")
            self.conf.set("resume", "total_articles", "0")
            # 将配置数据写入到配置文件中
            self.conf.write(open(self.cfgpath, "w"))
            return 0

    def Process(self):
        try:
            username = self.LineEdit_user.text()  # 自己公众号的账号
            pwd = self.LineEdit_pwd.text()  # 自己公众号的密码
            query_name = self.LineEdit_target.text()  # 公众号的英文名称
            self.time_gap = self.LineEdit_timegap.text() or 10  # 每页爬取等待时间
            self.time_gap = int(self.time_gap)
            self.timeStart = self.lineEdit_timeStart.text() or 1999  # 起始时间
            self.timeStart = int(self.timeStart)  # 将起始时间转换为整数
            self.timeEnd = self.lineEdit_timeEnd.text() or self.year_now + 1  # 获取结束时间，若为空则默认为当前年份+1
            self.timeEnd = int(self.timeEnd)  # 将结束时间转换为整数
            self.keyWord = self.lineEdit_keyword.text()  # 获取关键词                                          # 关键词
            # uin_key = self.LineEdit_wechat.text().strip()                                           # 微信 uin,key
            # if uin_key:
            #     self.wechat_uin = re.search(r'uin=(.*?)&', uin_key).group(1)
            #     self.wechat_key = re.search(r'key=(.*?)&', uin_key).group(1)

            if self.checkBox.isChecked() is True and pwd != "":  # 如果勾选了记住密码且密码不为空
                dicts = {'target': query_name, 'user': username, 'pwd': pwd, 'timegap': self.time_gap}  # 创建包含登录信息的字典
                with open(os.getcwd() + r'/login.json', 'w+') as p:  # 打开login.json文件
                    json.dump(dicts, p)  # 将登录信息写入文件
                    p.close()  # 关闭文件

            [token, cookies] = self.Login(username, pwd)  # 调用Login方法获取token和cookies
            self.Add_Cookies(cookies)  # 将cookies添加到session中
            if self.keyword_search_mode == 1:  # 如果处于关键词搜索模式
                self.keyWord_2 = self.lineEdit_keyword_2.text()  # 获取第二个关键词
                self.KeyWord_Search(token, self.keyWord_2)  # 调用关键词搜索方法
            else:  # 如果不是关键词搜索模式
                [fakeid, nickname] = self.Get_WeChat_Subscription(token, query_name)  # 获取公众号的fakeid和昵称
                if self.isresume == 0:  # 如果不是恢复模式
                    Index_Cnt = 0  # 初始化索引计数器
                    while True:  # 循环直到找到可用的目录
                        try:
                            # 构建保存路径，格式为：当前目录/spider-序号/公众号昵称
                            self.rootpath = os.path.join(os.getcwd(), "spider-%d" % Index_Cnt, nickname)
                            os.makedirs(self.rootpath)  # 创建目录
                            # 更新配置文件中的保存路径
                            self.conf.set("resume", "rootpath", self.rootpath)
                            # 将更新后的配置写入文件
                            self.conf.write(open(self.cfgpath, "r+", encoding="utf-8"))
                            break  # 成功创建目录后退出循环
                        except:  # 如果目录已存在
                            Index_Cnt = Index_Cnt + 1  # 增加序号，尝试下一个目录
                self.Get_Articles(token, fakeid)  # 调用获取文章的方法
        except Exception as e:
            self.Label_Debug("!!![%s]" % str(e))  # 在调试窗口显示错误信息
            print("!!![%s]" % str(e))
            if "list" in str(e):  # 如果错误信息包含"list"
                self.Label_Debug("请删除cookie.json")  # 提示用户删除cookie文件
                print("请删除cookie.json")

    def url_json_write(self, inputdict):
        with open(self.url_json_path, "w+") as f:
            f.write(json.dumps(inputdict))

    def url_json_read(self):
        with open(self.url_json_path, "r+") as f:
            json_read = json.loads(f.read())
        return json_read

    def url_json_update(self, source, adddict):
        source.append(adddict)

    def url_json_init(self):
        self.url_json_path = os.path.join(os.getcwd(), "url.json")
        if os.path.exists(self.url_json_path):
            print("[Yes] url.json")
            # 检查是否需要从零开始（isresume为0表示不需要断点续爬）。
            if self.isresume == 0:
                # 如果不需要断点续爬，删除现有的url.json文件。
                os.remove(self.url_json_path)
                # 创建一个新的空url.json文件。
                self.url_json_write([])
        else:
            print("[NO] url.json")
            # 如果url.json文件不存在,创建一个新的空 url.json 文件
            self.url_json_write([])

        # 读取url.json文件的内容
        self.json_read = self.url_json_read()
        # 获取url.json文件中记录的数量
        self.json_read_len = len(self.json_read)
        print("len(url.json):", self.json_read_len)

    def url_json_once(self, dict_add):
        self.url_json_update(self.json_read, dict_add)  # {"Title": 1, "Link": 2, "Img": 3}
        self.url_json_write(self.json_read)
        self.json_read = self.url_json_read()
        # print("url_json_once OK")
        # print(self.json_read)

    def Login(self, username, pwd):
        try:  # 开始尝试执行代码
            if self.freq_control == 1:  # 检查频率控制标志是否为1
                raise RuntimeError('freq_control=1')  # 如果是，抛出运行时错误
            print(self.initpath+"/cookie.json")  # 打印cookie文件路径
            with open(self.initpath+"/cookie.json", 'r+') as fp:  # 以读写模式打开cookie文件
                cookieToken_dict = json.load(fp)  # 从文件中加载JSON数据
                cookies = cookieToken_dict[0]['COOKIES']  # 获取cookies
                token = cookieToken_dict[0]['TOKEN']  # 获取token
                print(token)  # 打印token
                print(cookies)  # 打印cookies

                if cookies != "" and token != "":  # 检查cookies和token是否为空
                    self.Label_Debug("cookie.json读取成功")  # 在调试窗口显示成功信息
                    print("cookie.json读取成功")  # 在控制台打印成功信息
                self.Add_Cookies(cookies)  # 将cookies添加到session中

                html = self.sess.get(r'https://mp.weixin.qq.com/cgi-bin/home?t=home/index&lang=zh_CN&token=%s' % token, timeout=(30, 60))  # 使用token访问微信公众号主页
                if "登陆" not in html.text:  # 检查页面中是否包含"登陆"字样
                    self.Label_Debug("cookie有效,无需浏览器登陆")  # 在调试窗口显示信息
                    print("cookie有效,无需浏览器登陆")  # 在控制台打印信息
                    return token, cookies  # 返回token和cookies
        except Exception as e:  # 捕获异常
            print("无cookie.json或失效 -", e)  # 在控制台打印错误信息
            self.Label_Debug("无cookie.json或失效")  # 在调试窗口显示错误信息


        self.Label_Debug("正在打开浏览器,请稍等")  # 在调试窗口显示提示信息
        print("正在打开浏览器,请稍等")  # 在控制台打印提示信息
        options = Options()  # 创建Chrome浏览器选项对象
        # options.add_argument("--headless")  # 启用无头模式（已注释）
        options.add_argument("--incognito")  # 启用隐身模式
        options.add_argument("--disable-blink-features")  # 禁用Blink特性
        options.add_argument("--disable-blink-features=AutomationControlled")  # 禁用自动化控制特性
        options.add_argument("--no-default-browser-check")  # 禁用默认浏览器检查
        options.add_argument("--allow-running-insecure-content")  # 允许运行不安全内容
        options.add_argument("--ignore-certificate-errors")  # 忽略证书错误
        options.add_argument("--disable-single-click-autofill")  # 禁用单次点击自动填充
        options.add_argument("--disable-autofill-keyboard-accessory-view[8]")  # 禁用自动填充键盘辅助视图
        options.add_argument("--disable-full-form-autofill-ios")  # 禁用iOS全表单自动填充
        browser = webdriver.Chrome(options=options, service=Service(ChromeDriverManager().install()))  # 启动Chrome浏览器
        # browser = uc.Chrome(driver_executable_path=self.driver_path,  # 使用undetected_chromedriver启动浏览器（已注释）
        #                    browser_executable_path=self.browser_path,
        #                    suppress_welcome=False)

        browser.maximize_window()  # 最大化浏览器窗口

        browser.get(r'https://mp.weixin.qq.com')  # 访问微信公众号登录页面
        browser.implicitly_wait(60)  # 设置隐式等待时间为60秒
        # account = browser.find_element(by=By.NAME, value="account")  # 查找账号输入框（已注释）
        # password = browser.find_element(by=By.NAME, value="password")  # 查找密码输入框（已注释）
        # if (username != "" and pwd != ""):  # 如果提供了账号密码（已注释）
        #     account.click()  # 点击账号输入框（已注释）
        #     account.send_keys(username)  # 输入账号（已注释）
        #     password.click()  # 点击密码输入框（已注释）
        #     password.send_keys(pwd)  # 输入密码（已注释）
        #     browser.find_element(by=By.XPATH, value=r'//*[@id="header"]/div[2]/div/div/form/div[4]/a').click()  # 点击登录按钮（已注释）
        # else:  # 如果未提供账号密码（已注释）
        #     self.Label_Debug("* 请在10分钟内手动完成登录 *")  # 提示手动登录（已注释）
        pyautogui.alert(title='请手动完成登录', text='完成登录后，点击确认!', button='确认')  # 弹出提示框，要求用户手动登录
        WebDriverWait(browser, 60 * 10, 0.5).until(  # 显式等待，最多等待10分钟
            EC.presence_of_element_located((By.CSS_SELECTOR, r'.weui-desktop-account__info'))  # 等待用户信息元素出现
        )
        self.Label_Debug("登陆成功")  # 在调试窗口显示登录成功信息
        token = re.search(r'token=(.*)', browser.current_url).group(1)  # 从当前URL中提取token
        cookies = browser.get_cookies()  # 获取当前页面的cookies
        with open(os.getcwd()+"/cookie.json", 'w+') as fp:  # 打开cookie.json文件
            temp_list = {}  # 创建临时字典
            temp_array = []  # 创建临时列表
            temp_list['COOKIES'] = cookies  # 将cookies存入字典
            temp_list['TOKEN'] = token  # 将token存入字典
            temp_array.append(temp_list)  # 将字典添加到列表
            json.dump(temp_array, fp)  # 将列表写入文件
            fp.close()  # 关闭文件
            self.Label_Debug(">> 本地保存cookie和token")  # 在调试窗口显示保存成功信息
            print(">> 本地保存cookie和token")  # 在控制台打印保存成功信息
        browser.close()  # 关闭浏览器
        return token, cookies  # 返回token和cookies

    def Add_Cookies(self, cookie):
        c = requests.cookies.RequestsCookieJar()  # 创建一个空的CookieJar对象
        for i in cookie:  # 遍历传入的cookie列表
            c.set(i["name"], i["value"])  # 将每个cookie的name和value添加到CookieJar中
            self.sess.cookies.update(c)  # 更新会话中的cookies

    def KeyWord_Search(self, token, keyword):
        self.url_buf = []
        self.title_buf = []
        header = {
            'Content - Type': r'application/x-www-form-urlencoded;charset=UTF-8',
            'Host': 'mp.weixin.qq.com',
            'User-Agent': r'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/74.0.3729.131 Safari/537.36',
            'Referer': 'https://mp.weixin.qq.com/cgi-bin/appmsg?t=media/appmsg_edit&action=edit&type=10&isMul=1&isNew=1&share=1&lang=zh_CN&token=%d' % int(token)
        }
        url = r'https://mp.weixin.qq.com/cgi-bin/operate_appmsg?sub=check_appmsg_copyright_stat'
        data = {'token': token, 'lang': 'zh_CN', 'f': 'json', 'ajax': 1, 'random': random.uniform(0, 1), 'url': keyword, 'allow_reprint': 0, 'begin': 0, 'count': 10}
        html_json = self.sess.post(url, data=data, headers=header).json()
        total = html_json['total']
        total_page = ceil(total / 10)
        print(total_page, '-', total)
        table_index = 0
        for i in range(total_page):
            data = {
                'token': token,
                'lang': 'zh_CN',
                'f': 'json',
                'ajax': 1,
                'random': random.uniform(0, 1),
                'url': keyword,
                'allow_reprint': 0,
                'begin': i*10,
                'count': 10
            }
            html_json = self.sess.post(url, data=data, headers=header).json()
            page_len = len(html_json['list'])
            # print(page_len)
            for j in range(page_len):
                self.url_buf.append(html_json['list'][j]['url'])
                self.title_buf.append(html_json['list'][j]['title'])
                print(j+1, ' - ', html_json['list'][j]['title'])
                table_count = self.tableWidget_result.rowCount()
                if (table_index >= table_count):
                    self.tableWidget_result.insertRow(table_count)
                self.tableWidget_result.setItem(table_index, 0, QtWidgets.QTableWidgetItem(self.title_buf[j]))  # i*20+j
                self.tableWidget_result.setItem(table_index, 1, QtWidgets.QTableWidgetItem(self.url_buf[j]))  # i*20+j
                table_index = table_index + 1
                self.total_articles += 1
                with open(self.rootpath + "/spider.txt", 'a+', encoding="utf-8") as fp:
                    fp.write('*' * 60 + '\n【%d】\n  Title: ' % self.total_articles + self.title_buf[j] + '\n  Link: ' + self.url_buf[j] + '\n  Img: ' + '\r\n\r\n')
                    # fp.write('\n【%d】\n' % self.total_articles + '\n' + url_buf[j] + '\r\n')
                    fp.close()
                    self.Label_Debug(">> 第%d条写入完成：%s" % (j + 1, self.title_buf[j]))
                    print(">> 第%d条写入完成：%s" % (j + 1, self.title_buf[j]))
            print('*' * 60)
            self.get_content(self.title_buf, self.url_buf)
            self.url_buf.clear()
            self.title_buf.clear()

    def Get_WeChat_Subscription(self, token, query):
        if (query == ""):  # 如果查询名为空，使用默认值
            query = "xinhuashefabu1"  # 默认查询新华社公众号
        # 构造搜索公众号的URL
        url = r'https://mp.weixin.qq.com/cgi-bin/searchbiz?action=search_biz&token={0}&lang=zh_CN&f=json&ajax=1&random=0.5182749224035845&query={1}&begin=0&count=5'.format(
            token, query)
        # 发送GET请求获取公众号信息
        html_json = self.sess.get(url, headers=self.headers, timeout=(30, 60)).json()
        fakeid = html_json['list'][0]['fakeid']  # 获取公众号的唯一标识符
        nickname = html_json['list'][0]['nickname']  # 获取公众号的昵称
        self.Label_Debug("nickname: "+nickname)  # 在调试窗口显示公众号昵称
        return fakeid, nickname  # 返回fakeid和昵称

    def Get_Articles(self, token, fakeid):
        # 初始化变量
        img_buf = []  # 用于存储文章封面图片
        Total_buf = []  # 用于存储所有文章的标题，用于去重

        # 构造获取文章列表的URL
        url = r'https://mp.weixin.qq.com/cgi-bin/appmsg?token={0}&lang=zh_CN&f=json&ajax=1&random={1}&action=list_ex&begin=0&count=5&query=&fakeid={2}&type=9'.format(
            token, random.uniform(0, 1), fakeid)

        # 发送请求获取文章列表
        html_json = self.sess.get(url, headers=self.headers, timeout=(30, 60)).json()

        try:
            # 计算总页数
            Total_Page = ceil(int(html_json['app_msg_cnt']) / 5)
            self.progressBar.setMaximum(Total_Page)  # 设置进度条最大值
            QApplication.processEvents()  # 刷新UI
        except Exception as e:
            # 处理异常情况
            print(e)
            self.Label_Debug("!! 失败信息：" + html_json['base_resp']['err_msg'])
            return

        table_index = 0  # 用于表格行索引

        # 启动下载线程
        download_thread = threading.Thread(target=self.download_content)
        download_thread.start()
        self.thread_list.append(download_thread)

        _buf_index = 0  # 用于记录当前页的文章索引
        for i in range(Total_Page):  # 遍历每一页
            # 处理恢复模式
            if self.isresume == 1:
                i = i + self.pagenum

            # 更新进度信息
            self.Label_Debug(
                "第[%d/%d]页  url:%s, article:%s" % (i + 1, Total_Page, self.linkbuf_cnt, self.download_cnt))
            self.label_total_Page.setText("第[%d/%d]页  linkbuf_cnt:%s, download_cnt:%s" % (
            i + 1, Total_Page, self.linkbuf_cnt, self.download_cnt))

            # 构造分页URL
            begin = i * 5
            url = r'https://mp.weixin.qq.com/cgi-bin/appmsg?token={0}&lang=zh_CN&f=json&ajax=1&random={1}&action=list_ex&begin={2}&count=5&query=&fakeid={3}&type=9'.format(
                token, random.uniform(0, 1), begin, fakeid)

            # 发送请求获取分页数据
            while True:
                try:
                    html_json = self.sess.get(url, headers=self.headers, timeout=(30, 60)).json()
                    break
                except Exception as e:
                    print("连接出错，稍等2s", e)
                    self.Label_Debug("连接出错，稍等2s" + str(e))
                    sleep(2)
                    continue

            # 处理文章列表
            try:
                app_msg_list = html_json['app_msg_list']
            except Exception as e:
                self.Label_Debug("！！！操作太频繁，5s后重试！！！")
                print("！！！操作太频繁，5s后重试！！！", e)
                sleep(5)
                continue

            # 如果文章列表为空，结束循环
            if (str(app_msg_list) == '[]'):
                print('结束了')
                self.Label_Debug("结束了")
                break

            # 处理每篇文章
            for j in range(30):
                try:
                    # 检查文章是否已存在
                    if (app_msg_list[j]['title'] in Total_buf):
                        self.Label_Debug("本条已存在，跳过")
                        print("本条已存在，跳过")
                        continue

                    # 检查关键词匹配
                    if self.keyWord != "":
                        if self.keyWord not in app_msg_list[j]['title']:
                            self.Label_Debug("本条不匹配关键词[%s]，跳过" % self.keyWord)
                            print("本条不匹配关键词[%s]，跳过" % self.keyWord)
                            continue

                    # 检查时间范围
                    article_time = int(strftime("%Y", localtime(int(app_msg_list[j]['update_time']))))
                    if (self.timeStart > article_time):
                        self.Label_Debug(
                            "本条[%d]不在时间范围[%d-%d]内，跳过" % (article_time, self.timeStart, self.timeEnd))
                        print("本条[%d]不在时间范围[%d-%d]内，跳过" % (article_time, self.timeStart, self.timeEnd))
                        continue
                    if (article_time > self.timeEnd):
                        self.Label_Debug("达到结束时间，退出")
                        print("达到结束时间，退出")
                        self.Stop_Run()
                        return

                    # 保存文章信息
                    self.title_buf.append(app_msg_list[j]['title'])
                    self.link_buf.append(app_msg_list[j]['link'])
                    img_buf.append(app_msg_list[j]['cover'])
                    Total_buf.append(app_msg_list[j]['title'])

                    # 更新UI表格
                    table_count = self.tableWidget_result.rowCount()
                    if (table_index >= table_count):
                        self.tableWidget_result.insertRow(table_count)
                    self.tableWidget_result.setItem(table_index, 0,
                                                    QtWidgets.QTableWidgetItem(self.title_buf[_buf_index + j]))
                    self.tableWidget_result.setItem(table_index, 1,
                                                    QtWidgets.QTableWidgetItem(self.link_buf[_buf_index + j]))
                    table_index = table_index + 1

                    # 更新文章总数
                    self.total_articles += 1
                    dict_in = {"Title": self.title_buf[_buf_index + j], "Link": self.link_buf[_buf_index + j],
                               "Img": img_buf[_buf_index + j]}
                    self.url_json_once(dict_in)

                    # 创建目录并保存文章信息到文件
                    os.makedirs(self.rootpath, exist_ok=True)
                    with open(self.rootpath + "/spider.txt", 'a+', encoding="utf-8") as fp:
                        fp.write('*' * 60 + '\n【%d】\n  Title: ' % self.total_articles + self.title_buf[
                            _buf_index + j] + '\n  Link: ' + self.link_buf[_buf_index + j] + '\n  Img: ' + img_buf[
                                     _buf_index + j] + '\r\n\r\n')
                        fp.close()

                    # 更新调试信息和配置文件
                    self.Label_Debug(">> 第%d条写入完成：%s" % (self.total_articles, self.title_buf[_buf_index + j]))
                    print(">> 第%d条写入完成：%s" % (self.total_articles, self.title_buf[_buf_index + j]))
                    self.conf.set("resume", "total_articles", str(self.total_articles))
                    self.conf.write(open(self.cfgpath, "r+", encoding="utf-8"))
                except Exception as e:
                    print(">> 本页抓取结束 - ", e)
                    _buf_index += j
                    break

            # 更新进度信息
            self.Label_Debug(">> 一页抓取结束")
            print(">> 一页抓取结束")

            # 更新链接缓冲区计数
            if self.isresume == 1:
                self.linkbuf_cnt = len(self.link_buf) + self.json_read_len
            else:
                self.linkbuf_cnt = len(self.link_buf)
            self.conf.set("resume", "linkbuf_cnt", str(self.linkbuf_cnt))
            self.conf.write(open(self.cfgpath, "r+", encoding="utf-8"))
            self.conf.set("resume", "pagenum", str(i))
            self.conf.write(open(self.cfgpath, "r+", encoding="utf-8"))
            sleep(self.time_gap)  # 等待指定时间间隔

        # 完成提示
        self.Label_Debug_Clear()
        self.Label_Debug(">> 列表抓取结束!!! <<")
        print(">> 列表抓取结束!!! <<")
        self.download_end = 1


    def Get_comment_id(self, article_url):
        '''获取文章id'''
        try:
            resp = requests.get(article_url).text
            pattern = re.compile(r'comment_id\s*=\s*"(?P<id>\d+)"')
            return pattern.search(resp)['id']
        except:
            return None

    def Get_Comments(self, article_url, uin, key, offset=0):
        '''获取文章的评论'''
        # TODO: 微信uin和key失效后，弹窗提示更新，更新后继续运行
        comments = []
        if not uin or not key:
            return comments
        url = 'https://mp.weixin.qq.com/mp/appmsg_comment?'
        biz = re.search('__biz=(.*?)&', article_url).group(1)
        comment_id = self.Get_comment_id(article_url)
        
        datas = {
            'action': 'getcomment',
            # 与文章绑定
            'comment_id': str(comment_id),   # !import
            # 与微信绑定
            'uin': str(uin),                 # !import
            # 与微信绑定，约20分钟失效
            'key': str(key),                 # !import
            '__biz': str(biz),               # !import
            'offset': str(offset),
            'limit': '100',
            'f': 'json',
            # 'scene': '0',
            # 'appmsgid': appmsgid,
            # 'idx': idx,
            # 'send_time': '',
            # 'sessionid': sessionid,
            # 'enterid': enterid,
            # 'fasttmplajax': '1',
            # 'pass_ticket': pass_ticket,
            # 'wxtoken': '',
            # 'devicetype': 'Windows%2B11%2Bx64',
            # 'clientversion': '63090551',
            # 'appmsg_token': '',
            # 'x5': '0',            
        }
        params = ''
        for key,value in datas.items():
            params += key + '=' + value + '&'
        url += params
        try:
            resp = requests.get(url=url).json()
            if resp['elected_comment_total_cnt']:
                for item in resp['elected_comment']:
                    comments.append(item['nick_name'] + ": " + item['content'])
        except:
            pass
        return comments


    # 获取阅读数和点赞数(未测试)
    def Get_ReadsLikes(self, link):
        # 获得mid,_biz,idx,sn 这几个在link中的信息
        mid = link.split("&")[1].split("=")[1]
        idx = link.split("&")[2].split("=")[1]
        sn = link.split("&")[3].split("=")[1]
        _biz = link.split("&")[0].split("_biz=")[1]

        # fillder 中取得一些不变得信息
        pass_ticket = "这里也是输入你自己的数据"#从fiddler中获取 # ---------------------------------这里每次需要修改---------------------------------
        appmsg_token = "这里也是输入你自己的数据"#从fiddler中获取 # ---------------------------------这里每次需要修改---------------------------------

        # 目标url
        url = "http://mp.weixin.qq.com/mp/getappmsgext"#获取详情页的网址
        # 添加Cookie避免登陆操作，这里的"User-Agent"最好为手机浏览器的标识
        #phoneCookie = "自己的"
        phoneCookie = "这里也是输入你自己的数据（这虽然叫phoneCookie但其实就是fillder抓包得到的cookie）"# ---------------------------------这里每次需要修改---------------------------------
        headers = {
            "Cookie": phoneCookie,
            "User-Agent": "这里也是输入你自己的数据" # ---------------------------------这里需要修改---------------------------------
        }
        # 添加data，`req_id`、`pass_ticket`分别对应文章的信息，从fiddler复制即可。
        data = {
            "is_only_read": "1",
            "is_temp_url": "0",
            "appmsg_type": "9",
            'reward_uin_count': '-1'
        }
        """
        添加请求参数
        __biz对应公众号的信息，唯一
        mid、sn、idx分别对应每篇文章的url的信息，需要从url中进行提取
        key、appmsg_token从fiddler上复制即可
        pass_ticket对应的文章的信息，也可以直接从fiddler复制
        """
        params = {
            "__biz": _biz,
            "mid": mid,
            "sn": sn,
            "idx": idx,
            "key": "这里也是输入你自己的数据",# ---------------------------------这里每次需要修改---------------------------------
            "pass_ticket": pass_ticket,
            "appmsg_token": appmsg_token,
            "uin": "这里也是输入你自己的数据",
            "wxtoken": "777",
        }

        # 使用post方法进行提交
        requests.packages.urllib3.disable_warnings()
        content = requests.post(url, headers=headers, data=data, params=params).json()
        # 提取其中的阅读数和点赞数
        # print(content["appmsgstat"]["read_num"], content["appmsgstat"]["like_num"])
        try:
            readNum = content["appmsgstat"]["read_num"]
            print("阅读数:"+str(readNum))
        except:
            readNum = 0
        try:
            likeNum = content["appmsgstat"]["like_num"]
            print("喜爱数:"+str(likeNum))
        except:
            likeNum = 0
        try:
            old_like_num = content["appmsgstat"]["old_like_num"]
            print("在读数:"+str(old_like_num))
        except:
            old_like_num = 0
        # 歇3s，防止被封
        time.sleep(3)
        return readNum, likeNum,old_like_num

    def download_content(self):
        # global link_buf, title_buf
        # self.pri_index = 0
        # 启动一个无限循环，直到被显式中断。
        while 1:
            try:
                # 检查已下载文章数量（download_cnt）是否小于需要下载的文章总数（linkbuf_cnt）。
                if self.download_cnt < self.linkbuf_cnt:
                    # 如果是恢复之前的下载（isresume==1）：读取包含之前保存的URL的JSON文件使用JSON文件中当前下载计数索引的标题和链接调用
                    if self.isresume == 1:
                        self.json_read = self.url_json_read()
                        # print("download_cnt:", self.download_cnt, "; json_read:", len(self.json_read), "; linkbuf_cnt:", self.linkbuf_cnt)
                        self.get_content(self.json_read[self.download_cnt]["Title"],
                                         self.json_read[self.download_cnt]["Link"])
                    else:
                        # 如果不是恢复下载，使用当前缓冲区中的标题和链接调用get_content()。
                        self.get_content(self.title_buf[self.download_cnt], self.link_buf[self.download_cnt])
                    # 增加下载计数
                    self.download_cnt += 1
                    # 使用新的下载计数更新配置文件
                    self.conf.set("resume", "download_cnt", str(self.download_cnt))  # !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                    # 写入更新后的配置，以便后续恢复下载
                    self.conf.write(open(self.cfgpath, "r+", encoding="utf-8"))
                elif self.download_cnt >= self.linkbuf_cnt and self.download_end == 1:
                    # 如果所有文章都已下载且下载过程已结束：
                    # 清除调试标签
                    self.Label_Debug_Clear()
                    # 打印完成消息
                    self.Label_Debug(">> 程序结束, 欢迎再用!!! <<")
                    print(">> 程序结束, 欢迎再用!!! <<")
                    # 跳出无限循环
                    break
                elif self.download_cnt == self.linkbuf_cnt and self.download_end == 0:
                    # 如果下载计数与链接缓冲区计数匹配，但下载过程尚未结束，则等待2秒后再次检查。
                    sleep(2)
            except Exception as e:
                # 捕获并打印下载过程中发生的任何异常，并将其记录到调试标签中。
                print("download_content", e)
                self.Label_Debug(e)


    def get_content(self, title_buf, link_buf):  # 获取地址对应的文章内容
        # 初始化空字符串，准备存储单篇文章的标题。
        each_title = ""  # 初始化
        each_url = ""  # 初始化

        # 根据关键词搜索模式确定要处理的文章数量：关键词搜索模式：处理多篇文章（列表长度）普通模式：只处理一篇文章
        if self.keyword_search_mode == 1:
            length = len(title_buf)
        else:
            length = 1

        # 处理文章标题：使用正则表达式替换标题中的非法文件名字符根据搜索模式选择处理单篇或多篇文章的标题
        for index in range(length):
            if self.keyword_search_mode == 1:
                each_title = re.sub(r'[\|\/\<\>\:\*\?\\\"]', "_", title_buf[index])  # 剔除不合法字符
            else:
                each_title = re.sub(r'[\|\/\<\>\:\*\?\\\"]', "_", title_buf)  # 剔除不合法字符

            # 文件存储准备：创建以文章标题命名的文件夹如果文件夹不存在，创建它切换工作目录到该文件夹
            filepath = self.rootpath + "/" + each_title  # 为每篇文章创建文件夹
            if (not os.path.exists(filepath)):  # 若不存在，则创建文件夹
                os.makedirs(filepath)
            os.chdir(filepath)  # 切换至文件夹

            # 获取文章页面：根据搜索模式选择下载URL使用重试机制获取网页内容连接超时时等待2秒后重试
            download_url = link_buf[index] if self.keyword_search_mode==1 else link_buf
            while True:
                try:
                    html = self.sess.get(download_url, headers=self.headers, timeout=(30, 60))
                    break
                except Exception as e:
                    print("连接出错，稍等2s", e)
                    self.Label_Debug("连接出错，稍等2s" + str(e))
                    sleep(2)
                    continue
            # try:
            #     pdfkit.from_file(html.text, each_title + '.pdf')
            # except Exception as e:
            #     pass

            # 解析文章文本内容：使用BeautifulSoup解析HTML查找具有"rich_media_content"类的元素中的段落如果未找到文章内容，设置标志并记录错误
            soup = BeautifulSoup(html.text, 'lxml')
            try:
                article = soup.find(class_="rich_media_content").find_all("p")  # 查找文章内容位置
                No_article = 0
            except Exception as e:
                No_article = 1
                self.Label_Debug("本篇未匹配到文字 ->"+str(e))
                print("本篇未匹配到文字 ->", e)
                pass

            # 查找文章图片：在"rich_media_content"类中查找所有图片如果未找到图片，设置标志并记录错误
            try:
                img_urls = soup.find(class_="rich_media_content").find_all("img")  # 获得文章图片URL集
                No_img = 0
            except Exception as e:
                No_img = 1
                self.Label_Debug("本篇未匹配到图片 ->" + str(e))
                print("本篇未匹配到图片 ->", e)
                pass

            # 保存文章文本：遍历文章段落提取每个段落的文本将非空文本写入同名.txt文件记录保存完成的日志
            print("*" * 60)
            self.Label_Debug("*" * 30)
            self.Label_Debug(each_title)
            if No_article != 1:
                for i in article:
                    line_content = i.get_text()  # 获取标签内的文本
                    # print(line_content)
                    if (line_content != None):  # 文本不为空
                        with open(each_title + r'.txt', 'a+', encoding='utf-8') as fp:
                            fp.write(line_content + "\n")  # 写入本地文件
                            fp.close()
                self.Label_Debug(">> 保存文档 - 完毕!")
                # print(">> 标题：", each_title)
                print(">> 保存文档 - 完毕!")

            # 下载和保存图片：遍历图片URL使用重试机制下载图片最多重试3次，失败则跳过将图片保存为本地JPEG文件记录保存图片数量的日志
            if No_img != 1:
                for i in range(len(img_urls)):
                    re_cnt = 0
                    while True:
                        try:
                            pic_down = self.sess.get(img_urls[i]["data-src"], timeout=(30, 60))  # 连接超时30s，读取超时60s，防止卡死
                            break
                        except Exception as e:
                            print("下载超时 ->", e)
                            self.Label_Debug("下载超时->" + str(e))
                            re_cnt += 1
                            if re_cnt > 3:
                                print("放弃此图")
                                self.Label_Debug("放弃此图")
                                break
                    if re_cnt > 3:
                        f = open(str(i) + r'.jpeg', 'ab+')
                        f.close()
                        continue
                    img_urls[i]["src"] = str(i)+r'.jpeg'  # 更改图片地址为本地
                    with open(str(i) + r'.jpeg', 'ab+') as fp:
                        fp.write(pic_down.content)
                        fp.close()
                self.Label_Debug(">> 保存图片%d张 - 完毕!" % len(img_urls))
                print(">> 保存图片%d张 - 完毕!" % len(img_urls))

            # 保存完整的HTML文件：将解析后的HTML内容写入同名.html文件记录HTML保存完成的日志
            with open(each_title+r'.html', 'w', encoding='utf-8') as f:  # 保存html文件
                f.write(str(soup))
                f.close()
                self.Label_Debug(">> 保存html - 完毕!")
                # pdfkit.from_file('test.html','out1.pdf')
                print(">> 保存html - 完毕!")

            # 获取并保存文章评论：调用Get_Comments()方法获取评论将评论写入同名_comments.txt文件记录评论保存完成的日志
            # 下载文章评论
            comments = self.Get_Comments(download_url, self.wechat_uin, self.wechat_key)
            with open(each_title + r'_comments.txt', 'a+', encoding='utf-8') as fp:
                fp.write('\n'.join(comments))  # 写入本地文件
                fp.close()
                self.Label_Debug(">> 保存评论 - 完毕!")

            # 关键词搜索模式下的等待：在处理每篇文章后等待指定时间防止过快地发送请求
            if self.keyword_search_mode == 1:
                self.Label_Debug(">> 休息 %d s" % self.time_gap)
                print(">> 休息 %d s" % self.time_gap)
                sleep(self.time_gap)

################################强制关闭线程##################################################
    def _async_raise(self, tid, exctype):
        """raises the exception, performs cleanup if needed"""
        tid = ctypes.c_long(tid)
        if not inspect.isclass(exctype):
            exctype = type(exctype)
        res = ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, ctypes.py_object(exctype))
        if res == 0:
            raise ValueError("invalid thread id")
        elif res != 1:
            # """if it returns a number greater than one, you're in trouble,
            # and you should call it again with exc=NULL to revert the effect"""
            ctypes.pythonapi.PyThreadState_SetAsyncExc(tid, None)
            raise SystemError("PyThreadState_SetAsyncExc failed")

    def stop_thread(self, thread):
        self._async_raise(thread.ident, SystemExit)
###############################################################################################

def main():
    app = QtWidgets.QApplication(sys.argv)
    MainWindow = QtWidgets.QMainWindow()
    ui = MyMainWindow()
    ui.setupUi(MainWindow)
    MainWindow.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()



















#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : lusheng
import email
import threading
from tkinter import *
from tkinter import messagebox
import sqlite3
from datetime import datetime, date, timedelta
import time
import requests
import re
import json
from email.utils import parseaddr, formataddr
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from email.mime.multipart import MIMEMultipart
import logging
import win32api, win32con, win32gui_struct, win32gui
import os


def init_db(root):
    # 连接
    conn = sqlite3.connect("announcement.db")
    c = conn.cursor()

    # 创建表
    c.execute('''DROP TABLE IF EXISTS announcement ''')  # 删除旧表，如果存在（因为这是临时数据）
    c.execute('''
        CREATE TABLE announcement (
        id INTEGER PRIMARY KEY AUTOINCREMENT, 
        companyCd INTEGER , 
        companyName text,
        title text,
        publishDate text, 
        filePath text)
        ''')
    conn.commit()
    conn.close()
    label70 = Label(root, text='数据库初始化成功', font=('楷体', 10), fg='black')
    label70.grid(row=7, column=0, columnspan=2, sticky=W)


def _format_addr(s):
    name, addr = parseaddr(s)
    return formataddr((
        Header(name, 'utf-8').encode(),
        addr))


# 发送邮件
def sendMails(receivers, companyCd, companyName, disclosureTitle, publishDate, destFilePath):
    mail_host = "smtp.qq.com"  # 设置服务器

    mail_user = "228383562@qq.com"  # 用户名
    mail_pass = "waajnvtmdhiucbef"  # 口令
    sender = '228383562@qq.com'
    # mail_host = "smtp.163.com"  # 设置服务器
    # mail_user = "zj_mao@163.com"  # 用户名
    # mail_pass = "ZWTBWGYTNLOOSEDC"  # 口令
    # sender = 'zj_mao@163.com'

    receiversName = '收件邮箱'
    receivers = receivers
    mail_msg = """
    公司代码：%s <br>
    公司名称：%s <br>
    公告日期：%s <br>
    公告标题：%s <br>
    公告链接：%s <br>
    <p>这是全国中小企业股份转让系统查询推送</p>
    """ % (companyCd, companyName, publishDate, disclosureTitle, destFilePath)

    main_msg = MIMEMultipart()
    main_msg.attach(MIMEText(mail_msg, 'html', 'utf-8'))
    main_msg['From'] = _format_addr(u'<%s>' % mail_user)
    main_msg['To'] = _format_addr(u'<%s>' % receivers)
    main_msg['Subject'] = Header(disclosureTitle, 'utf-8').encode()

    try:
        smtpObj = smtplib.SMTP('smtp.qq.com', 587)
        # smtp = smtplib.SMTP_SSL(mailserver)
        # smtpObj = smtplib.SMTP('smtp.163.com', 25)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.login(mail_user, mail_pass)
        smtpObj.sendmail(sender, receivers, main_msg.as_string())
        smtpObj.quit()
        print("邮件发送成功")
        logging.info("邮件发送成功")
        return 1
    except smtplib.SMTPException as e:
        print("无法发送邮件")
        logging.error("Error: 无法发送邮件, %s" % e)
        return 0

def fun_timer(s,companyCd,companyName,startTime,endTime,receiveMail):
    print('Hello Timer!')
    global timer
    run(s,companyCd,companyName,startTime,endTime,receiveMail)
    timer = threading.Timer(600, fun_timer(s,companyCd,companyName,startTime,endTime,receiveMail))
    timer.start()



def run(s,companyCd,companyName,startTime,endTime,receiveMail):
    # companyCd = root.entry.get()
    # companyName = root.entry11.get().strip()
    # startTime = root.entry21.get().strip()
    # endTime = root.entry31.get().strip()
    # receiveMail = root.entry41.get().strip()
    if companyCd == '' and companyName == '':
        messagebox.showinfo('提示', '公司名称和公司代码不能都为空，请检查')
        return
    if companyCd != '' and companyName != '':
        messagebox.showinfo('提示', '公司名称和公司代码不能同时填写，请检查')
        return
    if len(startTime) != 10:
        messagebox.showinfo('提示', '开始日期格式不对，请检查')
        return
    # if len(endTime) != 10:
    #     messagebox.showinfo('提示', '结束日期格式不对，请检查')
    #     return
    # h = win32gui.FindWindow('TkTopLevel','中小企业股份转让系统公告查询工具')
    # win32gui.ShowWindow(h,win32con.SW_HIDE)
    while 1:
        # companyCd = entry.get().strip()
        if companyCd != '':
            companyCd_list = companyCd.split(';')
            for ccd in companyCd_list:
                db = sqlite3.connect("announcement.db")
                c = db.cursor()
                data = {
                    "disclosureType[]": 5,
                    "disclosureSubtype[]": None,
                    "page": 0,
                    "startTime": startTime,
                    "endTime": endTime,
                    "companyCd": ccd,
                    "isNewThree": 1,
                    "keyword": None,
                    "xxfcbj[]": None,
                    "hyType[]": None,
                    "needFields[]": ["companyCd", "companyName", "disclosureTitle", "destFilePath", "publishDate",
                                     "xxfcbj",
                                     "destFilePath", "fileExt", "xxzrlx"],
                    "sortfield": "xxssdq",
                    "sorttype": "asc",
                }
                news_list = []
                response1 = requests.post(URL, data)
                # print(response1.text)
                response2 = re.search('(?<=\(\[)(.*?)(?=]\))', response1.text).group()
                # print(response2)
                j = json.loads(response2)['listInfo']
                if j['content'] == []:
                    messagebox.showinfo('提示', '没有查询到信息，请检查公司代码或名称是否正确')
                    return
                else:
                    totalElements = j['totalElements']
                    totalPages = j['totalPages']
                    logging.info("通过代码%s查询到%d条公告，共%d页" % (ccd, totalElements, totalPages))
                    # label70 = Label(root, text="通过代码%s查询到%d条公告，共%d页" % (ccd, totalElements, totalPages),
                    #                 font=('楷体', 12), fg='black')
                    # label70.grid(row=7, column=0, columnspan=2, sticky=W)

                    for n in range(totalPages):
                        data = {
                            "disclosureType[]": 5,
                            "disclosureSubtype[]": None,
                            "page": n,
                            "startTime": startTime,
                            "endTime": endTime,
                            "companyCd": ccd,
                            "isNewThree": 1,
                            "keyword": None,
                            "xxfcbj[]": None,
                            "hyType[]": None,
                            "needFields[]": ["companyCd", "companyName", "disclosureTitle", "destFilePath",
                                             "publishDate",
                                             "xxfcbj",
                                             "destFilePath", "fileExt", "xxzrlx"],
                            "sortfield": "xxssdq",
                            "sorttype": "asc",
                        }
                        logging.info("正在处理第%d页" % (n + 1))
                        # label80 = Label(root, text="正在处理第%d页" % (n + 1), font=('楷体', 12), fg='black')
                        # label80.grid(row=8, column=0, columnspan=2, sticky=W)

                        response3 = requests.post(URL, data)
                        response4 = re.search('(?<=\(\[)(.*?)(?=]\))', response3.text).group()
                        j = json.loads(response4)['listInfo']
                        list = j['content']
                        # 循环本页内容查询数据库、发送邮件
                        for li in list:
                            # print(li)
                            companyCd2 = li['companyCd']
                            companyName2 = li['companyName']
                            destFilePath = "http://www.neeq.com.cn" + li['destFilePath']
                            disclosureTitle = li['disclosureTitle']
                            publishDate = li['publishDate']
                            xxfcbj = li['xxfcbj']
                            xxzrlx = li['xxzrlx']
                            result = c.execute("SELECT * FROM announcement where filePath = '%s'" % destFilePath)
                            if result.fetchone():
                                print(disclosureTitle, " 该公告数据库中已存在\n")
                                logging.info(disclosureTitle + " 该公告数据库中已存在")
                                # label90 = Label(root, text=disclosureTitle + " 该公告数据库中已存在", font=('楷体', 12), fg='black')
                                # label90.grid(row=9, column=0, columnspan=2, sticky=W)
                            else:
                                # 发送邮件
                                mailResult = sendMails(receiveMail, companyCd2, companyName2, disclosureTitle, publishDate,
                                          destFilePath)
                                # print(mailResult)
                                if mailResult == 1:
                                    data = "NULL,\'%s\',\'%s\',\'%s\',\'%s\',\'%s\'" % (
                                        companyCd2, companyName2, disclosureTitle, publishDate, destFilePath)
                                    # print(data, "\n")
                                    c.execute('INSERT INTO announcement VALUES (%s)' % data)
                                    db.commit()
                                    print(disclosureTitle, " 该公告已存入数据库\n")
                                    logging.info(disclosureTitle + " 该公告已存入数据库")
                                # label90 = Label(root, text=disclosureTitle + " 该公告已存入数据库", font=('楷体', 12), fg='black')
                                # label90.grid(row=9, column=0, columnspan=2, sticky=W)
                                time.sleep(5)
                    time.sleep(20)  # 获取一个页面后休息3秒，防止请求服务器过快
                db.close()

        if companyName != '':
            companyName_list = companyName.split(';')
            for keyword in companyName_list:
                db = sqlite3.connect("announcement.db")
                c = db.cursor()
                data = {
                    "disclosureType[]": 5,
                    "disclosureSubtype[]": None,
                    "page": 0,
                    "startTime": startTime,
                    "endTime": endTime,
                    "companyCd": None,
                    "isNewThree": 1,
                    "keyword": keyword,
                    "xxfcbj[]": None,
                    "hyType[]": None,
                    "needFields[]": ["companyCd", "companyName", "disclosureTitle", "destFilePath", "publishDate",
                                     "xxfcbj",
                                     "destFilePath", "fileExt", "xxzrlx"],
                    "sortfield": "xxssdq",
                    "sorttype": "asc",
                }
                news_list = []
                response1 = requests.post(URL, data)
                # print(response1.text)
                response2 = re.search('(?<=\(\[)(.*?)(?=]\))', response1.text).group()
                # print(response2)
                j = json.loads(response2)['listInfo']
                if j['content'] == []:
                    messagebox.showinfo('提示', '没有查询到信息，请检查公司代码或名称是否正确')
                    return
                else:
                    totalElements = j['totalElements']
                    totalPages = j['totalPages']
                    logging.info("通过关键字%s查询到%d条公告，共%d页" % (keyword, totalElements, totalPages))
                    # label70 = Label(root, text="通过代码%s查询到%d条公告，共%d页" % (ccd, totalElements, totalPages),
                    #                 font=('楷体', 12), fg='black')
                    # label70.grid(row=7, column=0, columnspan=2, sticky=W)

                    for n in range(totalPages):
                        data = {
                            "disclosureType[]": 5,
                            "disclosureSubtype[]": None,
                            "page": n,
                            "startTime": startTime,
                            "endTime": endTime,
                            "companyCd": companyCd,
                            "isNewThree": 1,
                            "keyword": keyword,
                            "xxfcbj[]": None,
                            "hyType[]": None,
                            "needFields[]": ["companyCd", "companyName", "disclosureTitle", "destFilePath",
                                             "publishDate",
                                             "xxfcbj",
                                             "destFilePath", "fileExt", "xxzrlx"],
                            "sortfield": "xxssdq",
                            "sorttype": "asc",
                        }
                        logging.info("正在处理第%d页" % (n + 1))
                        # label80 = Label(root, text="正在处理第%d页" % (n + 1), font=('楷体', 12), fg='black')
                        # label80.grid(row=8, column=0, columnspan=2, sticky=W)

                        response3 = requests.post(URL, data)
                        response4 = re.search('(?<=\(\[)(.*?)(?=]\))', response3.text).group()
                        j = json.loads(response4)['listInfo']
                        list = j['content']
                        # 循环本页内容查询数据库、发送邮件
                        for li in list:
                            # print(li)
                            companyCd2 = li['companyCd']
                            companyName2 = li['companyName']
                            destFilePath = "http://www.neeq.com.cn" + li['destFilePath']
                            disclosureTitle = li['disclosureTitle']
                            publishDate = li['publishDate']
                            xxfcbj = li['xxfcbj']
                            xxzrlx = li['xxzrlx']
                            result = c.execute("SELECT * FROM announcement where filePath = '%s'" % destFilePath)
                            if result.fetchone():
                                print(disclosureTitle, " 该公告数据库中已存在\n")
                                logging.info(disclosureTitle + " 该公告数据库中已存在")
                                # label90 = Label(root, text=disclosureTitle + " 该公告数据库中已存在", font=('楷体', 12), fg='black')
                                # label90.grid(row=9, column=0, columnspan=2, sticky=W)
                            else:
                                # 发送邮件
                                mailResult = sendMails(receiveMail, companyCd2, companyName2, disclosureTitle, publishDate,
                                          destFilePath)
                                # print(mailResult)
                                if mailResult == 1:
                                    data = "NULL,\'%s\',\'%s\',\'%s\',\'%s\',\'%s\'" % (
                                        companyCd2, companyName2, disclosureTitle, publishDate, destFilePath)
                                    # print(data, "\n")
                                    c.execute('INSERT INTO announcement VALUES (%s)' % data)
                                    db.commit()
                                    print(disclosureTitle, " 该公告已存入数据库\n")
                                    logging.info(disclosureTitle + " 该公告已存入数据库")

                                # label90 = Label(root, text=disclosureTitle + " 该公告已存入数据库", font=('楷体', 12), fg='black')
                                # label90.grid(row=9, column=0, columnspan=2, sticky=W)
                                time.sleep(1)
                    time.sleep(3)  # 获取一个页面后休息3秒，防止请求服务器过快
                db.close()
        logging.info("本次查询结束，10分钟后开始下次查询")
        time.sleep(600)
    # for companyCd in companyCd_list:


class SysTrayIcon(object):
    '''SysTrayIcon类用于显示任务栏图标'''
    QUIT = 'QUIT'
    SPECIAL_ACTIONS = [QUIT]
    FIRST_ID = 5320

    def __init__(s, icon, hover_text, menu_options, on_quit, tk_window=None, default_menu_index=None,
                 window_class_name=None):
        '''
        icon         需要显示的图标文件路径
        hover_text   鼠标停留在图标上方时显示的文字
        menu_options 右键菜单，格式: (('a', None, callback), ('b', None, (('b1', None, callback),)))
        on_quit      传递退出函数，在执行退出时一并运行
        tk_window    传递Tk窗口，s.root，用于单击图标显示窗口
        default_menu_index 不显示的右键菜单序号
        window_class_name  窗口类名
        '''
        s.icon = icon
        s.hover_text = hover_text
        s.on_quit = on_quit
        s.root = tk_window

        menu_options = menu_options + (('退出', None, s.QUIT),)
        s._next_action_id = s.FIRST_ID
        s.menu_actions_by_id = set()
        s.menu_options = s._add_ids_to_menu_options(list(menu_options))
        s.menu_actions_by_id = dict(s.menu_actions_by_id)
        del s._next_action_id

        s.default_menu_index = (default_menu_index or 0)
        s.window_class_name = window_class_name or "SysTrayIconPy"

        message_map = {win32gui.RegisterWindowMessage("TaskbarCreated"): s.restart,
                       win32con.WM_DESTROY: s.destroy,
                       win32con.WM_COMMAND: s.command,
                       win32con.WM_USER + 20: s.notify, }
        # 注册窗口类。
        window_class = win32gui.WNDCLASS()
        window_class.hInstance = win32gui.GetModuleHandle(None)
        window_class.lpszClassName = s.window_class_name
        window_class.style = win32con.CS_VREDRAW | win32con.CS_HREDRAW;
        window_class.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
        window_class.hbrBackground = win32con.COLOR_WINDOW
        window_class.lpfnWndProc = message_map  # 也可以指定wndproc.
        s.classAtom = win32gui.RegisterClass(window_class)
        s.update()

    def update(s):
        '''显示任务栏图标，不用每次都重新创建新的托盘图标'''
        # 创建窗口。
        hinst = win32gui.GetModuleHandle(None)
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        s.hwnd = win32gui.CreateWindow(s.classAtom,
                                       s.window_class_name,
                                       style,
                                       0,
                                       0,
                                       win32con.CW_USEDEFAULT,
                                       win32con.CW_USEDEFAULT,
                                       0,
                                       0,
                                       hinst,
                                       None)
        win32gui.UpdateWindow(s.hwnd)
        s.notify_id = None
        s.refresh_icon()

        win32gui.PumpMessages()

    def _add_ids_to_menu_options(s, menu_options):
        result = []
        for menu_option in menu_options:
            option_text, option_icon, option_action = menu_option
            if callable(option_action) or option_action in s.SPECIAL_ACTIONS:
                s.menu_actions_by_id.add((s._next_action_id, option_action))
                result.append(menu_option + (s._next_action_id,))
            else:
                result.append((option_text,
                               option_icon,
                               s._add_ids_to_menu_options(option_action),
                               s._next_action_id))
            s._next_action_id += 1
        return result

    def refresh_icon(s):
        # 尝试找到自定义图标
        hinst = win32gui.GetModuleHandle(None)
        if os.path.isfile(s.icon):
            icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
            hicon = win32gui.LoadImage(hinst,
                                       s.icon,
                                       win32con.IMAGE_ICON,
                                       0,
                                       0,
                                       icon_flags)
        else:  # 找不到图标文件 - 使用默认值
            hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)

        if s.notify_id:
            message = win32gui.NIM_MODIFY
        else:
            message = win32gui.NIM_ADD
        s.notify_id = (s.hwnd,
                       0,
                       win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP,
                       win32con.WM_USER + 20,
                       hicon,
                       s.hover_text)
        win32gui.Shell_NotifyIcon(message, s.notify_id)

    def restart(s, hwnd, msg, wparam, lparam):
        s.refresh_icon()

    def destroy(s, hwnd=None, msg=None, wparam=None, lparam=None, exit=1):
        nid = (s.hwnd, 0)
        win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
        win32gui.PostQuitMessage(0)  # 终止应用程序。
        if exit and s.on_quit:
            s.on_quit()  # 需要传递自身过去时用 s.on_quit(s)
        else:
            s.root.deiconify()  # 显示tk窗口

    def notify(s, hwnd, msg, wparam, lparam):
        if lparam == win32con.WM_LBUTTONDBLCLK:  # 双击左键
            pass
        elif lparam == win32con.WM_RBUTTONUP:  # 右键弹起
            s.show_menu()
        elif lparam == win32con.WM_LBUTTONUP:  # 左键弹起
            s.destroy(exit=0)
        return True
        """
        可能的鼠标事件：
          WM_MOUSEMOVE      #光标经过图标
          WM_LBUTTONDOWN    #左键按下
          WM_LBUTTONUP      #左键弹起
          WM_LBUTTONDBLCLK  #双击左键
          WM_RBUTTONDOWN    #右键按下
          WM_RBUTTONUP      #右键弹起
          WM_RBUTTONDBLCLK  #双击右键
          WM_MBUTTONDOWN    #滚轮按下
          WM_MBUTTONUP      #滚轮弹起
          WM_MBUTTONDBLCLK  #双击滚轮
        """

    def show_menu(s):
        menu = win32gui.CreatePopupMenu()
        s.create_menu(menu, s.menu_options)

        pos = win32gui.GetCursorPos()
        win32gui.SetForegroundWindow(s.hwnd)
        win32gui.TrackPopupMenu(menu,
                                win32con.TPM_LEFTALIGN,
                                pos[0],
                                pos[1],
                                0,
                                s.hwnd,
                                None)
        win32gui.PostMessage(s.hwnd, win32con.WM_NULL, 0, 0)

    def create_menu(s, menu, menu_options):
        for option_text, option_icon, option_action, option_id in menu_options[::-1]:
            if option_icon:
                option_icon = s.prep_menu_icon(option_icon)

            if option_id in s.menu_actions_by_id:
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                wID=option_id)
                win32gui.InsertMenuItem(menu, 0, 1, item)
            else:
                submenu = win32gui.CreatePopupMenu()
                s.create_menu(submenu, option_action)
                item, extras = win32gui_struct.PackMENUITEMINFO(text=option_text,
                                                                hbmpItem=option_icon,
                                                                hSubMenu=submenu)
                win32gui.InsertMenuItem(menu, 0, 1, item)

    def prep_menu_icon(s, icon):
        # 加载图标。
        ico_x = win32api.GetSystemMetrics(win32con.SM_CXSMICON)
        ico_y = win32api.GetSystemMetrics(win32con.SM_CYSMICON)
        hicon = win32gui.LoadImage(0, icon, win32con.IMAGE_ICON, ico_x, ico_y, win32con.LR_LOADFROMFILE)

        hdcBitmap = win32gui.CreateCompatibleDC(0)
        hdcScreen = win32gui.GetDC(0)
        hbm = win32gui.CreateCompatibleBitmap(hdcScreen, ico_x, ico_y)
        hbmOld = win32gui.SelectObject(hdcBitmap, hbm)
        brush = win32gui.GetSysColorBrush(win32con.COLOR_MENU)
        win32gui.FillRect(hdcBitmap, (0, 0, 16, 16), brush)
        win32gui.DrawIconEx(hdcBitmap, 0, 0, hicon, ico_x, ico_y, 0, 0, win32con.DI_NORMAL)
        win32gui.SelectObject(hdcBitmap, hbmOld)
        win32gui.DeleteDC(hdcBitmap)

        return hbm

    def command(s, hwnd, msg, wparam, lparam):
        id = win32gui.LOWORD(wparam)
        s.execute_menu_option(id)

    def execute_menu_option(s, id):
        menu_action = s.menu_actions_by_id[id]
        if menu_action == s.QUIT:
            win32gui.DestroyWindow(s.hwnd)
        else:
            menu_action(s)

class _Main:  #调用SysTrayIcon的Demo窗口
    def __init__(s):
        s.SysTrayIcon  = None  # 判断是否打开系统托盘图标

    def main(s):
        #tk窗口
        s.root = Tk()
        s.root.title('中小企业股份转让系统公告查询工具')
        s.root.geometry('430x190')
        s.root.geometry('+400+200')

        global entry,entry11,entry21,entry31,entry41
        # 文本输入框前的提示文本
        label = Label(s.root, text='公司代码：', width=10, font=('楷体', 12), fg='black')
        label.grid(row=0, column=0, )
        # 文本输入框-公司代码
        entry = Entry(s.root, font=('微软雅黑', 12), width=35)
        entry.grid(row=0, column=1, sticky=W)

        # 文本输入框前的提示文本
        label10 = Label(s.root, text='公司名称：', width=10, font=('楷体', 12), fg='black')
        label10.grid(row=1, column=0)
        # 文本输入框-公司名称
        entry11 = Entry(s.root, font=('微软雅黑', 12), width=35)
        entry11.grid(row=1, column=1, sticky=W)

        # 开始日期
        label20 = Label(s.root, text='起始日期：', width=10, font=('楷体', 12), fg='black')
        label20.grid(row=2, column=0)
        # 文本输入框-开始日期
        sd = StringVar()
        sd.set((datetime.today() + timedelta(days=-30)).strftime("%Y-%m-%d"))
        entry21 = Entry(s.root, textvariable=sd, font=('微软雅黑', 12), width=35)
        entry21.grid(row=2, column=1, sticky=W)
        # print(entry21.get())
        # 结束日期
        label30 = Label(s.root, text='结束日期：', width=10, font=('楷体', 12), fg='black')
        label30.grid(row=3, column=0)
        # 文本输入框-结束日期
        # datetime.today()
        # ed = StringVar()
        # ed.set(datetime.today().strftime("%Y-%m-%d"))
        # entry31 = Entry(root, textvariable=ed, font=('微软雅黑', 12), width=29)
        entry31 = Entry(s.root, font=('微软雅黑', 12), width=35)
        entry31.grid(row=3, column=1, sticky=W)

        # 收件邮箱
        label40 = Label(s.root, text='收件邮箱：', font=('楷体', 12), fg='black')
        label40.grid(row=4, column=0)
        # 文本输入框-收件邮箱
        receiveMail = StringVar()
        receiveMail.set('lusheng1234@126.com')
        # receiveMail.set('lusheng1234@126.com')
        entry41 = Entry(s.root, textvariable=receiveMail, font=('微软雅黑', 12), width=35)
        entry41.grid(row=4, column=1, sticky=W)

        # 初始化数据库
        button50 = Button(s.root, text='初始化', width=8, font=('幼圆', 12), fg='purple', command=lambda:init_db(s.root))
        button50.grid(row=5, column=0)

        # 开始按钮

        button51 = Button(s.root, text='开始', width=20, font=('幼圆', 12), fg='purple', command=lambda:fun_timer(s,entry.get().strip(),entry11.get().strip(),entry21.get().strip(),entry31.get().strip(),entry41.get().strip()))
        button51.grid(row=5, column=1)

        label60 = Label(s.root, text='  ', font=('楷体', 6), fg='black')
        label60.grid(row=6, column=0, columnspan=2, sticky=W)
        # label70 = Label(root, text='  ', font=('楷体', 8), fg='black')
        # label70.grid(row=6, column=0, columnspan=2, sticky=W)
        # label80 = Label(root, text='  ', font=('楷体', 8), fg='black')
        # label80.grid(row=6, column=0, columnspan=2, sticky=W)
        # label90 = Label(root, text='  ', font=('楷体', 8), fg='black')
        # label90.grid(row=6, column=0, columnspan=2, sticky=W)

        # 执行信息
        # label70 = Label(root, text='执行情况：' + 'dasdadas888888888888888888888888', font=('楷体', 12), fg='black')
        # label70.grid(row=7, column=0, columnspan=2, sticky=W)
        s.root.bind("<Unmap>", lambda event: s.Hidden_window() if s.root.state() == 'iconic' else False) #窗口最小化判断，可以说是调用最重要的一步
        s.root.protocol('WM_DELETE_WINDOW', s.exit) #点击Tk窗口关闭时直接调用s.exit，不使用默认关闭
        s.root.resizable(0,0)  #锁定窗口大小不能改变
        s.root.mainloop()

    def switch_icon(s, _sysTrayIcon, icon = 'icon.ico'):
        #点击右键菜单项目会传递SysTrayIcon自身给引用的函数，所以这里的_sysTrayIcon = s.sysTrayIcon
        #只是一个改图标的例子，不需要的可以删除此函数
        # _sysTrayIcon.icon = icon
        fun_timer(s, entry.get().strip(), entry11.get().strip(), entry21.get().strip(), entry31.get().strip(),
            entry41.get().strip())
        return

    def Hidden_window(s, icon = 'icon.ico', hover_text = "公告查询"):
        '''隐藏窗口至托盘区，调用SysTrayIcon的重要函数'''

        #托盘图标右键菜单, 格式: ('name', None, callback),下面也是二级菜单的例子
        #24行有自动添加‘退出’，不需要的可删除
        # menu_options = ((None, None, s.switch_icon),)
        menu_options = (('开始查询', None, s.switch_icon),)
        #                 ('二级 菜单', None, (('更改 图标', None, s.switch_icon), )))

        s.root.withdraw()   #隐藏tk窗口
        if s.SysTrayIcon: s.SysTrayIcon.update()  #已经有托盘图标时调用 update 来更新显示托盘图标
        else: s.SysTrayIcon = SysTrayIcon(icon,               #图标
                                          hover_text,         #光标停留显示文字
                                          menu_options,       #右键菜单
                                          on_quit = s.exit,   #退出调用
                                          tk_window = s.root, #Tk窗口
                                          )

    def exit(s, _sysTrayIcon = None):
        s.root.destroy()
        print ('exit...')

logging.basicConfig(filename='./log.log', format='[%(asctime)s-%(filename)s-%(levelname)s:%(message)s]',
                    level=logging.DEBUG, filemode='a', datefmt='%Y-%m-%d %I:%M:%S %p')
URL = 'http://www.neeq.com.cn/disclosureInfoController/infoResult_zh.do?callback=jQuery331_1596699678177'


if __name__ == '__main__':
    Main = _Main()
    Main.main()

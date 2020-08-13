#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : lusheng
import os
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
import win32api
import win32gui
import win32con


def init_db():
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

    # mail_user = "228383562@qq.com"  # 用户名
    # mail_pass = "waajnvtmdhiucbef"  # 口令
    # sender = '228383562@qq.com'

    mail_user = "610559273@qq.com"  # 用户名
    mail_pass = "xrljvzsvdzbbbfbb"  # 口令
    sender = '610559273@qq.com'

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
    main_msg['From'] = _format_addr(u'公告推送<%s>' % mail_user)
    main_msg['To'] = _format_addr(u'公告收件<%s>' % receivers)
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


def run():
    companyCd = entry.get().strip()
    companyName = entry11.get().strip()
    startTime = entry21.get().strip()
    endTime = entry31.get().strip()
    receiveMail = entry41.get().strip()
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
    h = win32gui.FindWindow('TkTopLevel','中小企业股份转让系统公告查询工具')
    win32gui.ShowWindow(h,win32con.SW_HIDE)

    hinst = win32api.GetModuleHandle(None)
    iconPathName = "icon.ico"
    if os.path.isfile(iconPathName):
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        hicon = win32gui.LoadImage(hinst, iconPathName, win32con.IMAGE_ICON, 0, 0, icon_flags)
    else:
        print('???icon???????')
        hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION)
    flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
    nid = (h, 0, flags, win32con.WM_USER + 20, hicon, "公告推送")
    try:
        win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)
    except:
        print("Failed to add the taskbar icon - is explorer running?")


    # flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
    # nid = (h, 0, win32gui.NIF_INFO, win32con.WM_USER + 20, 'icon.ico', "tooltip")
    # win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid)

    while 1:
        companyCd = entry.get().strip()
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


logging.basicConfig(filename='./log.log', format='[%(asctime)s-%(filename)s-%(levelname)s:%(message)s]',
                    level=logging.DEBUG, filemode='a', datefmt='%Y-%m-%d %I:%M:%S %p')
URL = 'http://www.neeq.com.cn/disclosureInfoController/infoResult_zh.do?callback=jQuery331_1596699678177'
root = Tk()
root.title('中小企业股份转让系统公告查询工具')
root.geometry('430x190')
root.geometry('+400+200')
# 文本输入框前的提示文本
label = Label(root, text='公司代码：', width=10, font=('楷体', 12), fg='black')
label.grid(row=0, column=0, )
# 文本输入框-公司代码
entry = Entry(root, font=('微软雅黑', 12), width=35)
entry.grid(row=0, column=1, sticky=W)

# 文本输入框前的提示文本
label10 = Label(root, text='公司名称：', width=10, font=('楷体', 12), fg='black')
label10.grid(row=1, column=0)
# 文本输入框-公司名称
entry11 = Entry(root, font=('微软雅黑', 12), width=35)
entry11.grid(row=1, column=1, sticky=W)

# 开始日期
label20 = Label(root, text='起始日期：', width=10, font=('楷体', 12), fg='black')
label20.grid(row=2, column=0)
# 文本输入框-开始日期
sd = StringVar()
sd.set((datetime.today() + timedelta(days=-30)).strftime("%Y-%m-%d"))
entry21 = Entry(root, textvariable=sd, font=('微软雅黑', 12), width=35)
entry21.grid(row=2, column=1, sticky=W)

# 结束日期
label30 = Label(root, text='结束日期：', width=10, font=('楷体', 12), fg='black')
label30.grid(row=3, column=0)
# 文本输入框-结束日期
# datetime.today()
# ed = StringVar()
# ed.set(datetime.today().strftime("%Y-%m-%d"))
# entry31 = Entry(root, textvariable=ed, font=('微软雅黑', 12), width=29)
entry31 = Entry(root, font=('微软雅黑', 12), width=35)
entry31.grid(row=3, column=1, sticky=W)

# 收件邮箱
label40 = Label(root, text='收件邮箱：', font=('楷体', 12), fg='black')
label40.grid(row=4, column=0)
# 文本输入框-收件邮箱
receiveMail = StringVar()
receiveMail.set('610559273@qq.com')
# receiveMail.set('lusheng1234@126.com')
entry41 = Entry(root, textvariable=receiveMail, font=('微软雅黑', 12), width=35)
entry41.grid(row=4, column=1, sticky=W)

# 初始化数据库
button50 = Button(root, text='初始化', width=8, font=('幼圆', 12), fg='purple', command=init_db)
button50.grid(row=5, column=0)

# 开始按钮

button51 = Button(root, text='开始', width=20, font=('幼圆', 12), fg='purple', command=run)
button51.grid(row=5, column=1)

label60 = Label(root, text='  ', font=('楷体', 6), fg='black')
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

root.mainloop()

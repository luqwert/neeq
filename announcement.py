#!/usr/bin/env python
# _*_ coding:utf-8 _*_
# @Author  : lusheng
import datetime
import sqlite3
import json

import requests
import re
import time


# 创建数据库新闻表
def create_db():
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


# create_db()  # 初始化数据库
url = 'http://www.neeq.com.cn/disclosureInfoController/infoResult_zh.do?callback=jQuery331_1596699678177'
companyCd = input("请输入需要查询的公司代码或公司名称：")
if len(companyCd) == 6 or len(companyCd) == 8:
    data = {
        "disclosureType[]": 5,
        "disclosureSubtype[]": None,
        "page": 0,
        "startTime": "2020-03-06",
        "endTime": "2020-08-06",
        "companyCd": companyCd,
        "isNewThree": 1,
        "keyword": None,
        "xxfcbj[]": None,
        "hyType[]": None,
        "needFields[]": ["companyCd", "companyName", "disclosureTitle", "destFilePath", "publishDate", "xxfcbj",
                         "destFilePath", "fileExt", "xxzrlx"],
        "sortfield": "xxssdq",
        "sorttype": "asc",
    }
    # 获取新闻列表
    news_list = []
    response1 = requests.post(url, data)
    # print(response1.text)
    response2 = re.search('(?<=\(\[)(.*?)(?=]\))', response1.text).group()
    # print(response2)
    j = json.loads(response2)['listInfo']
    if j['content'] == []:
        print("没有查询到信息，请检查公司代码或名称是否正确")
    else:
        totalElements = j['totalElements']
        totalPages = j['totalPages']
        print("查询到%d条公告，共%d页" % (totalElements, totalPages))
        db = sqlite3.connect("announcement.db")
        c = db.cursor()
        list = j['content']
        for li in list:
            print(li)
            companyCd = li['companyCd']
            companyName = li['companyName']
            destFilePath = "http://www.neeq.com.cn" + li['destFilePath']
            disclosureTitle = li['disclosureTitle']
            publishDate = li['publishDate']
            xxfcbj = li['xxfcbj']
            xxzrlx = li['xxzrlx']
            result = c.execute("SELECT * FROM announcement where filePath = '%s'" % destFilePath)
            if result.fetchone():
                print(disclosureTitle, " 该公告数据库中已存在\n")
            else:
                data = "NULL,\'%s\',\'%s\',\'%s\',\'%s\',\'%s\'" % (
                    companyCd, companyName, disclosureTitle, publishDate, destFilePath)
                # print(data, "\n")
                c.execute('INSERT INTO announcement VALUES (%s)' % data)
                db.commit()
        time.sleep(3)  # 获取一个页面后休息3秒，防止请求服务器过快
else:
    print("公司代码或名称格式错误，请检查")


# try:
#     while True:
#         # create_db()
#         db = sqlite3.connect("news.db")
#         c = db.cursor()
#         newspage(5)  # 每次获取5页内容
#         time.sleep(600)  # 10分钟运行一次
# except Exception as e:
#     print(str(e))
# finally:
#     db.close()

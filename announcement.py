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
    conn = sqlite3.connect("news.db")
    c = conn.cursor()

    # 创建表
    c.execute('''DROP TABLE IF EXISTS news_title ''')  # 删除旧表，如果存在（因为这是临时数据）
    c.execute(
        '''CREATE TABLE news_title (news_id INTEGER PRIMARY KEY AUTOINCREMENT, title text only ,date text , link text)''')
    conn.commit()
    conn.close()


url = ''
# 获取新闻列表
news_list = []
response1 = requests.request('get', url)
# print(response1.text)
response2 = re.search('(?<=\()(.*?)(?=\);}catch\(e\))', response1.text).group()
# print(response2)
j = json.loads(response2)
list = j['result']['data']
for news in reversed(list):
    time_stamp = int(news['ctime'])
    time_array = time.localtime(time_stamp)
    str_date = time.strftime("%Y-%m-%d %H:%M:%S", time_array)
    # print(news['title'], news['url'], str_date)
    news_list.append([news['title'], str_date, news['url']])

# def newspage(page):
#     for n in reversed(range(page)):
#         url = "https://feed.mix.sina.com.cn/api/roll/get?pageid=153&lid=2516&etime=NaN&stime=NaN&ctime=NaN&date=day&k=&num=50&page=%d&r=0.9980420077788352&callback=jQuery111205439543678332086_1595820470788&_=1595820470790" % (
#                 n + 1)
#         lists = get_news_list(url)
#         for li in lists:
#             result = c.execute("SELECT * FROM news_title where title = '%s'" % li[0])
#             if result.fetchone():
#                 print(li, "该新闻数据库中已存在\n")
#             else:
#                 data = "NULL,\'%s\',\'%s\',\'%s\'" % (li[0], li[2], li[1])
#                 print(data, "\n")
#                 c.execute('INSERT INTO news_title VALUES (%s)' % data)
#                 db.commit()
#         time.sleep(3)  # 获取一个页面后休息3秒，防止请求服务器过快


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

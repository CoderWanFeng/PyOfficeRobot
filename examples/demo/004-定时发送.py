# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@读者群     ：http://www.python4office.cn/wechat-group/
@学习网站      ：https://www.python-office.com
@Date    ：2023/2/13 21:19
@本段代码的视频说明     ：https://www.bilibili.com/video/BV1m8411b7LZ
'''
import PyOfficeRobot

# 定时发送文字信息
# PyOfficeRobot.chat.send_message_by_time(who='Yaaakaaang', message='这是定时自动发送的信息', time='15:51:00')

# 定时发送文件
PyOfficeRobot.file.send_file_by_time(who='Yaaakaaang', file=r'003-根据关键词回复.py', send_time='11:40:10')
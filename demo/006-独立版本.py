# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信     ：CoderWanFeng : https://mp.weixin.qq.com/s/Nt8E8vC-ZsoN1McTOYbY2g
@个人网站      ：www.python-office.com
@代码日期    ：2023/7/9 23:25 
@本段代码的视频说明     ：https://www.bilibili.com/video/BV1SY411y7Uh
'''

# 原始方式
import office

office.wechat.send_message(who='百度一下：程序员晚枫', message='点个star吧')

# 独立方式
import PyOfficeRobot

PyOfficeRobot.chat.send_message(who='百度一下：程序员晚枫', message='点个star吧')

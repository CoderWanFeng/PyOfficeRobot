# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信     ：CoderWanFeng : https://mp.weixin.qq.com/s/Nt8E8vC-ZsoN1McTOYbY2g
@个人网站      ：www.python-office.com
@代码日期    ：2023/8/9 23:05 
@本段代码的视频说明     ：
'''
import PyOfficeRobot as pr

pr.chat.receive_message(who='每天进步一点点', txt='userMessage.txt', output_path='./')

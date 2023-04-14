# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/4/2 21:59 
@Description     ：
'''

from PyOfficeRobot.core.WeChatType import WeChat

wx = WeChat()

who = '文件传输助手'
wx.GetSessionList()  # 获取会话列表
# wx.ChatWith(who)  # 打开`who`聊天窗口
while True:
    print(wx.GetAllMessage)
    friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
    print(friend_name, receive_msg)

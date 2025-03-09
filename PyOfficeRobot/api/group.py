# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@读者群     ：http://www.python4office.cn/wechat-group/
@个人网站 ：www.python-office.com
@Date    ：2023/4/27 21:50 
@Description     ：
'''

import sys
import time
import pandas as pd
from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction
from PySide6.QtWidgets import QApplication

from PyOfficeRobot.core.group.Start import MyWidget

from PyOfficeRobot.core.WeChatType import WeChat
wx = WeChat()

@instruction
def send():
    app = QApplication(sys.argv)
    # 初始化QApplication，界面展示要包含在QApplication初始化之后，结束之前
    window = MyWidget()
    window.show()
    # 初始化并展示我们的界面组件
    sys.exit(app.exec_())
    # 结束QApplication



def group_by_keywords(who: str, keywords: pd.DataFrame) -> None:
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    last_processed_msg = None
    while True:
        time.sleep(3)
        last_msg = wx.MsgList.GetChildren()[-1].Name
        if last_msg and last_msg != last_processed_msg:
            if last_msg not in keywords['回复内容'].values:
                matched_replies = keywords[keywords['关键词'].apply(lambda x: x in last_msg)]['回复内容']
                if not matched_replies.empty:
                    for reply in matched_replies:
                        wx.SendMsg(reply, who)
                    last_processed_msg = last_msg
                else:
                    print("没有匹配的关键字")
            else:
                print("最后一条消息是自动回复内容，跳过回复")




if __name__ == "__main__":
    data = {
            '序号': ['0', '1', '2'],
            '关键词': ["报名",  "学习",  "课程"],
            '回复内容': ["你好，这是报名链接：www.这是报名链接.com",  "你好，这是学习链接：www.这是学习链接.com", "你好，这是课程链接：www.这是课程链接.com"]
            }
    keywords = pd.DataFrame(data)
    group_by_keywords(who='测试群', keywords=keywords)
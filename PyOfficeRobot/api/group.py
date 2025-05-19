# -*- coding: UTF-8 -*-


import sys
import time

from PySide6.QtWidgets import QApplication

from PyOfficeRobot.core.WeChatType import WeChat
from PyOfficeRobot.core.group.Start import MyWidget
from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction

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


def chat_by_keywords(who: str, keywords: dict, match_type: str) -> None:
    """
    根据关键词进行聊天回复。

    Args:
        who (str): 要与之聊天的群名。
        keywords (dict): 匹配关键字的字典，键为关键词，值为回复内容。
        match_type (str): 匹配类型，'exact' 表示完全匹配，'contains' 表示包含匹配。

    Returns:
        None

    """

    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    last_processed_msg = None
    while True:
        time.sleep(3)
        last_msg = wx.MsgList.GetChildren()[-1].Name
        if last_msg and last_msg != last_processed_msg:
            if last_msg not in keywords.values():
                matched_reply = None
                if match_type == 'exact':
                    matched_reply = keywords.get(last_msg)
                elif match_type == 'contains':
                    matched_reply = next((value for key, value in keywords.items() if key in last_msg), None)
                if matched_reply:
                    wx.SendMsg(matched_reply, who)
                    last_processed_msg = last_msg
                else:
                    print(f"没有匹配的关键字,匹配类型：{match_type}")
            else:
                print("最后一条消息是自动回复内容，跳过回复")
        else:
            print("没有新消息")


def collect_msg():
    pass


if __name__ == "__main__":
    who = '测试群'
    keywords = {
        "报名": "你好，这是报名链接：www.python-office.com",
        "学习": "你好，这是学习链接：www.python-office.com",
        "课程": "你好，这是课程链接：www.python-office.com"
    }
    match_type = 'contains'  # 关键字匹配类型 包含：contains  精确：exact
    chat_by_keywords(who=who, keywords=keywords, match_type=match_type)

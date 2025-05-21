# -*- coding: UTF-8 -*-


import sys
import time
import pandas as pd
import win32api
import win32con

from pathlib import Path
from loguru import logger
from pofile import mkdir
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


def collect_msg(who='文件传输助手', output_excel_path='userMessage.txt', output_path='./',names=None, scroll_num=6):

    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口

    # 获取当前鼠标位置
    x, y = win32api.GetCursorPos()

    # 向右移动 250 像素
    win32api.SetCursorPos((x + 250, y))

    # 向上滚动鼠标 scroll_num 次
    for i in range(scroll_num):
        # 向上滚动鼠标
        scroll_up()
        print(f"向上滚动鼠标{i + 1}次")
        time.sleep(0.5)

    num = 0
    messages_list = []
    index = 0
    receive_msgs = wx.GetAllMessage  # 获取朋友的名字、发送的信息、ID
    # print(f"---------{receive_msgs}")

    while True:
        if 0 - num == len(receive_msgs):
            print('没有更多消息了')
            # print(num)
            break
        num -= 1
        friend_name = receive_msgs[num][0]
        receive_msg = receive_msgs[num][1]

        # 如果是系统信息，则跳过
        if friend_name == 'SYS':
            index += 1
            continue

        elif friend_name == 'Time':  # ('Time', '10:53', '4226902524247')
            index += 1
            if len(receive_msg.split(' ')) > 1:
                today = receive_msg.split(' ')[0]
                timestamp = receive_msg.split(' ')[1]
            else:
                today = None
                timestamp = receive_msg
            friend_name = None
            text = None
            image = None

        elif names is not None and friend_name not in names:
            continue

        else:
            index += 1
            today = None
            timestamp = None
            if receive_msg == '[图片]':
                text = None
                image = receive_msg
            else:
                text = receive_msg
                image = None

        data = {
            "序号": index,
            "日期": today,
            "时间": timestamp,
            "好友": friend_name,
            "文字聊天内容":text,
            "图片聊天内容":image,
            "备注": None,
        }
        # print(data)
        messages_list.append(data)
    # print(messages_list)

    mkdir(Path(output_path).absolute())  # 如果不存在，则创建输出目录
    if output_excel_path.endswith('.xlsx') or output_excel_path.endswith('xls'):  # 如果指定的输出excel结尾不正确，则报错退出
        abs_output_excel = Path(output_path).absolute() / output_excel_path
    else:  # 指定了，但不是xlsx或者xls结束
        raise BaseException(
            f'输出结果名：output_excel参数，必须以xls或者xlsx结尾，您的输入:{output_excel_path}有误，请修改后重新运行')

    df = pd.DataFrame(messages_list)
    df.to_excel(str(abs_output_excel), index=False, engine='openpyxl')
    logger.info(f'识别结果已保存到：{output_path}')

def scroll_up(amount=1):
    for _ in range(amount):
        win32api.mouse_event(win32con.MOUSEEVENTF_WHEEL, 0, 0, 500)  # 500 是滚轮“一个刻度”的值
        time.sleep(0.1)  # 可选，增加间隔让滚动生效


if __name__ == "__main__":
    # who = '测试群'
    # keywords = {
    #     "报名": "你好，这是报名链接：www.python-office.com",
    #     "学习": "你好，这是学习链接：www.python-office.com",
    #     "课程": "你好，这是课程链接：www.python-office.com"
    # }
    # match_type = 'contains'  # 关键字匹配类型 包含：contains  精确：exact
    # chat_by_keywords(who=who, keywords=keywords, match_type=match_type)
    collect_msg(who='✨python-office开源小组', output_excel_path='userMessage.xlsx', output_path='./', names=['程序员晚枫'], scroll_num=6)

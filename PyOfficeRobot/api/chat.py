import datetime
import os

import poai
import porobot as porobot
import schedule

from PyOfficeRobot.core.WeChatType import WeChat
from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction

wx = WeChat()


# @act_info(ACT_TYPE.MESSAGE)
@instruction
def send_message(who: str, message: str) -> None:
    """
    给指定人，发送一条消息
    :param who:
    :param message:
    :return:
    """

    # 获取会话列表
    wx.GetSessionList()
    wx.ChatWith(who)  # 打开`who`聊天窗口
    # for i in range(10):
    wx.SendMsg(message, who)  # 向`who`发送消息：你好~


@instruction
def chat_by_keywords(who: str, keywords: str):
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    temp_msg = ''
    while True:
        try:
            friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
            if (friend_name == who) & (receive_msg != temp_msg) & (receive_msg in keywords.keys()):
                """
                条件：
                朋友名字正确:(friend_name == who)
                不是上次的对话:(receive_msg != temp_msg)
                对方内容在自己的预设里:(receive_msg in kv.keys())
                """

                temp_msg = receive_msg
                wx.SendMsg(keywords[receive_msg], who)  # 向`who`发送消息
        except:
            pass


@instruction
def receive_message(who='文件传输助手', txt='userMessage.txt', output_path='./'):
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    while True:
        friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
        current_time = datetime.datetime.now()
        cut_line = '^^^----------^^^'
        print('--' * 88)
        with open(os.path.join(output_path, txt), 'a+') as output_file:
            output_file.write('\n')
            output_file.write(cut_line)
            output_file.write('\n')
            output_file.write(str(current_time))
            output_file.write('\n')
            output_file.write(str(friend_name))
            output_file.write('\n')
            output_file.write(str(receive_msg))
            output_file.write('\n')


@instruction
def send_message_by_time(who, message, time):
    schedule.every().day.at(time).do(send_message, who=who, message=message)
    while True:
        schedule.run_pending()


@instruction
def chat_by_gpt(who, api_key, model_engine="text-davinci-002", max_tokens=1024, n=1, stop=None, temperature=0.5,
                top_p=1,
                frequency_penalty=0.0, presence_penalty=0.6):
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    temp_msg = None
    while True:
        try:
            friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
            if (friend_name == who) & (receive_msg != temp_msg):
                """
                条件：
                朋友名字正确:(friend_name == who)
                不是上次的对话:(receive_msg != temp_msg)
                对方内容在自己的预设里:(receive_msg in kv.keys())
                """
                print(f'【{who}】发送：【{receive_msg}】')
                temp_msg = receive_msg
                reply_msg = poai.chatgpt.chat(api_key, receive_msg,
                                              model_engine, max_tokens, n, stop, temperature,
                                              top_p, frequency_penalty, presence_penalty)
                wx.SendMsg(reply_msg, who)  # 向`who`发送消息
        except:
            pass


@instruction
def chat_robot(who):
    wx.GetSessionList()  # 获取会话列表
    wx.ChatWith(who)  # 打开`who`聊天窗口
    temp_msg = None
    while True:
        try:
            friend_name, receive_msg = wx.GetAllMessage[-1][0], wx.GetAllMessage[-1][1]  # 获取朋友的名字、发送的信息
            if (friend_name == who) & (receive_msg != temp_msg):
                """
                条件：
                朋友名字正确:(friend_name == who)
                不是上次的对话:(receive_msg != temp_msg)
                对方内容在自己的预设里:(receive_msg in kv.keys())
                """
                print(f'【{who}】发送：【{receive_msg}】')
                temp_msg = receive_msg
                reply_msg = porobot.normal.chat(receive_msg)
                wx.SendMsg(reply_msg, who)  # 向`who`发送消息
        except:
            pass

import schedule

from PyOfficeRobot.core.WeChatType import WeChat
import datetime
import os

from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction

wx = WeChat()

# @act_info(ACT_TYPE.MESSAGE)
@instruction
def send_message(who, message):
    """
    给指定人，发送一条消息
    :param who:
    :param message:
    :return:
    """

    # 获取会话列表
    wx.GetSessionList()
    wx.ChatWith(who)  # 打开`文件传输助手`聊天窗口
    # for i in range(10):
    wx.SendMsg(message)  # 向`文件传输助手`发送消息：你好~

@instruction
def chat_by_keywords(who, keywords):
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
                wx.SendMsg(keywords[receive_msg])  # 向`who`发送消息
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

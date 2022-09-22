from PyOfficeRobot.core.WeChatType import WeChat
from PyOfficeRobot.lib.CONST import ACT_TYPE
from PyOfficeRobot.lib.dec.act_dec import act_info

wx = WeChat()


@act_info(ACT_TYPE.MESSAGE)
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

from PyOfficeRobot.core.WeChatType import *

wx = WeChat()  # 获取当前微信客户端
p = wx.GetSessionList()  # 获取会话列表
msg = '###'
who = '文件传输助手'
wx.ChatWith(who)  # 确定聊天对象
temp_msg, temp_id = 0, 0
while True:
    try:
        receive_name, receive_msg, receive_id = wx.GetLastMessage[0], wx.GetLastMessage[1], wx.GetLastMessage[2]
        if receive_msg != temp_msg and receive_id != temp_id:
            print(f'消息来自：{receive_name}, 消息内容：{receive_msg}, 消息编号：{receive_id}')
            temp_id = receive_id
            temp_msg = receive_msg
            # wx.SendMsg(msg)  # 向聊天对象who发送msg(消息)
        else:
            pass
    except Exception as ex:
        print(ex)

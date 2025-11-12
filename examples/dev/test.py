from wxauto4 import WeChat   # 开源版

# 初始化微信实例
wx = WeChat()

# 发送消息
wx.SendMsg("你好", who="程序员晚枫")

# 获取当前聊天窗口消息
msgs = wx.GetAllMessage()

for msg in msgs:
    print('==' * 30)
    print(f"{msg.sender}: {msg.content}")
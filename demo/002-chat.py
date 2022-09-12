"""
作者微信:https://mp.weixin.qq.com/s/dAm2B09i2ZaqCwhwP-AEdQ
原文链接：https://mp.weixin.qq.com/s/jUsWXJl5xDFueTESmc5CYg
"""

# 首先，将PyOfficeRobot模块导入到我们的代码块中。
from PyOfficeRobot.core.WeChatType import *

# 获取当前微信客户端
wx = WeChat()
# 获取会话列表
chat_list = wx.GetSessionList()
# # 向某人发送消息（以`文件传输助手`为例）
# msg = '你好~我是B站：程序员晚枫，感谢点赞 + 转发'
# who = '文件传输助手'
# wx.ChatWith(who)  # 打开`文件传输助手`聊天窗口
# # for i in range(10):
# wx.SendMsg(msg)  # 向`文件传输助手`发送消息：你好~


wx.ChatWith(who='测试群')
wx.SavePic()

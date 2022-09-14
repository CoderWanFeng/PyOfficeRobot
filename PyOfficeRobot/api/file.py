from PyOfficeRobot.core.WeChatType import WeChat

wx = WeChat()


def send_file(who, file):
    """
    发送任意类型的文件
    :param who:
    :param file: 文件路径
    :return:
    """
    # 向某人发送文件（以`文件传输助手`为例，发送三个不同类型文件）
    wx.ChatWith(who)  # 打开`文件传输助手`聊天窗口
    wx.SendFiles(file)  # 向`文件传输助手`发送上述三个文件
    # 注：为保证发送文件稳定性，首次发送文件可能花费时间较长，后续调用会缩短发送时间

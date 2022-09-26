from PR.Dumogu.WeChatType import *

def send_msg(msg):
    """
    直接发送消息，对核心文件修正后的封装，可以发送换行信息
    :param msg:
    """
    global wx
    if '\n' in msg:
        for k in msg.split('\n'):
            wx.SendMsg(k, clear=False)
    else:
        wx.SendMsg(msg)
    wx.SendEnd()  # 最终完成信息发送


def send_file(filename):
    """
    直接发送消文件，对核心文件修正后的封装，可以发送换行信息
    :param filename:
    """
    global wx
    wx.SendFiles(filename)
    wx.SendEnd()  # 最终完成信息发送


def send(msg, name):
    """
    确定聊天对象后发送消息，一般不使用~
    :param name: 发送对象或者对象列表
    :param msg: 发送的信息
    """
    global wx
    if type(name) == list:
        for n in name:
            wx.ChatWith(n)
    else:
        wx.ChatWith(name)
    send_msg(msg)


def main_reply(msg, chats):
    """主要回复函数，功能需要自己慢慢加，天气查询，机器回复的设定，诗词查询等可以自己封装，在try语句内"""
    global wx, bot_list, temp_msg, wind
    xinxi = msg.strip()
    admin = 1 if chats in bot_list else 0

    if "阿雨说" in msg:  # 阿雨说
        png(sp(xinxi, "阿雨说"), './Send/baojie.png')
        send_file(base_pan + "/Send/new.png")

    elif '拜年' == xinxi:  # 拜年
        s_xinxi = bainian()
        send(s_xinxi, ['文件传输助手', '小熊'])

    elif not listout(['也祝您', '也祝你', '祝亚飞', '和家人', '及家人'], xinxi):  # 拜年自动回复
        time.sleep(3)
        send_msg(f"谢谢~")

    elif "天气" in xinxi[:6]:  # 天气查询
        send_msg(tian_qi(xinxi))

    elif "关机" in xinxi:  # 关机控制
        send_msg(shutdown(xinxi))

    elif "打卡" == xinxi:  # 打卡码发送
        send_file(base_pan + "/Send/打卡.png")

    elif "核酸" in xinxi and len(xinxi) < 7 and admin:  # 核酸检测码的发送
        if "思雨" in xinxi:
            if "周口" in xinxi:
                send_file(base_pan + "/Send/syzkhs.png")
            elif "郑州" in xinxi:
                send_file(base_pan + "/Send/syzzhs.png")
        else:
            if "周口" in xinxi:
                send_file(base_pan + "/Send/zkhs.png")
            elif "郑州" in xinxi:
                send_file(base_pan + "/Send/zzhs.png")
    elif "日历" in xinxi:
        xinxi = sp(msg.text, "日历")
        year, month = int(xinxi[0]), int(xinxi[1])
        cal(year, month)

    elif "诗词" in xinxi:  # 诗词
        xinxi = sp(xinxi, "诗词")
        send_msg(shi(xinxi))

    elif "夸姐姐" == xinxi:  # 夸姐姐
        s_xinxi = kua()
        send_msg(s_xinxi + '[BOT]')

    elif "截图" == xinxi and admin:
        s_xinxi = screen()
        send_file(s_xinxi)

    elif "@@" in xinxi and admin:
        s_xinxi = sp(xinxi, '@@')
        if s_xinxi == '':
            s_xinxi = '文件助手'
        send(f'Chat with {s_xinxi}！/偷笑', '文件助手')
        wx.ChatWith(s_xinxi)
        if not wind:
            wx.MinWind.Click()

    elif "发票识别" == xinxi and admin:
        send(f'{ocr_main()}', '文件助手')

    elif "@窗口" in xinxi and admin:
        wind = 0 if '关闭' in xinxi else 1
        send(f'窗口 {"打开" if wind else "关闭"}！/偷笑', '文件助手')

    elif "学习" == xinxi[:2]:  # 小熊机器人学习模块儿
        s_xinxi = data.putin(sp(xinxi, "学习"))
        send_msg(s_xinxi + '[BOT]')
    elif "遗忘" == xinxi[:2]:
        s_xinxi = data.delete(sp(xinxi, "遗忘"))
        send_msg(s_xinxi + '[BOT]')
    elif "问题" == xinxi and admin:
        love = ' | '.join(list(data.sp_question[0].keys()))
        lang = ' | '.join(list(data.sp_question[1].keys()))
        s_xinxi = f'{love}\n{lang}' if admin else lang
        send_msg(f'{s_xinxi}[BOT]')

    elif "升级" == xinxi[:2]:  # 数据升级
        xin = sp(xinxi, "升级")
        if "+" in xin:
            q, a = xin.split('+')[0], xin.split('+')[1]
        else:
            q, a = xin, ''
        s_xinxi = data.update(q, a)
        send_msg(s_xinxi)
    elif data.all(xinxi):  # 数据库回复
        s_xinxi = data.all(xinxi) if admin else data.lang(xinxi)
        send_msg(s_xinxi)


def check_msg(temp, msg, t):
    """
    检查 temp[t] 和 msg[t] 是否一致
    :rtype: boll
    """
    if t in temp:
        if temp[t] == msg[t]:
            return False
        else:
            return True
    else:
        return True


wx = WeChat()  # 获取当前微信客户端
wx.ChatWith('文件助手')
wind = 1
temp_msg, msg_dic, temp_dic = '', {}, {}  # 存储最后一条消息，最后一条消息字典，存储所有消息字典
bot_list = ['小熊', '阿雨', '']  # admin的权限人员，可以远程调整pdbot
data = BRAIN(f'I:/EXE/WX/Log/Data_brain.db', 'question,answer,type')
while True:
    try:
        receive_msg, chat = wx.GetLastMessage
        msg_dic[chat] = receive_msg
        chat = '小熊' if not chat else chat
        if receive_msg != temp_msg:
            if chat == '小熊':
                if check_msg(temp_dic, msg_dic, chat):
                    prt(f'时间：{rtime()}, {chat}--> 消息内容：{receive_msg}')
                    temp_dic[chat] = receive_msg
            else:
                if check_msg(temp_dic, msg_dic, chat):
                    prt(f'时间：{rtime()}, From {chat}--> 消息内容：{receive_msg}', zi='紫')
                    temp_dic[chat] = receive_msg
            try:
                main_reply(receive_msg, chat)  # 向聊天对象反馈消息
            except Exception as ex:
                print(traceback.format_exc(ex))
    except IndexError:
        pass
    except TypeError:
        pass

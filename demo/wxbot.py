from PR.Dumogu.WeChatType import *

# pip install -i https://pypi.tuna.tsinghua.edu.cn/simple python-office -U
import subprocess
from Panda.Base import *
from  PR.Dumogu.Pdbot import *
# from Android import *

def send_msg(msg, name):
    """直接发送消息，对核心文件修正后的封装，可以发送换行信息"""
    if name:
        wx.ChatWith(name)
    for m in msg.split('\n'):
        os.system(f'echo {m} | clip')
        wx.SendMsg('{Ctrl}v', False)
    wx.SendEnd()

def send_file(filename, name):
    real_path = os.path.realpath(filename)
    if name:
        wx.ChatWith(name)
    subprocess.Popen(args=['powershell', f'Get-Item {real_path} | Set-Clipboard'])
    wx.SendMsg('{Ctrl}v')
    wx.SendEnd()

def main_reply(msg, chat, talk):
    """主要回复函数，功能需要自己慢慢加，天气查询，机器回复的设定，诗词查询等可以自己封装，在try语句内不会报错"""
    global wx, admin_list, temp_msg, wind
    logging.info(msg)
    xinxi = msg.strip()
    admin = 1 if chat in admin_list else 0

    if "阿雨说" in msg:  # 阿雨说
        png(sp(xinxi, "阿雨说"), './Send/baojie.png')
        send_file(base_pan + "/Send/new.png", talk)

    elif "考勤打卡" == msg:
        send_msg("准备开始打卡...", talk)
        dk_dd()
        dk_ykt()

    elif "聊天记录" in msg and admin:
        if '@' in msg:
            wx.ChatWith(msg.split('@')[1])
        send_msg(show_all_msg(), talk)

    elif "天气" in xinxi[:6]:  # 天气查询
        path = tian_qi(xinxi)
        send_file(path, talk)

    elif "关机" in xinxi:  # 关机控制
        send_msg(shutdown(xinxi), talk)

    elif "打卡" == xinxi:  # 打卡码发送
        path = base_pan + "/Send/打卡.png"
        send_file(path, talk)

    elif "核酸" in xinxi and len(xinxi) < 7 and admin:  # 核酸检测码的发送
        if "思雨" in xinxi:
            path = base_pan + "/Send/syzzhs.png" if "郑州" in xinxi else base_pan + "/Send/syzkhs.png"
        else:
            path = base_pan + "/Send/yfzzhs.png" if "郑州" in xinxi else base_pan + "/Send/yfzkhs.png"
        send_file(path, talk)

    elif "诗词" in xinxi:  # 诗词
        xinxi = sp(xinxi, "诗词")
        send_msg(shici(xinxi), talk)

    elif "截图" == xinxi and admin:
        path = pc_screen()
        send_file(path, talk)

    elif "色图" == xinxi or "涩图" == xinxi and admin and "群" not in talk:
        path = setu()
        send_file(path, talk)

    elif "@@" in xinxi and admin:
        message = sp(xinxi, '@@')
        if message == '':
            message = '文件助手'
        send_msg(f'Chat with {message}！/偷笑', '小熊')
        wx.ChatWith(message)

    elif "学习" == xinxi[:2] and admin:  # 小熊机器人学习模块儿
        message = data.putin(sp(xinxi, "学习"))
        send_msg(message + '[BOT]', talk)
    elif "遗忘" == xinxi[:2] and admin:
        message = data.delete(sp(xinxi, "遗忘"))
        send_msg(message + '[BOT]', talk)
    elif "问题" == xinxi and admin:
        love = ' - '.join(list(data.sp_question[0].keys()))
        lang = ' - '.join(list(data.sp_question[1].keys()))
        message = f'{love}\n{lang}' if admin else lang
        send_msg(f'{message}[BOT]', talk)

    elif "升级" == xinxi[:2] and admin:  # 数据升级
        xin = sp(xinxi, "升级")
        ques, ans = xin.split('+')[0], xin.split('+')[1] if "+" in xin else xin, ''
        message = data.update(ques, ans)
        send_msg(message, talk)
    elif data.all(xinxi):  # 数据库回复
        message = data.all(xinxi) if admin else data.lang(xinxi)
        send_msg(message, talk)

def check_msg(temp, msg: dict, t) -> bool:
    """ 检查 temp[t] 和 msg[t] 是否一致"""
    if t in temp:
        msg_check = False if temp[t] == msg[t] else True
    else:
        msg_check = True
    return msg_check

def nick(name):
    nick_list = ['总', '部长', '主任', '老师', '姐', '哥', '兄', '伯', '叔', '处', '班长', '书记', '兄弟', '姨', '嫂', '娘', '行长', '律师', '经理']
    for n in nick_list:
        if n in name:
            return name[:1] + n
    return name

def show_all_msg():
    msgs = wx.GetAllMessage
    txt = []
    for msg in msgs:
        if ':' in msg[0] and len(msg[0]) == 5:
            line = '*' * 12 + msg[0] + '*' * 12
            prt('*' * 30 + msg[0] + '*' * 30, zi='蓝')
        elif msg[1] == msg[2]:
            line = msg[1] + '：' + msg[0]
            prt(msg[1] + '-->' + msg[0], zi='红')
        else:
            line = msg[1] + '：' + msg[0]
            prt('\t\t' + msg[1] + '-->' + msg[0], zi='蓝')
        txt.append(line)
    return "\n".join(txt)

def main():
    while True:
        try:
            time.sleep(0.5)
            receive_msg, chat, talk = wx.GetLastMessage
            chat = self if not chat else chat
            msg_dic[chat] = receive_msg
            if receive_msg != temp_msg and receive_msg != "停止":
                logging.info(f'{temp_dic},{msg_dic},{chat}-->{check_msg(temp_dic, msg_dic, chat)}')
                # print(receive_msg, chat, talk)
                if chat == self:
                    if check_msg(temp_dic, msg_dic, chat):
                        prt(f'时间：{time_str()}{chat}--> 消息内容：{receive_msg}', zi='蓝')
                        temp_dic[chat] = receive_msg
                elif chat != self and chat == talk:
                    if check_msg(temp_dic, msg_dic, chat):
                        prt(f'时间：{time_str()}From {chat}--> 消息内容：{receive_msg}', zi='红')
                        temp_dic[chat] = receive_msg
                elif chat != self and chat != talk:
                    if check_msg(temp_dic, msg_dic, chat):
                        prt(f'时间：{time_str()}From [{talk}]{chat}--> 消息内容：{receive_msg}', zi='黄')
                        temp_dic[chat] = receive_msg
                try:
                    main_reply(receive_msg, chat, talk)  # 向聊天对象反馈消息
                except Exception as ex:
                    print(traceback.format_exc(ex.args))
            elif receive_msg == "停止":
                time.sleep(600)
        except IndexError:
            pass
        except TypeError:
            pass

if __name__ == "__main__":
    logging.basicConfig(format='[%(asctime)s][%(levelname)s]:%(message)s', level=logging.ERROR)
    wx = WeChat()  # 获取当前微信客户端
    self = wx.SelfName
    admin_list = [self, '小熊', '大熊', '阿雨']  # admin的权限人员，可以远程调整pdbot
    data = BRAIN(f'D:/Panda/Data_brain.db', 'question,answer,type')
    temp_msg, msg_dic, temp_dic = '', {}, {}  # 存储最后一条消息，最后一条消息字典，存储所有消息字典
    main()

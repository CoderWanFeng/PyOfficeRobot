# -*-coding:utf-8-*-
# sys.stdout = io.TextIOWrapper(sys.stdout.buffer,encoding='gb18030')

import uiautomation as uia
import win32gui, win32con
import win32clipboard as wc
import time
import os

def SavePic(savepath=None, filename: str = None) -> object:
    Pic = uia.WindowControl(ClassName='ImagePreviewWnd', Name='图片查看')
    Pic.SendKeys('{Ctrl}s')
    SaveAs = Pic.WindowControl(ClassName='#32770', Name='另存为...')
    SaveAsEdit = SaveAs.EditControl(ClassName='Edit', Name='文件名:')
    SaveButton = Pic.ButtonControl(ClassName='Button', Name='保存(S)')
    PicName, Ex = os.path.splitext(SaveAsEdit.GetValuePattern().Value)
    if not savepath:
        savepath = os.getcwd()
    if not filename:
        filename = PicName
    FilePath = os.path.realpath(os.path.join(savepath, filename + Ex))
    SaveAsEdit.SendKeys(FilePath)
    SaveButton.Click()
    Pic.SendKeys('{Esc}')


class WeChat:
    def __init__(self):
        # 主要窗体
        uia.PaneControl(Name='任务栏').ButtonControl(Name='微信').Click()  # 启动微信
        self.sema = threading.BoundedSemaphore(1)
        self.UiaAPI = uia.WindowControl(ClassName='WeChatMainWndForPC')

        self.ChatList = self.UiaAPI.ListControl(Name='会话')
        self.EditMsg = self.UiaAPI.EditControl(Name='输入')
        self.SearchBox = self.UiaAPI.EditControl(Name='搜索')
        self.MsgList = self.UiaAPI.ListControl(Name='消息')
        self.SelfName = self.UiaAPI.ButtonControl(searchDepth=4).Name
        # 必须处于界面可视化方可自动化
        # self.MaxWind = self.UiaAPI.ButtonControl(Name='最大化')
        # self.MinWind = self.UiaAPI.ButtonControl(Name='最小化')
        self.Chats = self.UiaAPI.ButtonControl(Name='聊天')
        self.Tongxunlu = self.UiaAPI.ButtonControl(Name='通讯录')

    def GetContacts(self, end_name, rang=2, times=1000, wheels=4, wait=0):
        contacts = []
        no_list = ['', '新的朋友', '公众号', '微信团队', '文件助手']
        self.UiaAPI.SwitchToThisWindow()  # 找到微信窗口
        self.Tongxunlu.Click()
        contact_list = self.UiaAPI.ListControl(Name='联系人')
        item = contact_list.ListItemControl()
        for x in range(rang):
            for i in range(times):
                try:
                    name = item.Name  # 返回会话列表名称
                    if (name not in contacts) and (name not in no_list):
                        contacts.append(name)
                    elif name == end_name:
                        break
                except Exception as ex:
                    contact_list.WheelDown(wheelTimes=wheels, waitTime=wait)
                    item = contact_list.GetFirstChildControl()
                    None if 'Name' in ex.args[0] else None
                if item:
                    item = item.GetNextSiblingControl()
            for i in range(times):
                try:
                    name = item.Name  # 返回会话列表名称
                    if (name not in contacts) and (name not in no_list):
                        contacts.append(name)
                    elif name == '新的朋友':
                        break
                except Exception as ex:
                    contact_list.WheelUp(wheelTimes=wheels + 1, waitTime=wait)
                    item = contact_list.GetLastChildControl()
                    None if 'Name' in ex.args[0] else None
                if item:
                    item = item.GetPreviousSiblingControl()
        self.Chats.Click()
        return contacts

    def GetChatList(self):
        """获取当前会话列表，更新会话列表"""
        item = self.ChatList.ListItemControl()
        temp_chat_list = []
        for i in range(20):
            try:
                name = item.Name  # 返回会话列表名称
            except:
                break
            if name not in temp_chat_list:
                temp_chat_list.append(name)
            item = item.GetNextSiblingControl()  # ???
        return temp_chat_list

    def Search(self, keyword):
        """查找微信好友或关键词"""
        self.UiaAPI.SetFocus()  # 找到微信窗口
        time.sleep(0.5)
        self.UiaAPI.SendKeys('{Ctrl}f{Ctrl}a', waitTime=0.5)
        self.SearchBox.SendKeys(keyword, waitTime=0.5)
        self.SearchBox.SendKeys('{Enter}')

    def ChatWith(self, who):
        """打开某个聊天框
        who : 要打开的聊天框好友名，str;"""
        self.UiaAPI.SwitchToThisWindow()
        if who in self.GetChatList()[:-1]:
            self.ChatList.ListItemControl(Name=who).Click(simulateMove=False)
        else:
            self.Search(who)
            self.ChatList.ListItemControl().GetFirstChildControl().Click(simulateMove=False)

    def SendMsg(self, msg, clear=False):
        """向当前窗口发送消息
        msg : 要发送的消息
        clear : 是否清除当前已编辑内容
        基本已弃用，主要是传递快捷键操作"""
        self.UiaAPI.SwitchToThisWindow()
        if clear:
            self.EditMsg.SendKeys('{Ctrl}a', waitTime=0)
        self.EditMsg.SendKeys(msg, waitTime=0)

    def SendEnd(self):
        self.UiaAPI.SwitchToThisWindow()
        self.EditMsg.SendKeys('{Enter}', waitTime=0)

    def split_msg(self, msgs):
        """分析消息列表，返回（消息，发消息人，聊天对象），聊天对象可能是群聊名字"""
        try:
            uia.SetGlobalSearchTimeout(0.5)
            if self.UiaAPI.PaneControl(searchDepth=1, foundIndex=2).Name != "":
                talk = self.UiaAPI.PaneControl(searchDepth=1, foundIndex=3).PaneControl(searchDepth=1, foundIndex=2)
            else:
                talk = self.UiaAPI.PaneControl(searchDepth=1, foundIndex=2).PaneControl(searchDepth=1, foundIndex=2)
            talk = talk.PaneControl(searchDepth=1, foundIndex=3).ButtonControl(RuntimeId='[2A.780632.4.15A]', foundIndex=1).Name
            uia.SetGlobalSearchTimeout(0.5)
        except Exception as ex:
            talk = '' if ex.args else ''
        chat = msgs.ButtonControl(searchDepth=2)  # 聊天对象应当从消息子目录获取

        if msgs.Name not in ['[图片]', '[文件]', '[音乐]', '[名片]']:
            msg = msgs.Name
            msg = "https://" + msg if 'www.' in msg[:4] and 'http' not in msg else msg
            try:
                chat = chat.Name
            except Exception as ex:
                chat = '系统' if ex.args else ''
        elif msgs.Name == '[名片]':
            card = msgs.TextControl()
            msg = f'{msgs.Name} {card.Name}'
            chat = chat.Name
        elif msgs.Name == '[文件]':
            chat = chat.Name
            if chat != self.SelfName:
                text = msgs.TextControl(foundIndex=2)
                size = msgs.TextControl(foundIndex=3)
            else:
                text = msgs.TextControl(foundIndex=1)
                size = msgs.TextControl(foundIndex=2)
                # print(text,size)
            if text:
                msg = f'{msgs.Name} ({size.Name}) {text.Name}'
        elif msgs.Name == '[音乐]':
            auth = msgs.TextControl(searchDepth=9)
            music = msgs.EditControl(searchDepth=9)
            if music:
                msg = f'{msgs.Name} {auth.Name}-->{music.Name}'
                chat = chat.Name
        elif "撤回了一条消息" in msgs.Name:
            chat = msgs.Name.replace("撤回了一条消息", '').replace("\"", '').strip()
            msg = f'[撤回消息]'
        elif msgs.Name == '[图片]':
            self.sema.acquire()
            try:
                SavePic()
            except Exception as ex:
                ex.args
            self.sema.release()
            msg = '[图片]'
            chat = chat.Name
        else:
            chat = chat.Name
            msg = msgs.EditControl().Name
        return msg, chat, talk  # 返回正确的聊天信息和聊天对象

    @property
    def GetAllMessage(self):
        """获取当前窗口中加载的所有聊天记录"""
        all_msg = []
        MsgItems = self.MsgList.GetChildren()[1:]
        count = 0
        lens = len(MsgItems)
        for msgs in MsgItems:
            count += 1
            print(f'\r全部聊天记录获取进度：[{count}/{lens}]...', flush=True, end='')
            all_msg.append(self.split_msg(msgs))
        print('')
        return all_msg

    @property
    def GetLastMessage(self):
        """获取当前窗口中最后一条聊天记录"""
        try:
            uia.SetGlobalSearchTimeout(0.5)
            msgs = self.MsgList.GetChildren()[-1]  # todo
            return self.split_msg(msgs)
        except LookupError:
            pass

    @property
    def LoadMoreMessage(self):
        """定位到当前聊天页面，并往上滚动鼠标滚轮，加载更多聊天记录到内存"""
        self.MsgList.WheelUp(wheelTimes=100, waitTime=0.1)

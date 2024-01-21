# -*- coding: UTF-8 -*-
'''
@作者  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫，微信：CoderWanFeng
@读者群     ：http://www.python4office.cn/wechat-group/
@学习网站      ：www.python-office.com
@代码日期    ：2023/11/9 23:00 
@本段代码的视频说明     ：
'''

#微信定时群发功能
#微信调用模拟类 start
import time

import win32api
import win32clipboard as w
import win32con
import win32gui

#微信调用模拟类 end

#调用HTTP插件类
import requests


# 把文字放入剪贴板
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_UNICODETEXT, aString)
    w.CloseClipboard()


# 模拟ctrl+V
def ctrlV():
    win32api.keybd_event(17, 0, 0, 0)  # 按下ctrl
    win32api.keybd_event(86, 0, 0, 0)  # 按下V
    win32api.keybd_event(86, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放V
    win32api.keybd_event(17, 0, win32con.KEYEVENTF_KEYUP, 0)  # 释放ctrl


# 模拟alt+s
def altS():
    win32api.keybd_event(18, 0, 0, 0)
    win32api.keybd_event(83, 0, 0, 0)
    win32api.keybd_event(83, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(18, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟enter
def enter():
    win32api.keybd_event(13, 0, 0, 0)
    win32api.keybd_event(13, 0, win32con.KEYEVENTF_KEYUP, 0)


# 模拟鼠标单击
def click():
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


# 移动鼠标的位置
def movePos(x, y):
    win32api.SetCursorPos((x, y))


if __name__ == "__main__":
    target_time = ['17:55', '22:45', '20:00']  # 这里是发送时间
    name_list = ['易']  # 这里是要发送信息的联系人 ['易', '天极']

    payload = dict(date='2023-11-09', store_name='万达二')
    r = requests.post('http://wd2.xwdgp.vip/public/admin/selectTodayDeclaration', data=payload)
    # print(r.text)
    #print(r.json()['copy_text'])
    #send_content = r.json()['copy_text'];
    # send_content = r.json()['copy_data']['beautician']
    send_content = ''
    next_line = '{ctrl}{ENTER}'
    for word in r.json()['copy_data']['beautician']:
        # print(word)是数组
        send_content += word + next_line

    #send_content = "这里是需要发送的信息内容"  # 这里是需要发送的信息内容
    while True:
        now = time.strftime("%m月%d日%H:%M", time.localtime())  # 返回格式化时间
        print(now)
        if now[-5:] in target_time:  # 判断时间是否为设定时间
            hwnd = win32gui.FindWindow("WeChatMainWndForPC", '微信')  # 返回微信窗口的句柄信息
            win32gui.ShowWindow(hwnd, win32con.SW_SHOW)  # 激活并显示微信窗口
            win32gui.MoveWindow(hwnd, 0, 0, 1000, 700, True)  # 将微信窗口移动到指定位置和大小
            time.sleep(1)
            for name in name_list:
                movePos(28, 147)
                click()
                movePos(148, 35)
                click()
                time.sleep(1)
                setText(name)
                ctrlV()
                time.sleep(1)  # 等待联系人搜索成功
                enter()
                time.sleep(1)
                setText(send_content)
                ctrlV()
                time.sleep(1)
                altS()
                time.sleep(1)
            win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
        time.sleep(60)
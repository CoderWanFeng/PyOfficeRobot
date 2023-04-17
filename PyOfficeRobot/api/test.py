# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/4/16 23:21 
@Description     ：
'''
from ctypes import *

import win32clipboard
import win32clipboard as wc

PUBLISH_ID = '公众号：程序员晚枫'

COPYDICT = {}


def SendFiles(self, filepath, not_exists='ignore'):
    """向当前聊天窗口发送文件
    not_exists: 如果未找到指定文件，继续或终止程序
    filepath: 要复制文件的绝对路径"""

    class DROPFILES(Structure):
        _fields_ = [
            ("pFiles", c_uint),
            ("x", c_long),
            ("y", c_long),
            ("fNC", c_int),
            ("fWide", c_bool),
        ]

    def setClipboardFiles(paths):
        files = ("\0".join(paths)).replace("/", "\\")
        data = files.encode("U16")[2:] + b"\0\0"
        wc.OpenClipboard()
        try:
            wc.EmptyClipboard()
            wc.SetClipboardData(
                wc.CF_HDROP, matedata + data)
        finally:
            wc.CloseClipboard()

    def readClipboardFilePaths():
        wc.OpenClipboard()
        try:
            return wc.GetClipboardData(win32clipboard.CF_HDROP)
        finally:
            wc.CloseClipboard()

    pDropFiles = DROPFILES()
    pDropFiles.pFiles = sizeof(DROPFILES)
    pDropFiles.fWide = True
    matedata = bytes(pDropFiles)
    filename = [r"%s" % filepath]
    setClipboardFiles(filename)

    # wc.CloseClipboard()
    self.SendClipboard()
    return 1

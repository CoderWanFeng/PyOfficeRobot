# -*- coding: UTF-8 -*-
'''
@Author  ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@WeChat     ：CoderWanFeng
@Blog      ：www.python-office.com
@Date    ：2023/4/27 21:50 
@Description     ：
'''

import sys

from PySide6.QtWidgets import QApplication

from PyOfficeRobot.core.group.Start import MyWidget


def send():
    app = QApplication(sys.argv)
    # 初始化QApplication，界面展示要包含在QApplication初始化之后，结束之前
    window = MyWidget()
    window.show()
    # 初始化并展示我们的界面组件
    sys.exit(app.exec_())
    # 结束QApplication

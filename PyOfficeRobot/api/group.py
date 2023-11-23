# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/Nt8E8vC-ZsoN1McTOYbY2g
@个人网站 ：www.python-office.com
@Date    ：2023/4/27 21:50 
@Description     ：
'''

import sys

from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction
from PySide6.QtWidgets import QApplication

from PyOfficeRobot.core.group.Start import MyWidget

@instruction
def send():
    app = QApplication(sys.argv)
    # 初始化QApplication，界面展示要包含在QApplication初始化之后，结束之前
    window = MyWidget()
    window.show()
    # 初始化并展示我们的界面组件
    sys.exit(app.exec_())
    # 结束QApplication

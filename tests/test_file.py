import unittest

from PyOfficeRobot.api import file

from PyOfficeRobot.api.chat import *
from PyOfficeRobot.api.file import *

keywords = {
    '你好': "在干嘛？"
}


class TestFile(unittest.TestCase):
    def test_send_file(self):
        send_file(who='文件传输助手', file=r'./test_files/0816.jpg')

    def test_chat_by_keyword(self):
        chat_by_keywords(who='程序员晚枫', keywords=keywords)

    def test_receive_message(self):
        receive_message(who='程序员晚枫')

    def test_sm_by_time(self):
        send_message_by_time(who='每天进步一点点', message='你好', time='17:38')

    def test_chat_by_gpt(self):
        chat_by_gpt(who='每天进步一点点')

    def test_weixin_file(self):
        who = '每天进步一点点'
        file_path = r'./chat.py'
        # file_path = r'D:\workplace\code\github\PyOfficeRobot\tests\chat.py'
        # PyOfficeRobot.file.send_file(who, file)
        file.send_file(who, file_path)
        # file.send_file(who='每天进步一点点', file=r'D:\workplace\code\github\PyOfficeRobot\dev\contributor\gen\自动加好友\Excel_File\【企查查】查企业-高级搜索“涂料”(202302030683).xls')

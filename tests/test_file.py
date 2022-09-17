import unittest

from PyOfficeRobot.api.chat import chat_by_keyword
from PyOfficeRobot.api.file import *

keywords = {
    '你好': "在干嘛？"
}


class TestFile(unittest.TestCase):
    def test_send_file(self):
        send_file(who='文件传输助手', file=r'./test_files/0816.jpg')

    def test_chat_by_keyword(self):
        chat_by_keyword(who='兆锋', keywords=keywords)

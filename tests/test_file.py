import unittest

from PyOfficeRobot.api.file import *


class TestFile(unittest.TestCase):
    def test_send_file(self):
        send_file(who='文件传输助手', file=r'./test_files/0816.jpg')

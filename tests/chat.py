import unittest

from PyOfficeRobot.api.chat import *


class TestFile(unittest.TestCase):

    def test_chat_ali(self):
        chat_ali(who='程序员晚枫', key=os.getenv('TY_KEY'))

    def test_chat_zhipu(self):
        chat_by_zhipu(who='程序员晚枫', key='e29c16C0HFKSpNHzIm')

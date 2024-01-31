import unittest

from PyOfficeRobot.api.chat import *


class TestFile(unittest.TestCase):

    def test_chat_ali(self):
        chat_ali(who='每天进步一点点', key=os.getenv('TY_KEY'))

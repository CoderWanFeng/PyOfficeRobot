import unittest

from PyOfficeRobot.api import file, group
from PyOfficeRobot.api.chat import *
from PyOfficeRobot.api.file import *

keywords = {
    '你好': "在干嘛？"
}


class TestFile(unittest.TestCase):
    def test_send_file(self):
        send_file(who='文件传输助手', file=r'../test_files/0816.jpg')

    def test_chat_by_keyword(self):
        chat_by_keywords(who='程序员晚枫', keywords=keywords)

    def test_receive_message(self):
        receive_message(who='程序员晚枫')

    def test_sm_by_time(self):
        send_message_by_time(who='程序员晚枫', message='你好', time='17:38')

    def test_send_message(self):
        send_message(who='晚枫', message='你好')

    def test_chat_by_gpt(self):
        chat_by_gpt(who='程序员晚枫')

    def test_weixin_file(self):
        who = '程序员晚枫'
        file_path = r'./dfasd.py'
        # file_path = r'D:\workplace\code\github\PyOfficeRobot\tests\chat.py'
        # PyOfficeRobot.file.send_file(who, file)
        file.send_file(who, file_path)
        # file.send_file(who='程序员晚枫', file=r'D:\workplace\code\github\PyOfficeRobot\dev\contributor\gen\自动加好友\Excel_File\【企查查】查企业-高级搜索“涂料”(202302030683).xls')

    def test_group_send(self):
        group.send()

    def test_chat_ali(self):
        chat_ali(who='程序员晚枫', key=os.getenv('TY_KEY'))

    def test_chat_ds(self):
        chat_by_deepseek(who='晚枫', api_key="sk-pRdhASKn0wm5i7DjkdDfj5ENbRcpsqGrtV7hdFZZ6laV5aMk")

    def test_chat_zhipu(self):
        chat_by_zhipu(who='绋嬪簭鍛樻櫄鏋?', key='')

    def test_group_chat_by_keywords(self):
        who = '测试群'
        keywords = {
            "报名": "你好，这是报名链接：www.python-office.com",
            "学习": "你好，这是学习链接：www.python-office.com",
            "课程": "你好，这是课程链接：www.python-office.com"
        }
        match_type = 'contains'  # 关键字匹配类型 包含：contains  精确：exact
        chat_by_keywords(who=who, keywords=keywords, match_type=match_type)

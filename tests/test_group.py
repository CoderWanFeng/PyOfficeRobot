from PyOfficeRobot.api.group import WeChatBot


def test_group():
    bot = WeChatBot('keywords_data.csv', '测试群2')
    if bot.wx:
        bot.select_conversation()
        bot.process_messages()

if __name__ == '__main__':
    test_group()
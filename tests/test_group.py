from PyOfficeRobot.api.group import WeChatBot, WeChatBot2


# def test_group():
#     bot = WeChatBot('keywords_data.csv', '测试群2')
#     if bot.wx:
#         bot.select_conversation()
#         bot.process_messages()

def test_group2():
    print(111111)
    bot2 = WeChatBot2('InfoData.xlsx','测试群2')
    if bot2.wx:
        bot2.select_conversation()
        bot2.process_messages()

if __name__ == '__main__':
    # test_group()
    test_group2()
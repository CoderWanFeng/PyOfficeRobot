import pandas as pd
import uiautomation as auto
import time
import logging

# 配置日志
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class WeChatBot:
    def __init__(self, csv_path, contact_name):
        self.df = pd.read_csv(csv_path)
        self.contact_name = contact_name
        self.wx = self.find_wechat_window()
        self.last_processed_msg = None

    def find_wechat_window(self):
        try:
            wx_window = auto.WindowControl(searchDepth=3, Name='微信')
            if wx_window.Exists(0):
                wx_window.SwitchToThisWindow()
                return wx_window
            else:
                logging.warning("未找到微信窗口")
                return None
        except Exception as e:
            logging.error(f"查找微信窗口出错: {e}")
            return None

    def select_conversation(self):
        conversations = self.wx.ListControl()
        for conversation in conversations.GetChildren():
            if conversation.Name == self.contact_name:
                conversation.Click(simulateMove=False)
                break

    def get_last_message(self):
        try:
            message_list = self.wx.ListControl(Name='消息').GetChildren()
            if message_list:
                last_msg = message_list[-1].Name
                logging.info(f"获取到最后一条消息: {last_msg}")
                return last_msg
            else:
                logging.info("消息列表为空")
                return None
        except Exception as e:
            logging.error(f"获取最后一条消息出错: {e}")
            return None

    def send_reply(self, reply):
        try:
            self.wx.SendKeys(reply, waitTime=0)
            self.wx.SendKeys('{Enter}', waitTime=1)
            logging.info(f"回复内容是: {reply}")
        except Exception as e:
            logging.error(f"发送消息出错: {e}")

    def process_messages(self):
        while True:
            time.sleep(3)
            last_msg = self.get_last_message()
            if last_msg and last_msg != self.last_processed_msg:
                if last_msg not in self.df['回复内容'].values:
                    matched_replies = self.df[self.df['关键词'].apply(lambda x: x in last_msg)]['回复内容']
                    if not matched_replies.empty:
                        for reply in matched_replies:
                            reply = reply.replace('{br}', '\n')
                            self.send_reply(reply)
                        self.last_processed_msg = last_msg
                    else:
                        logging.info("没有匹配的关键字")
                else:
                    logging.info("最后一条消息是自动回复内容，跳过回复")
            else:
                logging.info("没有新消息或消息已处理")

# 示例调用
bot = WeChatBot('keywords_data.csv', '测试群2')
if bot.wx:
    bot.select_conversation()
    bot.process_messages()
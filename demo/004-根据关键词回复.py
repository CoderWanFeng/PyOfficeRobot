"""
已经更新了：发消息、发文件
今日更新：根据关键词，自动回复
1、安装：pip install python-office
2、代码演示
"""
import office

keywords = {
    "我要报名": "你好，这是报名链接：www.python-office.com",
    "点赞了吗？": "点了",
    "关注了吗？": "必须的",
    "投币了吗？": "三连走起",
}
office.wechat.chat_by_keywords(who='每天进步一点点', keywords=keywords)

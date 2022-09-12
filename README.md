又一个微信聊天机器人横空出世了，人人可用！

![](https://www.python-office.com/api/img-cdn/robot/wechat/cover.jpg)

之前给大家分享过一个微信机器人：[一个15分钟的视频，教你用Python创建自己的微信聊天机器人！](http://t.cn/A66p30bI)

但是这个机器人，需要基于网页版才能用；然而很多朋友的微信，是不能登录网页版微信的。

> 有没有一种微信机器人，任何人的微信都可以用，不需要网页微信功能呢？


在经过技术检索和开发以后，支持所有微信使用的：**PyOfficeRobot**来啦~

## 1、安装PyOfficeRobot

1行命令，安装PyOfficeRobot这个库
```
pip install -i https://pypi.tuna.tsinghua.edu.cn/simple PyOfficeRobot -U
```

## 2、微信机器人

先来一个简单的功能：**自动给指定好友发送消息。**

```
# 首先，将PyOfficeRobot模块导入到我们的代码块中。
from PyOfficeRobot.core.WeChatType import *

# 获取当前微信客户端
wx = WeChat()
# 获取会话列表
wx.GetSessionList()
# 向某人发送消息（以`文件传输助手`为例）
msg = '你好~我是程序员晚枫，全网同名，感谢点赞 + 转发'
who = '文件传输助手'
wx.ChatWith(who)  # 打开`文件传输助手`聊天窗口
# for i in range(10):
wx.SendMsg(msg)  # 向`文件传输助手`发送消息：你好~
```

## 3、功能说明

我最近开源了这个库的全部源代码，功能正在开发中，欢迎大家参与开发~

- ⭐GitHub：https://github.com/CoderWanFeng/PyOfficeRobot

---
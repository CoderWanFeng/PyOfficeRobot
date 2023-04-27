#!/usr/bin/env python
# -*- coding:utf-8 -*-

#############################################
# File Name: instruction_url.py
# Mail: 1957875073@qq.com
# Created Time:  2022-12-17 08:14:34
# Description: 有关 方法说明 的装饰器
#############################################

import os
# 每个文件的具体方法说明
from functools import wraps

from PyOfficeRobot.lib.CONST import SPLIT_LINE

chat_dict = {"chat_by_keywords": "https://www.bilibili.com/video/BV1fV4y1M7ju",
             "receive_message": "",
             "send_message": "https://www.bilibili.com/video/BV1Jt4y1j7F1",
             "send_message_by_time": "https://www.bilibili.com/video/BV1m8411b7LZ",
             "chat_by_gpt": "https://blog.51cto.com/u_15493782/6131326",
             }
file_dict = {
    "send_file": "https://www.bilibili.com/video/BV1te4y1y7Ro",

}

friend_dict = {
    "add": "https://www.bilibili.com/video/BV1DV4y1o7t2",

}

# 有多少文件需要说明
instruction_file_dict = {
    "chat.py": chat_dict,
    "file.py": file_dict,
    "friend.py": friend_dict,
}


def instruction(func):
    @wraps(func)
    def instruction_wrapper(*args, **kwargs):
        func_filename = os.path.basename(func.__code__.co_filename)  # 取出方法所在的文件名
        # 如果有这个文件，并且已经配置了方法名对应的说明链接，则打印出来
        if func_filename in instruction_file_dict.keys() and instruction_file_dict[func_filename][func.__name__]:
            print(SPLIT_LINE)
            print('【PyOfficeRobot，微信机器人全部功能】：https://www.python-office.com/office/robot.html')
            print(SPLIT_LINE)
            print(
                f'正在运行：office.{os.path.basename(func_filename)[:-3]}.{func.__name__} , 这个方法的使用说明：{instruction_file_dict[func_filename][func.__name__]}')
            print(SPLIT_LINE)
        instruction_res = func(*args, **kwargs)
        return instruction_res

    return instruction_wrapper


#############################################
# 以下是本文件的工具模块，用来更新方法和链接
from inspect import getmembers, isfunction


# 获取模块包含的方法名
def get_method_name(file):
    for method_name in getmembers(file):
        if isfunction(method_name[1]):
            print(f'"{method_name[0]}":"",')

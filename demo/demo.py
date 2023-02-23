# pip install PyOfficeRobot
# 建议使用清华大学的仓库，教程：https://www.bilibili.com/video/BV1SM411y7vw/
import PyOfficeRobot

# 注意：目前不用加参数，自动收集当前打开的微信群，未来会优化
PyOfficeRobot.file.get_group_list()
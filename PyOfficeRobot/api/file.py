from PyOfficeRobot.core.WeChatType import WeChat
from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction

wx = WeChat()


# @act_info(ACT_TYPE.FILE)
@instruction
def send_file(who: str, file: str):
    """
    发送任意类型的文件
    :param who:
    :param file: 文件路径
    :return:
    """
    wx.ChatWith(who)  # 打开聊天窗口
    # wx.SendFiles(file)
    wx.test_SendFiles(filepath=file, who=who)  # 添加who参数：雷杰

    # 注：为保证发送文件稳定性，首次发送文件可能花费时间较长，后续调用会缩短发送时间


import uiautomation as uia

""" isinstance()函数知识点:

type判断函数类型是什么，而isinstance是通过判断对象是否是已知的类型
但是isinstance比type高级一些（功能上的差异）
type() ，不考虑继承关系（子类不是父类类型）
isinstance() ，考虑继承关系（子类是父类类型）

a = 10086
isinstance (a,int)  # true
isinstance (a,str) # false
isinstance (a,(str,int,list))    # 只要满足元组类型中的其中一个即可，答案是满足，所以为false
"""


def get_group_list():
    """
    author: Zijian
    :return:
    """
    uia.uiautomation.SetGlobalSearchTimeout(2)
    wechat_win = uia.WindowControl(Depth=1, Name="微信", ClassName="WeChatMainWndForPC")
    wechat_win.SwitchToThisWindow()
    wechat_win.SetActive()
    wechat_win.SetTopmost()

    ### 点击右上角"聊天信息"
    chatinfo_button = wechat_win.ButtonControl(Name="聊天信息")
    chatinfo_button.Click(waitTime=0.5)

    ### 点击下拉菜单按钮"查看更多" 如果群里人数少这个按钮可能不存在
    more_button = wechat_win.ButtonControl(Name="查看更多")
    if more_button.Exists(0, 0):
        more_button.Click(0, 0, simulateMove=False)

    ### 获取"群聊名称"   先定位"群聊名称",再获取下一个兄弟节点,再获取第一个子节点,.Name取出里面的文字
    group_name = wechat_win.TextControl(Name="群聊名称").GetNextSiblingControl().GetFirstChildControl().Name
    print(group_name)

    ### 获取聊天成员列表控件,并且要获取边框范围
    nums = wechat_win.ListControl(Name="聊天成员")
    rect = nums.GetParentControl().BoundingRectangle  # 获取父级范围
    print(rect)
    numbers = nums.GetChildren()
    # for member in numbers:
    sum_info = []
    for index, member in enumerate(numbers):
        # print(type(member))
        # print(uia.ListItemControl)

        # 判断列表的每一个子成员 是不是列表项目控件,如果不是就跳过
        if not isinstance(member, uia.ListItemControl):
            continue

        # 判断列表的每一个子成员 是不是列表项目控件的名称是否存在名字,或有添加,删除,移除的关键字,如果不是就跳过。
        if not member.Name or member.Name in ("添加", "删除", "移出"):
            continue

        # 再判断当前时候有 快捷键被按下的状态,如果有直接break跳出循环停止采集.
        if uia.IsKeyPressed(uia.Keys.VK_F2):
            print("F2已被按下，停止采集")
            break

        # print("  -------  " + member.Name + "  -------  ")
        # 判断 每个成员的边界底部是否超出整个列表成员控件边框的底部,如果超过就发送下滚轮热键
        pos = member.BoundingRectangle
        if pos.bottom >= rect.bottom:
            uia.WheelDown(waitTime=0.5)

        ### 先左键点击每个成员,再用右键点击取消
        # 左键点击每个成员
        member.Click(waitTime=0.01)

        # 左键后会出现一个窗格控件  WalkControl返回的是一个元组类型(control, depth)控件,换句话说就是一个长度为2的元组对象 for循环迭代器时需要用两个变量去接收 如果用一个变量去接受返回的是一个元组长度为2的对象

        user_info_pane = uia.PaneControl(Name="微信")
        # user_info_pane拿到用户信息弹窗后截图保存本地文件
        # user_info_pane.CaptureToImage(savePath=f"{index}.png")

        # 用户信息字典
        user_info = {}
        user_info['群列表显示'] = member.Name
        for control, depth in uia.WalkControl(user_info_pane):
            # print(control)  for循环迭代器时需要用两个变量去接收 如果用一个变量去接受返回的是一个元组对象长度为2的

            if control.ControlTypeName == "TextControl":
                # k = c.Name.rstrip("：")  #rstrip() 删除 string 字符串末尾的指定字符
                text = control.Name
                if "：" in text:  # 如果这个文本控件类型里面的文本包含中文冒号":",就Get他的下一个兄弟节点 再用·Name的方法取出里面的内容
                    k = text.rstrip("：")
                    v = control.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
                if "来源" in text:
                    k = text
                    v = control.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
                if "备注名" in text:
                    k = text
                    v = control.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
                if "企业" == text:
                    k = text
                    a = control.GetNextSiblingControl()
                    v = a.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
                if "实名" == text:
                    k = text
                    a = control.GetNextSiblingControl()
                    v = a.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
                if "商城" == text:
                    k = text
                    v = control.GetNextSiblingControl().Name
                    # print(k,v)
                    user_info[k] = v
        print(user_info)
        sum_info.append(user_info)

        # 右键点击每个成员取消个人弹框 或者 发送ESC 快捷键
        # member.RightClick(waitTime=0.01)
        uia.SendKeys('{ESC}')

    print(sum_info)

    ### 将列表字典传给pandas库分析数据
    import pandas as pd
    save_path = "微信群信息采集.xlsx"

    df = pd.DataFrame(sum_info)
    df = df.reset_index(drop=True)  ### False是保留原先索引建立新的索引 True是重新建立索引
    df.to_excel(excel_writer=save_path, sheet_name='微信群信息采集', index=False)
    print(df)

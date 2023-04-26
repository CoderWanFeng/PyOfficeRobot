# -*- codeing = utf-8 -*-
# @Time : 2023/2/1 14:47
# @Project_File : 项目工程文档
# @Dir_Path : Python工程文档/成品区/微信大漠
# @File : Start.py
# @IDE_Name : PyCharm
# ============================================================
# ============================================================
# ============================================================
# ============================================================
# ============================================================


# excel_file = 77 line
# log_file = 56 line


import random
from datetime import datetime
from time import sleep

import xlrd
from pywinauto.application import *
from pywinauto.base_wrapper import *
from pywinauto.controls.uia_controls import *
from pywinauto.controls.uiawrapper import *
from pywinauto.keyboard import send_keys

Count = int(0)
MSG_LINE = int(1)


def Get_NowTime():
    "【当前时间：%s-%s-%s-：%s:%s:%s】正在读取第%s行数据,该公司的联系电话有如下:\n【%s】"
    year = datetime.now().year
    month = datetime.now().month
    day = datetime.now().day
    hour = datetime.now().hour
    minute = datetime.now().minute
    second = datetime.now().second
    return "【当前时间：{}-{}-{}：{}:{}:{}】".format(year, month, day, hour, minute, second)


def Printf_Log(str_log):
    """
    打印日志到指定文件\n
    :return:
    :rtype:
    """
    with open(r"./Auto_Log.txt", 'a', encoding='utf-8') as f:
        f.write(str_log + "\n")


def Is_Tel(TelNumber):
    """
    判断当前输入的联系电话是否为电话号码
    :param TelNumber:
    :type TelNumber:
    :return:
    :rtype:
    """
    if '-' in TelNumber or len(TelNumber) != 11 or TelNumber[0] != "1":
        return False
    else:
        return True


def Read_Excel(row):
    """
    读取整行数据,返回一个字典,字典的key为xls文件中的表头索引,其中 '电话' 这个可key是一个list列表对象\n
    :param row: 读取的行数据
    :type row: int类型
    :return: 整行数据
    :rtype: dict类型
    """
    # <editor-fold desc="代码块 : 读取整行数据">
    Excel_Worker = xlrd.open_workbook(filename=r"./Excel_File/【企查查】查企业-高级搜索“涂料”(202302030683).xls")
    Excel_Sheet = Excel_Worker.sheet_by_index(0)
    Excel_Dict_Data = dict()
    Excel_Dict_Data['企业名称'] = Excel_Sheet.cell(row, 0).value
    Excel_Dict_Data['登记状态'] = Excel_Sheet.cell(row, 1).value
    Excel_Dict_Data['法定代表人'] = Excel_Sheet.cell(rowx=row, colx=2).value
    Excel_Dict_Data['注册资本'] = Excel_Sheet.cell(rowx=row, colx=3).value
    Excel_Dict_Data['成立日期'] = Excel_Sheet.cell(rowx=row, colx=4).value
    Excel_Dict_Data['核准日期'] = Excel_Sheet.cell(rowx=row, colx=5).value
    Excel_Dict_Data['所属省份'] = Excel_Sheet.cell(rowx=row, colx=6).value
    Excel_Dict_Data['所属城市'] = Excel_Sheet.cell(rowx=row, colx=7).value
    Excel_Dict_Data['所属区县'] = Excel_Sheet.cell(rowx=row, colx=8).value
    Excel_Dict_Data['电话'] = Excel_Sheet.cell(rowx=row, colx=9).value
    Excel_Dict_Data['更多电话'] = Excel_Sheet.cell(rowx=row, colx=10).value
    Excel_Dict_Data['邮箱'] = Excel_Sheet.cell(rowx=row, colx=11).value
    Excel_Dict_Data['更多邮箱'] = Excel_Sheet.cell(rowx=row, colx=12).value
    Excel_Dict_Data['统一社会信用代码'] = Excel_Sheet.cell(rowx=row, colx=13).value
    Excel_Dict_Data['纳税人识别号'] = Excel_Sheet.cell(rowx=row, colx=14).value
    Excel_Dict_Data['注册号'] = Excel_Sheet.cell(rowx=row, colx=15).value
    Excel_Dict_Data['组织机构代码'] = Excel_Sheet.cell(rowx=row, colx=16).value
    Excel_Dict_Data['参保人数'] = Excel_Sheet.cell(rowx=row, colx=17).value
    Excel_Dict_Data['企业（机构）类型'] = Excel_Sheet.cell(rowx=row, colx=18).value
    Excel_Dict_Data['国标行业门类'] = Excel_Sheet.cell(rowx=row, colx=19).value
    Excel_Dict_Data['国标行业大类'] = Excel_Sheet.cell(rowx=row, colx=20).value
    Excel_Dict_Data['国标行业中类'] = Excel_Sheet.cell(rowx=row, colx=21).value
    Excel_Dict_Data['国标行业小类'] = Excel_Sheet.cell(rowx=row, colx=22).value
    Excel_Dict_Data['曾用名'] = Excel_Sheet.cell(rowx=row, colx=23).value
    Excel_Dict_Data['英文名'] = Excel_Sheet.cell(rowx=row, colx=24).value
    Excel_Dict_Data['网址'] = Excel_Sheet.cell(rowx=row, colx=25).value
    Excel_Dict_Data['企业地址'] = Excel_Sheet.cell(rowx=row, colx=26).value
    Excel_Dict_Data['最新年报地址'] = Excel_Sheet.cell(rowx=row, colx=27).value
    Excel_Dict_Data['经营范围'] = Excel_Sheet.cell(rowx=row, colx=28).value
    Excel_Dict_Data['电话'] = str(Excel_Dict_Data['电话']) + "；" + str(Excel_Dict_Data['更多电话'])
    Excel_Dict_Data['电话'] = list(Excel_Dict_Data['电话'].split("；"))
    Printf_Log("\n")
    Printf_Log("%s-----正在读取第%s行数据,当前公司名字为:【%s】" % (Get_NowTime(), row, Excel_Dict_Data['企业名称']))
    Printf_Log(
        "%s-----正在读取第%s行数据,该公司的联系电话有如下:\n【%s】" % (Get_NowTime(), row, Excel_Dict_Data['电话']))

    return Excel_Dict_Data
    # </editor-fold>


def FindFriend(WX_Windows, Tel_Number, ErrorCount=Count):
    """
    该函数用于在进入通讯录界面,输入手机号码一系列的操作\n
    :param WX_Windows: 微信窗口对象
    :param Tel_Number: 需要搜索的手机号码
    :param ErrorCount:
    :return:
    """
    # <editor-fold desc="代码块 : 执行搜索手机号码操作">
    print("提示:当前在正在对输入的手机号码进行一个查找")
    Button_AddFrient = WX_Windows.child_window(title="添加朋友")
    Button_AddFrient_Wrapper = Button_AddFrient.wrapper_object()
    """:type : pywinauto.controls.uiawrapper.UIAWrapper"""
    Button_AddFrient_Wrapper.draw_outline(colour='red', thickness=5)
    Button_AddFrient_Wrapper.click_input()

    Edit_Number = WX_Windows.child_window(title="微信号/手机号")
    Edit_Number_Wrapper = Edit_Number.wrapper_object()
    """:type : pywinauto.controls.uiawrapper.UIAWrapper"""
    Edit_Number_Wrapper.draw_outline(colour='red', thickness=5)
    Edit_Number_Wrapper.click_input()
    Edit_Number_Wrapper.type_keys(Tel_Number)

    Button_Find = WX_Windows.child_window(title='搜索：')
    Button_Find_Wrapper = Button_Find.wrapper_object()
    """:type : pywinauto.controls.uiawrapper.UIAWrapper"""
    Button_Find_Wrapper.draw_outline(colour='red', thickness=5)
    Button_Find_Wrapper.click_input()

    try:

        Button_AddTXL = WX_Windows.child_window(title='添加到通讯录')
        Button_AddTXL_Wrapper = Button_AddTXL.wrapper_object()
        """:type : pywinauto.controls.uiawrapper.UIAWrapper"""
        Button_AddTXL_Wrapper.draw_outline(colour="red", thickness=5)
        Button_AddTXL_Wrapper.click_input()
        print("提示:当前号码[ %s ]可以进行一个添加" % (Tel_Number))

        return True
    except:
        try:
            print(
                "错误:当前号码出现错误,可能原因为 1.已经添加过该好友  2.该号码不存在  3.出现频繁次数了  当前号码为 : %s" % (
                    Tel_Number))
            print("注意:当前出错次数为 : %d" % (ErrorCount))

            find_result = WX_Windows.child_window(title=r'搜索结果')
            find_result_wrapper = find_result.wrapper_object()
            assert isinstance(find_result_wrapper, UIAWrapper)
            find_childs = find_result_wrapper.children()

            if (len(find_childs) == 1):
                print("\t\t 【警告】：%s-----当前号码并没有进行一个点击,请后续再进行操作" % (Get_NowTime()))

            elif (len(find_childs) == 2):
                find_childs[0].draw_outline(colour='red', thickness=5)
                find_resule_info = find_childs[0].element_info
                assert isinstance(find_resule_info, UIAElementInfo)
                print("\t\t 【错误】：%s-----%s" % (Get_NowTime(), find_resule_info.name))
                Printf_Log("\t\t 【错误】：%s-----%s" % (Get_NowTime(), find_resule_info.name))

        except:
            print("\t\t %s提示:当前遍历号码[ %s ]已经是微信好友" % (Get_NowTime(), Tel_Number))
            Printf_Log("\t\t %s提示:当前遍历号码[ %s ]已经是微信好友" % (Get_NowTime(), Tel_Number))

        return False
    # </editor-fold>


def Open_TXL():
    # <editor-fold desc="代码块 : 进入通讯录界面">
    """
    该函数为打开通讯录界面\n
    :return:
    """
    print("提示:当前正在进入通讯录界面")
    Button_SC_Wrapper.draw_outline(colour="red", thickness=5)
    Button_SC_Wrapper.click_input()
    sleep(0.5)
    Button_EXE_Wrapper.draw_outline(colour="red", thickness=5)
    Button_EXE_Wrapper.click_input()
    sleep(0.5)
    Button_LT_Wrapper.draw_outline(colour="red", thickness=5)
    Button_LT_Wrapper.click_input()
    sleep(0.5)
    Button_EXE_Wrapper.draw_outline(colour="red", thickness=5)
    Button_EXE_Wrapper.click_input()
    sleep(0.5)
    Button_TXL_Wrapper.draw_outline(colour="red", thickness=5)
    Button_TXL_Wrapper.click_input()
    sleep(0.5)
    # </editor-fold>


def Carry_TXL(App_Object, Hello_Str, Tel_Number, Excel_Data):
    """
    该函数为申请好友验证时的函数胡\n
    :param App_Object: 微信窗口操作对象
    :param Hello_Str: 打招呼的字符串
    :param Tel_Number: 备注名,用于备注该用户的手机号码
    :return: None
    """
    # <editor-fold desc="代码块 : 填写验证信息以及确认操作">
    # Anchor为锚点 Target为目标点
    Anchor_1 = App_Object.child_window(title='发送添加朋友申请')
    Anchor_1.draw_outline(colour='green', thickness=5)
    Anchor_1_Wrapper = Anchor_1.wrapper_object()
    Anchor_1_Wrapper_Parent = Anchor_1_Wrapper.element_info.parent
    Children_1 = Anchor_1_Wrapper_Parent.children()
    Target_1 = Children_1[1]
    Target_1_Wrapper = UIAWrapper(Target_1)
    Target_1_Wrapper.draw_outline(colour='red', thickness=5)
    Target_1_Wrapper.click_input()
    Target_1_Wrapper.click_input()
    for _ in range(50):
        send_keys("{VK_END}")
        send_keys("{VK_BACK}")
    Target_1_Wrapper.type_keys(Hello_Str)

    Anchor_2 = App_Object.child_window(title='备注名')
    Anchor_2.draw_outline(colour='green', thickness=5)
    Anchor_2_Wrapper = Anchor_2.wrapper_object()
    Anchor_2_Wrapper_Parent = Anchor_2_Wrapper.element_info.parent
    Children_2 = Anchor_2_Wrapper_Parent.children()
    Target_2 = Children_2[1]
    Target_2_Wrapper = UIAWrapper(Target_2)
    Target_2_Wrapper.draw_outline(colour='red', thickness=5)
    Target_2_Wrapper.click_input()
    Target_2_Wrapper.click_input()
    for _ in range(30):
        send_keys("{VK_END}")
        send_keys("{VK_BACK}")
    Target_2_Wrapper.type_keys("【%s-%s-%s】" % (Excel_Data['所属城市'], Excel_Data['法定代表人'], Tel_Number))
    # </editor-fold>
    # App_Object.capture_as_image().save(Tel_Number + ".jpg")

    Button_Yes = App_Object.child_window(title='确定')
    Button_Yes_Wrapper = Button_Yes.wrapper_object()
    Button_Yes_Wrapper.draw_outline(colour='red', thickness=5)
    Button_Yes_Wrapper.click_input()

    # Button_Yes = App_Object.child_window(title='确定')
    # Button_Yes_Wrapper = Button_Yes.wrapper_object()
    # Button_Parent = Button_Yes_Wrapper.element_info.parent
    # Button_No = Button_Parent.children()[1]
    # Button_No_Wrapper = UIAWrapper(Button_No)
    # Button_No_Wrapper.draw_outline(colour='red',thickness=5)
    # Button_No_Wrapper.click_input()


if __name__ == "__main__":

    # <editor-fold desc="代码块 : 获取微信窗口句柄">
    Get_WeChat_Hwnd = lambda: win32gui.FindWindow("WeChatMainWndForPC", "微信")
    WeChat_Hwnd = Get_WeChat_Hwnd()
    # </editor-fold>

    # <editor-fold desc="代码块 : 设置窗口状态以及位置">
    win32gui.ShowWindow(WeChat_Hwnd, 9)
    # win32gui.SetWindowPos(WeChat_Hwnd, win32con.HWND_TOPMOST, 0,0,800,800, win32con.SWP_NOMOVE | win32con.SWP_NOACTIVATE| win32con.SWP_NOOWNERZORDER|win32con.SWP_SHOWWINDOW)
    win32gui.MoveWindow(WeChat_Hwnd, 1920 - 800, 0, 800, 800, True)
    print("Msg:窗口设置完毕")
    # </editor-fold>

    # <editor-fold desc="代码块 : 获取微信窗口对象以及微信窗口的动作操作包装对象">
    App_Object = Application(backend="uia").connect(handle=WeChat_Hwnd)
    WX_Windows = App_Object['微信']
    """:type : pywinauto.application.WindowSpecification"""
    WX_Wrapper = WX_Windows.wrapper_object()
    """:type : pywinauto.controls.uiawrapper.UIAWrapper"""
    WX_Wrapper.draw_outline(colour="red", thickness=5)
    # </editor-fold>

    # <editor-fold desc="代码块 : 获取左侧按钮对象以及操作包装对象(合计:3对)">
    Button_LT = WX_Windows.child_window(title='聊天')
    Button_TXL = WX_Windows.child_window(title='通讯录')
    Button_SC = WX_Windows.child_window(title='收藏')
    Button_EXE = WX_Windows.child_window(title='小程序面板')

    Button_LT_Wrapper = Button_LT.wrapper_object()
    Button_TXL_Wrapper = Button_TXL.wrapper_object()
    Button_SC_Wrapper = Button_SC.wrapper_object()
    Button_EXE_Wrapper = Button_EXE.wrapper_object()

    # assert isinstance(Button_LT_Wrapper,UIAWrapper)
    # assert isinstance(Button_TXL_Wrapper,UIAWrapper)
    # assert isinstance(Button_SC_Wrapper,UIAWrapper)
    # assert isinstance(Button_EXE_Wrapper,UIAWrapper)

    # </editor-fold>

    row = int(input("请输入在多少行开始 : "))
    while True:
        Excel_Data = Read_Excel(row=row)

        print("\n")
        print("\n")
        print("\n")
        print("\n")
        print("\n")
        print("提示:当前运行至Excel中的第%d行数据" % (row))
        print("提示:当前运行至Excel中的第%d行数据,该公司名字为 : %s" % (row, Excel_Data['企业名称']))
        print("提示:当前运行至Excel中的第%d行数据,该公司法定代表人为 : %s" % (row, Excel_Data['法定代表人']))
        print("提示:当前运行至Excel中的第%d行数据,该公司所属城市为 : %s" % (row, Excel_Data['所属城市']))
        print("提示:当前运行至Excel中的第%d行数据,该公司网址为 : %s" % (row, Excel_Data['网址']))
        print("提示:当前运行至Excel中的第%d行数据,该公司电话为 : %s" % (row, Excel_Data['电话']))
        print("\n")

        for _ in Excel_Data["电话"]:

            if Is_Tel(_):
                print("\t\t %s提示:当前遍历号码[ %s ] 是一个手机号码,可以进行一个添加尝试" % (Get_NowTime(), _))
                Printf_Log("\t\t %s提示:当前遍历号码[ %s ] 是一个手机号码,可以进行一个添加尝试" % (Get_NowTime(), _))

                Open_TXL()
                result = FindFriend(WX_Windows=WX_Windows, Tel_Number=_, ErrorCount=Count)

                if (result == False):
                    Count = Count + 1
                    print("\t\t %s提示:当前遍历号码[ %s ] 不存在有微信号" % (Get_NowTime(), _))
                    Printf_Log("\t\t %s提示:当前遍历号码[ %s ] 不存在有微信号" % (Get_NowTime(), _))

                elif (result == True):
                    Carry_TXL(App_Object=WX_Windows, Hello_Str="你好,老板。我们是做大小防白的,有需要可以了解下",
                              Tel_Number=_, Excel_Data=Excel_Data)
                    print("\t\t %s提示:当前遍历号码[ %s ] 存在有微信号,已经完成添加好友操作" % (Get_NowTime(), _))
                    Printf_Log("\t\t %s提示:当前遍历号码[ %s ] 存在有微信号,已经完成添加好友操作" % (Get_NowTime(), _))
                Random_Wait_Time = random.randint(15, 25)
                print("\t\t %s-----随机等待时间为%d秒" % (Get_NowTime(), Random_Wait_Time))
                sleep(Random_Wait_Time)

            else:
                print("\t\t %s提示:当前遍历号码[ %s ] 不为手机号码,此处直接跳过" % (Get_NowTime(), _))
                Printf_Log("\t\t %s提示:当前遍历号码[ %s ] 不为手机号码,此处直接跳过" % (Get_NowTime(), _))

            if Count > 10:
                print("\t\t %s错误:当前遍历出现错误过多,此处开始不再添加,具体情况查看日志" % (Get_NowTime()))
                Printf_Log("\t\t %s错误:当前遍历出现错误过多,此处开始不再添加,具体情况查看日志" % (Get_NowTime()))

                sleep(8888)
                exec
        row = row + 1

# -*- coding: UTF-8 -*-
'''
@作者 ：B站/抖音/微博/小红书/公众号，都叫：程序员晚枫
@微信 ：CoderWanFeng : https://mp.weixin.qq.com/s/Nt8E8vC-ZsoN1McTOYbY2g
@个人网站 ：www.python-office.com
@Date    ：2023/4/23 23:01 
@Description     ：
'''
from time import sleep

from pywinauto.application import *
from pywinauto.base_wrapper import *
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.keyboard import send_keys
from pywinauto.uia_element_info import UIAElementInfo

from PyOfficeRobot.lib.decorator_utils.instruction_url import instruction


def _Carry_TXL(App_Object, Hello_Str, Tel_Number, notes):
    """
    该函数为申请好友验证时的函数胡\n
    :param App_Object: 微信窗口操作对象
    :param Hello_Str: 打招呼的字符串
    :param Tel_Number: 备注名,用于备注该用户的手机号码
    :param notes: 备注名
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
    Target_2_Wrapper.type_keys(notes)
    # </editor-fold>
    # App_Object.capture_as_image().save(Tel_Number + ".jpg")

    Button_Yes = App_Object.child_window(title='确定')
    Button_Yes_Wrapper = Button_Yes.wrapper_object()
    Button_Yes_Wrapper.draw_outline(colour='red', thickness=5)
    Button_Yes_Wrapper.click_input()


def Get_NowTime():
    "【当前时间：%s-%s-%s-：%s:%s:%s】正在读取第%s行数据,该公司的联系电话有如下:\n【%s】"
    return time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())


def _find_friend(WX_Windows, Tel_Number, ErrorCount):
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

    Button_Find = WX_Windows.child_window(title='搜索：' + Tel_Number)
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
                print("\t\t 【错误】：%s-----%s" % (Get_NowTime(), find_resule_info.name))

        except:
            print("\t\t %s提示:当前遍历号码[ %s ]已经是微信好友" % (Get_NowTime(), Tel_Number))
            print("\t\t %s提示:当前遍历号码[ %s ]已经是微信好友" % (Get_NowTime(), Tel_Number))

        return False
    # </editor-fold>


def _Open_TXL(Button_LT_Wrapper, Button_TXL_Wrapper, Button_SC_Wrapper, Button_EXE_Wrapper):
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


@instruction
def add(num_notes, msg):
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

    Count = 1
    # </editor-fold>
    for num in num_notes.keys():
        _Open_TXL(Button_LT_Wrapper, Button_TXL_Wrapper, Button_SC_Wrapper, Button_EXE_Wrapper)
        result = _find_friend(WX_Windows=WX_Windows, Tel_Number=num, ErrorCount=Count)
        _Carry_TXL(App_Object=WX_Windows, Hello_Str=msg,
                   Tel_Number=num, notes=num_notes[num])


if __name__ == "__main__":
    msg = "你好，我是程序员晚枫，全网同名。"
    # num_list = ['15603052573', '19112440257']
    num_notes = {
        # 'CoderWanFeng': '北京-晚枫-学生',
        'hdylw1024': '上海-晚枫-乞丐',
    }
    add(msg=msg, num_notes=num_notes)

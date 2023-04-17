# -*- codeing = utf-8 -*-
# @Time : 2023/3/7 15:43
# @Project_File : Start.py
# @Dir_Path : 成品区/获取微信群聊名称
# @File : TEMPLATE.py.py
# @IDE_Name : PyCharm
# ============================================================
# ============================================================
# ============================================================
# ============================================================
# ============================================================

import time

import win32gui
from pywinauto.application import Application
from pywinauto.controls.uiawrapper import UIAWrapper
from pywinauto.element_info import ElementInfo
from pywinauto.controls.uia_controls import ListItemWrapper
from pywinauto.controls.uia_controls import ButtonWrapper
from pywinauto.keyboard import send_keys

from win32gui import FindWindow
from win32gui import ShowWindow
from win32gui import MoveWindow
from win32gui import SetForegroundWindow

from xlwt.Workbook import Workbook
from xlwt.Worksheet import Worksheet



Group_Names = set()

def Open_the_address_book_manager():
    """
    打开通讯录管理器\n
    :return:
    :rtype:
    """

    # <editor-fold desc="代码段 : 打开微信中的通讯录管理器">
    Hwnd = FindWindow("WeChatMainWndForPC", "微信")
    ShowWindow(Hwnd, 9)
    SetForegroundWindow(Hwnd)
    MoveWindow(Hwnd,1920-800,0,800,800,True)
    App_Object = Application(backend="uia").connect(handle=Hwnd)
    WX_Windows = App_Object['微信']

    WX_Wrapper : UIAWrapper =  WX_Windows.wrapper_object()
    WX_Wrapper.draw_outline(colour='red',thickness=10)
    BUTTON_ADDRESS_ELEMENT = WX_Windows.child_window(title='通讯录')
    BUTTON_ADDRESS_WRAPPER : UIAWrapper = BUTTON_ADDRESS_ELEMENT.wrapper_object()

    BUTTON_ADDRESS_WRAPPER.draw_outline(colour='red',thickness=10)
    BUTTON_ADDRESS_WRAPPER.click_input()

    BUTTON_ADDRESS_MANAGE_ELEMENT = WX_Windows.child_window(title='通讯录管理')
    BUTTON_ADDRESS_MANAGE_WRAPPER : UIAWrapper = BUTTON_ADDRESS_MANAGE_ELEMENT.wrapper_object()
    BUTTON_ADDRESS_MANAGE_WRAPPER.draw_outline(colour='red',thickness=10)
    BUTTON_ADDRESS_MANAGE_WRAPPER.click_input()

    if(FindWindow("ContactManagerWindow","通讯录管理")):
        win32gui.CloseWindow(Hwnd)
        print(True)
        return True
    else:
        print(False)
        return False
    # </editor-fold>

def Get_all_group_names() -> list:
    """
    获取所有群聊的名字,以列表形式返回\n
    :return:
    :rtype:
    """
    # <editor-fold desc="代码段 : 设置 '通讯录管理' 窗口位置">
    Hwnd = FindWindow("ContactManagerWindow", "通讯录管理")
    x1,y1,x2,y2 = win32gui.GetWindowRect(Hwnd)
    width = x2 - x1
    height = y2 - y1
    point_x1 = 1920-width
    point_y1 = 0
    win32gui.MoveWindow(Hwnd,point_x1,point_y1,width,height,1)
    # </editor-fold>


    # <editor-fold desc="代码段 : 获取锚点元素">
    App_Object = Application(backend="uia").connect(handle=Hwnd)
    TXLGQ_Windows = App_Object['通讯录管理']
    TXLGQ_Wrapper: UIAWrapper = TXLGQ_Windows.wrapper_object()
    TXLGQ_Wrapper.draw_outline(colour='red', thickness=10)

    Pane_FriendAuthority = TXLGQ_Windows.child_window(title='朋友权限', control_type='Pane')
    Pane_Label = TXLGQ_Windows.child_window(title='标签', control_type='Pane')
    Pane_Recent_Groups = TXLGQ_Windows.child_window(title='最近群聊', control_type='Pane')
    Pane_List_Groups = object()

    Wrapper_Pane_FriendAuthority : UIAWrapper = Pane_FriendAuthority.wrapper_object()
    Wrapper_Pane_Label : UIAWrapper = Pane_Label.wrapper_object()
    Wrapper_Pane_Recent_Groups : UIAWrapper = Pane_Recent_Groups.wrapper_object()

    Wrapper_Pane_FriendAuthority.draw_outline(colour='red',thickness=5)
    Wrapper_Pane_Label.draw_outline(colour='red',thickness=5)
    Wrapper_Pane_Recent_Groups.draw_outline(colour='red',thickness=5)

    Wrapper_Parent : UIAWrapper = Wrapper_Pane_FriendAuthority.parent()


    # </editor-fold>


    # <editor-fold desc="代码段 : 打开最近群聊">
    if(len(Wrapper_Parent.children()) == 4):
        Wrapper_Pane_Recent_Groups.draw_outline()
        Wrapper_Pane_Recent_Groups.click_input()
        Pane_List_Groups : UIAWrapper = Wrapper_Parent.children()[4]
    else:
        Pane_List_Groups : UIAWrapper = Wrapper_Parent.children()[4]
    #BUTTON_RECENT_GROUPS_WRAPPER.click_input()
    # </editor-fold>

    #内部函数
    def Get_ListItem_At_Present() -> list:
        """
        返回当前页面所有的群名字\n
        流程 :
            1.获取所有的 'ListItem' 元素\n
            2.对所有的 'ListItem' 进行迭代遍历,并且在迭代中定位Pane的位置,并且获取其 '文本'\n
                定位 Pane 时,因为 Pane 是 Button 元素的同级元素,需先定位至 Button 元素,然后获取其父元素,再由其 父元素 去获取Pane元素\n
                其中 Pane 元素跟 Button 的位置是固定的 \n
                Button 位于 父元素的索引为 0 的位置 \n
                Pane 位于 父元素的索引为 1 的位置 \n
        :return:
        :rtype:
        """
        names = list()
        LIST_LISTITEM: list = Pane_List_Groups.descendants(control_type='ListItem')
        for _ in LIST_LISTITEM:
            Button : ButtonWrapper = _.descendants(control_type='Button')[0]
            Parent = Button.parent()
            Pane = Parent.children()[1]
            Pane.draw_outline()
            names.append(Pane.element_info.name)
        return names

    Pane_List_Groups.descendants(control_type='ListItem')[0].draw_outline(colour='red',thickness=5)
    Pane_List_Groups.descendants(control_type='ListItem')[0].click_input()

    #预估群数量(必须大于实际数量)
    for _ in range(100):
        namelist = Get_ListItem_At_Present()
        for _ in namelist:
            Group_Names.add(_)
        send_keys("{VK_DOWN}")
        print("提示:当前已经获取的群数量为 : %d"%(len(Group_Names)))


def Save_Groups_To_Xls(groups=Group_Names,savefile="微信群名字.xls"):
    """

    :return:
    :rtype:
    """
    workbook = Workbook()
    worksheet: Worksheet = Workbook.add_sheet(workbook,"sheet1")

    row = int(0)
    for _ in Group_Names:
        print(_)
        worksheet.write(row,0,_)
        worksheet.write(row,1,"True")
        row = row + 1
    workbook.save("微信群名字.xls")


if __name__ == "__main__":
    result : bool = Open_the_address_book_manager()

    if(result == False):
        exit()

    Get_all_group_names()

    Save_Groups_To_Xls()
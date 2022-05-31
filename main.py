# -*- coding: utf-8 -*-
#
# @Time    : 2022-03-03 15:30:20
# @Author  : Yunlong Shi
# @Email   : syljoy@163.com
# @FileName: main.py
# @Software: PyCharm
# @Github  : https://github.com/syljoy
# @Desc    : 划题搜答案主程序

import PyHook3 as pyHook
import pythoncom
import win32api
import win32con
import win32clipboard  # 剪贴板
import time
import pyperclip

from op_excel import get_qb

import warnings
warnings.filterwarnings("ignore")

VK_CODE = {
    'ctrl': 0x11,
    'c': 0x43,
}

OLD_S = "s"

FLAG = 1

def key_copy():
    win32api.keybd_event(VK_CODE["ctrl"], 0, 0, 0)
    time.sleep(0.01)
    win32api.keybd_event(VK_CODE["c"], 0, 0, 0)
    win32api.keybd_event(VK_CODE["c"], 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(VK_CODE["ctrl"], 0, win32con.KEYEVENTF_KEYUP, 0)

def get_text():
    win32clipboard.OpenClipboard()
    try:
        d = win32clipboard.GetClipboardData(win32con.CF_TEXT)
    except TypeError as e:
        # print(e)
        return None

    win32clipboard.CloseClipboard()
    return d.decode('GBK')

def onKeyboardEvent(event):
    # print('按下键盘：', event.Key)  # 返回按下的键
    global FLAG
    if event.Key == 'Q':
        print("退出程序，并终止代码")
        hm.UnhookMouse()  # 取消监听事件
        hm.UnhookKeyboard()  # 取消监听事件
        # 虽然取消了，但是程序不终止
        win32api.PostQuitMessage(0)
    elif event.Key == '1' and FLAG != 1:
        FLAG = 1
        print("已切换至单选题模式")
    elif event.Key == '2' and FLAG != 2:
        FLAG = 2
        print("已切换至多选题模式")
    elif event.Key == '3' and FLAG != 3:
        FLAG = 3
        print("已切换至判断题模式")
    return True
def search_answer(qb, s):
    answer = qb.loc[qb['* 题干'].str.contains(s)]
    if len(answer) == 0:
        return '未搜索到答案'
    elif len(answer) == 1:
        return answer['* 题干'].values[0][:20]+'...\t\t答案：'+answer['* 标答'].values[0]
    elif len(answer) > 1:
        return '找到多个答案，请重新选择'

def onMouseEvent(event):
    # 鼠标移动：mouse move
    # print(event.MessageName)
    if event.MessageName in ['mouse left up']:
        # 当鼠标按键按下，并抬起后执行
        key_copy()
        time.sleep(0.1)
        s = get_text()
        # 清空粘贴版
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
        # print('剪贴板：', s)
        global OLD_S
        if s and OLD_S != s:
            OLD_S = "".join(s)
            if FLAG == 1:  # 单选题模式：
                print(search_answer(single_qb, s))
                print("\t\t现为单选题模式(按2切换多选题模式，按3切换判断题模式)")
            elif FLAG == 2:
                print(search_answer(multiple_qb, s))
                print("\t\t现为多选题模式(按1切换单选题模式，按3切换判断题模式)")
            elif FLAG == 3:
                print(search_answer(tf_qb, s))
                print("\t\t现为判断题模式(按1切换单选题模式，按2切换多选题模式)")
            else:
                print('请选择题目模式：\n\t按1切换单选题模式\t按2切换多选题模式\t按3切换判断题模式')

    # 返回True代表将事件继续传给其他句柄，为False则停止传递，即被拦截
    return True


if __name__ == '__main__':
    # 获取题库
    single_qb, multiple_qb, tf_qb = get_qb()
    # 题目模式

    # 创建pyhook管理器
    hm = pyHook.HookManager()
    # 监听键盘 # KeyUp/KeyDown/KeyChar/KeyAll
    hm.KeyDown = onKeyboardEvent  # 事件绑定
    # 4
    hm.HookKeyboard()  # 开始监听

    hm.MouseAll = onMouseEvent  # 将OnMouseEvent函数绑定到MouseAllButtonsDown事件上
    hm.HookMouse()

    # 循环监听
    pythoncom.PumpMessages()  # 会进入一个循环，不会执行后续语句，想退出只能使用win32api.PostQuitMessage(0) 退出。

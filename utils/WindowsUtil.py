import os

import win32con
import win32gui
import win32api
import win32process

from utils import TimeUtil


def find_window(window_name, class_name=None) -> int:
    """
    查找窗口句柄
    :param window_name: 窗口名
    :param class_name: 类型
    :return:
    """
    return win32gui.FindWindow(class_name, window_name)


def find_child_window(hwnd, window_name, class_name=None) -> int:
    """
    查找窗口子句柄
    :param hwnd: 父窗口句柄
    :param window_name: 窗口名
    :param class_name: 类名
    :return:
    """
    temp_hwnd = win32gui.GetWindow(hwnd, win32con.GW_CHILD)
    while temp_hwnd != 0:
        if window_name is not None:
            temp_text = get_window_text(temp_hwnd)
            if temp_text == window_name:
                if class_name is None:
                    return temp_hwnd
                else:
                    if get_class_name(temp_hwnd) == class_name:
                        return temp_hwnd
        else:
            if class_name is None:
                raise ValueError("window_name和class_name不能都为None")
            else:
                if get_class_name(temp_hwnd) == class_name:
                    return temp_hwnd

        temp_hwnd = win32gui.GetWindow(temp_hwnd, win32con.GW_HWNDNEXT)
    return 0


def get_window_text(hwnd) -> str:
    """
    获取窗口标题
    :param hwnd: 句柄
    :return: 窗口标题
    """
    return win32gui.GetWindowText(hwnd)


def get_class_name(hwnd) -> str:
    """
    获取窗口类名
    :param hwnd:
    :return:
    """
    return win32gui.GetClassName(hwnd)


def ergodic_window_hwnd(window_name, class_name=None, vague=True) -> list:
    """
    遍历句柄
    :param window_name:
    :param class_name:
    :param vague: 模糊匹配
    :return: 句柄list
    """
    result_list = []
    temp_hwnd = win32gui.GetDesktopWindow()
    temp_hwnd = win32gui.GetWindow(temp_hwnd, win32con.GW_CHILD)
    while temp_hwnd != 0:
        temp_text = get_window_text(temp_hwnd)
        # 完全匹配
        if not vague:
            if temp_text == window_name:
                if class_name is None:
                    result_list.append(temp_hwnd)
                else:
                    if get_class_name(temp_hwnd) == class_name:
                        result_list.append(temp_hwnd)
        else:
            if temp_text.find(window_name) >= 0:
                if class_name is None:
                    result_list.append(temp_hwnd)
                else:
                    if get_class_name(temp_hwnd) == class_name:
                        result_list.append(temp_hwnd)
        temp_hwnd = win32gui.GetWindow(temp_hwnd, win32con.GW_HWNDNEXT)

    return result_list


def get_foreground_window() -> int:
    """
    获取当前窗口句柄
    :return:
    """
    return win32gui.GetForegroundWindow()


def set_window_text(hwnd, text):
    """
    修改窗口标题
    :param hwnd:要修改的窗口
    :param text:要修改的标题
    :return:
    """
    win32gui.SetWindowText(hwnd, text)


def get_client_rect(hwnd) -> (int, int, int, int):
    """
    获取窗口大小（不包括标题栏、边框、滚动条等装饰的区域）
    :param hwnd: 句柄
    :return: （0,0,宽,高）
    """
    return win32gui.GetClientRect(hwnd)


def get_window_rect(hwnd):
    """
    获取窗口在屏幕的位置（左上角坐标,右下角坐标）
    :param hwnd: 句柄
    :return: 左上角坐标x,左上角坐标y,右下角坐标x,右下角坐标y）
    """
    return win32gui.GetWindowRect(hwnd)


def set_window_rect(hwnd, width, height):
    """
    设置窗口大小
    :param hwnd: 句柄
    :param width: 宽
    :param height: 高
    :return:
    """
    rc_window = win32gui.GetWindowRect(hwnd)
    rc_client = win32gui.GetClientRect(hwnd)
    # 标题栏、边框、滚动条等装饰的区域长宽度
    border_width = (rc_window[2] - rc_window[0]) - (rc_client[2] - rc_client[0])
    border_height = (rc_window[3] - rc_window[1]) - (rc_client[3] - rc_client[1])
    win32gui.SetWindowPos(hwnd, 0, 0, 0, width + border_width, height + border_height,
                          win32con.SWP_NOMOVE | win32con.SWP_NOZORDER)


def client_to_screen(hwnd, x, y) -> (int, int):
    """
    客户端坐标转屏幕坐标
    :param hwnd:
    :param x:
    :param y:
    :return:
    """
    return win32gui.ClientToScreen(hwnd, (x, y))


def is_window(hwnd) -> bool:
    """
    判断窗口存在
    :param hwnd:
    :return:
    """
    return win32gui.IsWindow(hwnd)


def start_program(path):
    """
    启动程序
    :param path: 路径
    :return:
    """
    # 获取文件名
    file_name = os.path.basename(path)
    # 获取文件目录
    dir_path = os.path.dirname(path)
    win32api.ShellExecute(None, 'open', file_name, None, dir_path, win32con.SW_SHOWNORMAL)


def close_window(hwnd):
    """
    关闭窗口
    :param hwnd:
    :return:
    """
    win32gui.SendMessage(hwnd, win32con.WM_CLOSE, win32con.WM_NULL, win32con.WM_NULL)


def close_process(hwnd):
    """
    关闭进程
    :param hwnd:
    :return:
    """
    tid, pid = win32process.GetWindowThreadProcessId(hwnd)
    process_handle = win32api.OpenProcess(win32con.PROCESS_TERMINATE, False, pid)
    if process_handle.handle > 0:
        win32api.TerminateProcess(process_handle, 0)


def display_window(hwnd):
    """
    显示窗口（取消最小化）
    :param hwnd:
    :return:
    """
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        TimeUtil.delay(50)


def activate_window(hwnd):
    """
    激活窗口
    :param hwnd:
    :return:
    """
    win32gui.SetForegroundWindow(hwnd)


def hide_window(hwnd):
    """
    隐藏窗口
    :param hwnd:
    :return:
    """
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0,
                          win32con.SWP_NOSIZE | win32con.SWP_NOMOVE | win32con.SWP_HIDEWINDOW)


def show_window(hwnd):
    """
    显示窗口
    :param hwnd:
    :return:
    """
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOMOVE |
                          win32con.SWP_SHOWWINDOW)


def move_window(hwnd, x, y):
    """
    移动窗口
    :param hwnd:
    :param x:
    :param y:
    :return:
    """
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOP, x, y, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOZORDER)


def window_top(hwnd):
    """
    窗口置前
    :param hwnd:
    :return:
    """
    win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


def cancel_window_top(hwnd):
    """
    窗口取消置前
    :param hwnd:
    :return:
    """
    win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 0, 0, 0, 0, win32con.SWP_NOSIZE | win32con.SWP_NOMOVE)


def minimize_window(hwnd):
    """
    最小化窗口
    :param hwnd:
    :return:
    """
    win32gui.SendMessage(hwnd, win32con.WM_SYSCOMMAND, win32con.SC_MINIMIZE, win32con.WM_NULL)


def back_activate_window(hwnd):
    """
    后台激活
    :param hwnd:
    :return:
    """
    win32gui.PostMessage(hwnd, win32con.WM_ACTIVATEAPP, 1, 0)

import ctypes

import win32api
import win32con
import win32gui

from utils import TimeUtil

"""
    0-15位：指定当前消息的重复次数。其值就是用户按下该键后自动重复的次数，但是重复次数不累积
    16-23位：指定其扫描码，其值依赖于OEM厂商
    24位：指定该按键是否为扩展按键，所谓扩展按键就是Ctrl,Alt之类的，如果是扩展按键，其值为1，否则为0
    25-28位：保留字段，暂时不可用
    29位：指定按键时的上下文，其值为1时表示在按键时Alt键被按下，其值为0表示WM_SYSKEYDOWN消息因没有任何窗口有键盘焦点而被发送到当前活动窗口。
    30位：指定该按键之前的状态，其值为1时表示该消息发送前，该按键是被按下的，其值为0表示该消息发送前该按键是抬起的。
    31位：指定其转换状态，对WM_SYSKEYDOWN消息而言，其值总为0
"""


def get_down_lparam(virtual_key) -> int:
    """
    获取PostMessage的lparam参数，当WM_KEYDOWN时
    :param virtual_key: 按键的ASCII码
    :return:
    """
    scancode = win32api.MapVirtualKey(virtual_key, 0)
    secondbyte = ('00' + hex(scancode)[2:].zfill(2))[-2:]
    # 见最上面的说明
    s = '00' + secondbyte + '0001'
    return int(s, 16)


def get_alt_down_lparam(virtual_key) -> int:
    """
    获取PostMessage的lparam参数，当WM_KEYDOWN且ALT按下时
    :param virtual_key: 按键的ASCII码
    :return:
    """
    scancode = win32api.MapVirtualKey(virtual_key, 0)
    secondbyte = ('00' + hex(scancode)[2:].zfill(2))[-2:]
    # 见最上面的说明
    s = '20' + secondbyte + '0001'
    return int(s, 16)


def get_alt_up_lparam(virtual_key) -> int:
    """
    获取PostMessage的lparam参数，当WM_KEYUP且ALT放开时
    :param virtual_key: 按键的ASCII码
    :return:
    """
    scancode = win32api.MapVirtualKey(virtual_key, 0)
    secondbyte = ('00' + hex(scancode)[2:].zfill(2))[-2:]
    # 见最上面的说明
    s = 'E0' + secondbyte + '0001'
    return int(s, 16)


def get_up_lparam(virtual_key) -> int:
    """
    获取PostMessage的lparam参数，当WM_KEYUP时
    :param virtual_key: 按键的ASCII码
    :return:
    """
    scancode = win32api.MapVirtualKey(virtual_key, 0)
    secondbyte = ('00' + hex(scancode)[2:].zfill(2))[-2:]
    # 见最上面的说明
    s = 'C0' + secondbyte + '0001'
    return int(s, 16)


def key_press(hwnd, virtual_key):
    key_down(hwnd, virtual_key)
    TimeUtil.delay(50)
    key_up(hwnd, virtual_key)


def key_down(hwnd, virtual_key):
    win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, virtual_key, get_down_lparam(virtual_key))


def key_up(hwnd, virtual_key):
    win32gui.PostMessage(hwnd, win32con.WM_KEYUP, virtual_key, get_up_lparam(virtual_key))


def move_to(hwnd, x, y):
    point = win32api.MAKELONG(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_MOUSEMOVE, None, point)


def left_down(hwnd, x, y, delay=50):
    move_to(hwnd, x, y)
    TimeUtil.delay(delay)
    point = win32api.MAKELONG(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONDOWN, win32con.MK_LBUTTON, point)


def left_up(hwnd, x, y, delay=50):
    move_to(hwnd, x, y)
    TimeUtil.delay(delay)
    point = win32api.MAKELONG(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_LBUTTONUP, win32con.MK_LBUTTON, point)


def left_click(hwnd, x, y, delay=50):
    left_down(hwnd, x, y, delay)
    TimeUtil.delay(delay)
    left_up(hwnd, x, y, delay)
    TimeUtil.delay(delay)


def right_down(hwnd, x, y, delay=50):
    move_to(hwnd, x, y)
    TimeUtil.delay(delay)
    point = win32api.MAKELONG(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_RBUTTONDOWN, win32con.MK_RBUTTON, point)


def right_up(hwnd, x, y, delay=50):
    move_to(hwnd, x, y)
    TimeUtil.delay(delay)
    point = win32api.MAKELONG(x, y)
    win32gui.PostMessage(hwnd, win32con.WM_RBUTTONUP, win32con.MK_RBUTTON, point)


def right_click(hwnd, x, y, delay=50):
    right_down(hwnd, x, y, delay)
    TimeUtil.delay(delay)
    right_up(hwnd, x, y, delay)


def alt_combo_key(hwnd, virtual_key):
    """
    alt的组合键
    :param hwnd:
    :param virtual_key:
    :return:
    """
    win32gui.PostMessage(hwnd, win32con.WM_SYSKEYDOWN, virtual_key, get_alt_down_lparam(virtual_key))
    TimeUtil.delay(50)
    win32gui.PostMessage(hwnd, win32con.WM_SYSKEYUP, virtual_key, get_alt_up_lparam(virtual_key))


def send_text(hwnd, text):
    """
    输入字符串
    :param hwnd:
    :param text:
    :return:
    """
    for i in range(len(text)):
        win32gui.SendMessage(hwnd, win32con.WM_IME_CHAR, ord(text[i]), 0)
        TimeUtil.delay(10)


def foreground_move_to(x, y):
    """
    鼠标移动
    :param x: x坐标
    :param y: y坐标
    """
    ctypes.windll.user32.SetCursorPos(x, y)


def foreground_mouse_down(x, y, button):
    """
    鼠标按下
    :param x: x坐标
    :param y: y坐标
    :param button: left middle right 鼠标的左键 右键 滚轮
    """
    ev = None
    if button not in ("left", "middle", "right"):
        raise ValueError('button 取值 "left", "middle", 或 "right", 现在值为：%s' % button)

    if button == "left":
        ev = 0x0002
    elif button == "middle":
        ev = 0x0020
    elif button == "right":
        ev = 0x0008
    # 屏幕的宽度和高度
    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))
    converted_x = 65536 * x // width + 1
    converted_y = 65536 * y // height + 1
    ctypes.windll.user32.mouse_event(ev, ctypes.c_long(converted_x), ctypes.c_long(converted_y), 0, 0)


def foreground_mouse_up(x, y, button):
    """
    鼠标弹起
    :param x: x坐标
    :param y: y坐标
    :param button: left middle right 鼠标的左键 右键 滚轮
    """
    ev = None
    if button not in ("left", "middle", "right"):
        raise ValueError('button 取值 "left", "middle", 或 "right", 现在值为：%s' % button)

    if button == "left":
        ev = 0x0004
    elif button == "middle":
        ev = 0x0040
    elif button == "right":
        ev = 0x0010

    # 屏幕的宽度和高度
    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))
    converted_x = 65536 * x // width + 1
    converted_y = 65536 * y // height + 1
    ctypes.windll.user32.mouse_event(ev, ctypes.c_long(converted_x), ctypes.c_long(converted_y), 0, 0)


def linear(n):
    if not 0.0 <= n <= 1.0:
        raise ValueError("tween参数取值应在0.0到1.0之间.")
    return n


def get_point_on_line(x1, y1, x2, y2, n):
    """
    Returns an (x, y) tuple of the point that has progressed a proportion ``n`` along the line defined by the two
    ``x1``, ``y1`` and ``x2``, ``y2`` coordinates.

    This function was copied from pytweening module, so that it can be called even if PyTweening is not installed.
    """
    x = ((x2 - x1) * n) + x1
    y = ((y2 - y1) * n) + y1
    return x, y


def foreground_mouse_move_drag(x1, y1, x2, y2, duration, tween_x=linear):
    """
    鼠标按住拖动
    :param x1: 开始拖动的x坐标
    :param y1: 开始拖动的y坐标
    :param x2: 结束拖动的x坐标
    :param y2: 结束拖动的y坐标
    :param duration:
    :param tween_x:
    """
    # 先移动到要按住的位置
    foreground_move_to(x1, y1)
    # 按住鼠标左键
    foreground_mouse_down(x1, y1, "left")
    # 延时用
    sleep_amount = 0
    # 屏幕的宽度和高度
    width, height = (ctypes.windll.user32.GetSystemMetrics(0), ctypes.windll.user32.GetSystemMetrics(1))
    # 当前鼠标位置
    start_x, start_y = win32api.GetCursorPos()
    # 如果持续时间足够短，只需立即将光标移动到那里即可。
    steps = [(x2, y2)]
    if duration > 0.1:
        # 非即时移动拖动涉及补间
        num_steps = max(width, height)
        sleep_amount = duration / num_steps
        if sleep_amount < 0.05:
            num_steps = int(duration / 0.05)
            sleep_amount = duration / num_steps

        steps = [get_point_on_line(start_x, start_y, x2, y2, tween_x(n / num_steps)) for n in range(num_steps)]
        # 确保最后一个位置是实际目的地
        steps.append((x2, y2))

    for tween_x, tween_y in steps:
        if len(steps) > 1:
            TimeUtil.delay(sleep_amount * 1000)

        tween_x = int(round(tween_x))
        tween_y = int(round(tween_y))

        ctypes.windll.user32.SetCursorPos(tween_x, tween_y)
    # 释放鼠标左键
    foreground_mouse_up(x2, y2, "left")

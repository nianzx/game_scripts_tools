import ctypes
import math

import threading

import cv2
import win32con
import win32gui
from struct import pack, calcsize

import numpy as np
from PIL import Image, ImageOps


def screenshot_to_bitmap_array(hwnd, left, top, right, bottom):
    """
    截图转成image对象
    :param hwnd: 要截图的窗口句柄
    :param left: 窗口中截图区域左上角x坐标
    :param top: 窗口中截图区域左上角y坐标
    :param right: 右下角x坐标
    :param bottom: 右下角y坐标
    :return: 返回PIL的image对象
    """
    width = right - left
    height = bottom - top
    # 创建互斥锁对象
    mutex = threading.Lock()
    # 源DC
    src_dc = win32gui.GetDC(hwnd)
    # 内存DC
    mem_dc = win32gui.CreateCompatibleDC(src_dc)
    # 创建与指定的设备环境相关的设备兼容的位图
    save_bitmap = ctypes.windll.gdi32.CreateCompatibleBitmap(src_dc, width, height)
    # 选择一对象到指定的设备上下文环境中
    win32gui.SelectObject(mem_dc, save_bitmap)
    # 使用互斥锁保护代码块
    with mutex:
        # 实际截图
        win32gui.BitBlt(mem_dc, 0, 0, width, height, src_dc, left, top, win32con.SRCCOPY)
        # 分配内存空间
        _data = ctypes.create_string_buffer(height * width * 4)
        bmi = pack('LHHHH', calcsize('LHHHH'), width, height, 1, 32)
        # 获取的图像信息按照底部-顶部的顺序排列
        ctypes.windll.gdi32.GetDIBits(mem_dc, save_bitmap, 0, height, _data, bmi, win32con.DIB_RGB_COLORS)
    # 将图像数据转为numpy数组
    image_array = np.ctypeslib.as_array(_data)
    # 翻转图像数据
    image_array = image_array.reshape((height, width, 4))
    image_array = np.flipud(image_array)
    # 调整颜色通道的顺序 (BGR -> RGB)
    image_array = image_array[:, :, [2, 1, 0, 3]]
    # 将 numpy 数组转换为一维，以便创建 PIL Image 对象
    image_array = image_array.flatten()
    # 创建PIL Image对象
    image = Image.frombytes('RGBA', (width, height), image_array)
    # 释放内存
    win32gui.DeleteObject(save_bitmap)
    win32gui.DeleteDC(mem_dc)
    win32gui.ReleaseDC(hwnd, src_dc)
    return image


def screenshot_to_file(hwnd, left, top, right, bottom, path):
    """
    截图到文件
    :param hwnd: 要截图的窗口句柄
    :param left: 窗口中截图区域左上角x坐标
    :param top: 窗口中截图区域左上角y坐标
    :param right: 右下角x坐标
    :param bottom: 右下角y坐标
    :param path: 保存的文件路径
    """
    image = screenshot_to_bitmap_array(hwnd, left, top, right, bottom)
    image.save(path)


def get_color_rgb(color_str):
    """
    获取颜色RGB信息
    :param color_str:  颜色字符串 如：B9B9B9-101010
    :return: 格式：{'x': 0, 'y': 0, 'DR': 16, 'DG': 16, 'DB': 16, 'R': 142, 'G': 89, 'B': 35}
    """
    result = {}

    point = color_str.split("|")
    if len(point) == 3:
        x = int(point[0])
        y = int(point[1])
        result['x'] = x
        result['y'] = y
        color_str = point[2]
    else:
        result['x'] = 0
        result['y'] = 0

    result['DR'] = 0
    result['DG'] = 0
    result['DB'] = 0

    # 判断颜色字符串中是否包含 '-'
    i_pos = color_str.find('-')
    if i_pos > 0:
        # 字符串长度为13，则解析RGB和RGB差值
        if len(color_str) == 13:
            # 解析RGB值
            result['R'] = int(color_str[0:2], 16)
            result['G'] = int(color_str[2:4], 16)
            result['B'] = int(color_str[4:6], 16)
            # 解析RGB差值
            result['DR'] = int(color_str[i_pos + 1:i_pos + 3], 16)
            result['DG'] = int(color_str[i_pos + 3:i_pos + 5], 16)
            result['DB'] = int(color_str[i_pos + 5:i_pos + 7], 16)
            return result
            # 字符串长度为6，则只解析RGB值
    elif len(point[2]) == 6:
        result['R'] = int(color_str[0:2], 16)
        result['G'] = int(color_str[2:4], 16)
        result['B'] = int(color_str[4:6], 16)
        return result
    else:
        return {}
        # 如果以上两种情况都不满足，返回错误信息
    return {}


def rgb_compare(pixel_rgb, rgb_attribute) -> bool:
    """
    比较颜色值
    :param pixel_rgb: 截图中获取的像素点颜色值,用Image.getpixel()获取的值
    :param rgb_attribute: 多点颜色值，包含偏差值 格式：{'x': 0, 'y': 0, 'DR': 16, 'DG': 16, 'DB': 16, 'R': 142, 'G': 89, 'B': 35}
    :return: 匹配返回true，反之返回false
    """
    if abs(pixel_rgb[0] - rgb_attribute['R']) <= rgb_attribute['DR'] and abs(pixel_rgb[1] - rgb_attribute['G']) <= \
            rgb_attribute['DG'] and abs(pixel_rgb[2] - rgb_attribute['B']) <= rgb_attribute['DB']:
        return True
    return False


def multi_point_find_color(image, multi_point_color_str, similarity=1.0):
    """
    多点找色
    :param image: 要比对的图片
    :param multi_point_color_str: 多点字符串，字符串格式参考按键精灵，例如
    ['844B1B-101010','1|4|763C12-101010,-4|8|864D1C-101010,21|2|7F4417-101010,24|2|EDC060-101010,-2|-7|F1C968-101010,23|14|ECAD55-101010,1|4|763C12-101010']
    :param similarity: 相似度
    :return: 找到的坐标，没找到返回(-1，-1)，注：这边返回的是相对坐标值
    """
    # 截图的宽高
    width, height = image.size
    # 第一个点颜色
    single_color = multi_point_color_str[0]
    # 多点颜色list
    multi_color = multi_point_color_str[1]
    color_list = multi_color.split(',')
    # 多点数量
    multi_num = len(color_list)
    # 最大不匹配数
    max_mismatch_num = multi_num - math.trunc(multi_num * similarity)
    # 不匹配数
    mismatch_num = 0

    for x in range(width):
        for y in range(height):
            # 获取像素点颜色
            pixel_rgb = image.getpixel((x, y))
            # 第一个点颜色值
            rgb_attribute = get_color_rgb(single_color)

            # 颜色比较
            if rgb_compare(pixel_rgb, rgb_attribute):
                # 用来判断是否得到坐标点
                flag = True
                # 进入比较其他多点
                for color in color_list:
                    temp_rgb_attribute = get_color_rgb(color)
                    # 点的坐标不能越界
                    tx = x + temp_rgb_attribute['x']
                    ty = y + temp_rgb_attribute['y']
                    if tx >= width or ty >= height or tx < 0 or ty < 0:
                        flag = False
                        break

                    temp_pixel_rgb = image.getpixel((tx, ty))

                    if not rgb_compare(temp_pixel_rgb, temp_rgb_attribute):
                        # 不匹配
                        mismatch_num += 1

                        if mismatch_num > max_mismatch_num:
                            flag = False
                            break

                if flag:
                    # 这边获取的是相对坐标
                    return x, y
    return -1, -1


def screenshot_multi_point_find_color(hwnd, left, top, right, bottom, multi_point_color_str, similarity=1.0):
    """
    截图多点找色
    :param hwnd: 要截图的窗口句柄
    :param left: 窗口中截图区域左上角x坐标
    :param top: 窗口中截图区域左上角y坐标
    :param right: 右下角x坐标
    :param bottom: 右下角y坐标
    :param multi_point_color_str: 多点字符串，格式参考按键精灵，例如
    ['844B1B-101010','1|4|763C12-101010,-4|8|864D1C-101010,21|2|7F4417-101010,24|2|EDC060-101010,-2|-7|F1C968-101010,23|14|ECAD55-101010,1|4|763C12-101010']
    :param similarity: 相似度
    :return: 找到的绝对坐标 找不到返回(-1,-1)
    """
    image = screenshot_to_bitmap_array(hwnd, left, top, right, bottom)
    x, y = multi_point_find_color(image, multi_point_color_str, similarity)
    if x == -1 and y == -1:
        return x, y
    # 转成绝对路径
    return x + left, y + top


def _kmp(needle, haystack):
    """
    Knuth-Morris-Pratt (KMP) 字符串搜索算法
    :param needle: 要搜索的短字符串
    :param haystack: 在其中搜索模式的长字符串
    :return: needle在haystack中第一次出现的位置
    """
    shifts = [1] * (len(needle) + 1)
    shift = 1
    for pos in range(len(needle)):
        while shift <= pos and needle[pos] != needle[pos - shift]:
            shift += shifts[pos - shift]
        shifts[pos + 1] = shift

    start_pos = 0
    match_len = 0
    for c in haystack:
        while match_len == len(needle) or match_len >= 0 and needle[match_len] != c:
            start_pos += shifts[match_len]
            match_len -= shifts[match_len]
        match_len += 1
        if match_len == len(needle):
            yield start_pos


def screenshot_find_picture(hwnd, left, top, right, bottom, dest_image_url, grayscale=True):
    """
    截图找图(全匹配)
    :param hwnd: 要截图的窗口句柄
    :param left: 窗口中截图区域左上角x坐标
    :param top: 窗口中截图区域左上角y坐标
    :param right: 右下角x坐标
    :param bottom: 右下角y坐标
    :param dest_image_url: 目标图片url
    :param grayscale: 是否转成灰度图片进行比较 这大概提升30%的效率 默认True
    :return: 坐标
    """
    src_image = screenshot_to_bitmap_array(hwnd, left, top, right, bottom)
    dest_image = Image.open(dest_image_url)
    if grayscale:
        # 图片转灰度图片
        src_image = ImageOps.grayscale(src_image)
        dest_image = ImageOps.grayscale(dest_image)
    else:
        # 如果不使用灰度，请确保我们比较的是RGB图像，而不是RGBA图像。
        if src_image.mode == 'RGBA':
            src_image = src_image.convert('RGB')
        if dest_image.mode == 'RGBA':
            dest_image = dest_image.convert('RGB')

    dest_width, dest_height = dest_image.size
    src_width, src_height = src_image.size

    dest_image_data = tuple(dest_image.getdata())
    src_image_data = tuple(src_image.getdata())

    dest_image_rows = [
        dest_image_data[y * dest_width: (y + 1) * dest_width] for y in range(dest_height)
    ]  # LEFT OFF - check this
    dest_image_first_row = dest_image_rows[0]

    assert (
            len(dest_image_first_row) == dest_width
    ), 'The calculated width of first row of the needle image is not the same as the width of the image.'
    assert [len(row) for row in dest_image_rows] == [
        dest_width
    ] * dest_height, 'The dest_image_rows aren\'t the same size as the original image.'

    # 从最左边的一列开始
    for y in range(src_height):
        # 拿出目标图片的第一行 在 源图片中一行一行对比 使用kmp算法
        for matchx in _kmp(dest_image_first_row, src_image_data[y * src_width: (y + 1) * src_width]):
            found_match = True
            # 在源图片中找到了目标图片的第一行，对比剩下的
            for searchy in range(1, dest_height):
                dest_start = (searchy + y) * src_width + matchx
                if (
                        dest_image_data[searchy * dest_width: (searchy + 1) * dest_width]
                        != src_image_data[dest_start: dest_start + dest_width]
                ):
                    found_match = False
                    break
            if found_match:
                # 返回找到的坐标，转成绝对坐标
                return matchx + left, y + top
    # 没找到
    return -1, -1


def screenshot_find_picture2(hwnd, left, top, right, bottom, dest_image_url, confidence=0.9, grayscale=True, step=1):
    """
    截图找图(模糊匹配)
    :param hwnd: 句柄
    :param hwnd: 要截图的窗口句柄
    :param left: 窗口中截图区域左上角x坐标
    :param top: 窗口中截图区域左上角y坐标
    :param right: 右下角x坐标
    :param bottom: 右下角y坐标
    :param dest_image_url: 目标图片url
    :param grayscale: 是否转成灰度图片进行比较 这大概提升30%的效率 默认True
    :param step: 取值1或者2，为2时confidence默认0.95
    :param confidence: 相似度
    :return: 坐标
    """
    src_image = screenshot_to_bitmap_array(hwnd, left, top, right, bottom)
    dest_image = Image.open(dest_image_url)

    if dest_image.size[0] > src_image.size[0] or dest_image.size[1] > src_image.size[1]:
        # 要找的图片比源图大
        return -1, -1

    # 加载图片并转灰度图片
    dest_array = np.array(dest_image.convert('RGB'))
    dest_image = dest_array[:, :, ::-1].copy()
    if grayscale:
        dest_image = cv2.cvtColor(dest_image, cv2.COLOR_BGR2GRAY)

    src_array = np.array(src_image.convert('RGB'))
    src_image = src_array[:, :, ::-1].copy()
    if grayscale:
        src_image = cv2.cvtColor(src_image, cv2.COLOR_BGR2GRAY)

    if step == 2:
        # 等于2的时候 速度可以提升3倍 相似度默认0.95
        confidence *= 0.95
        dest_image = dest_image[::step, ::step]
        src_image = src_image[::step, ::step]
    else:
        step = 1

    result = cv2.matchTemplate(src_image, dest_image, cv2.TM_CCOEFF_NORMED)
    match_indices = np.arange(result.size)[(result > confidence).flatten()]
    matches = np.unravel_index(match_indices[:10000], result.shape)

    if len(matches[0]) == 0:
        return -1, -1

    matchx = matches[1] * step
    matchy = matches[0] * step

    for x, y in zip(matchx, matchy):
        return x, y

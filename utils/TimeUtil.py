import time


def delay(ms):
    """
    延时
    :param ms: 毫秒
    :return:
    """
    start_time = time.perf_counter()
    while time.perf_counter() - start_time < ms / 1000:
        pass

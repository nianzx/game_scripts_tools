import ctypes


class OCR(object):

    def __init__(self, DLL_PATH, TESSDATA_PREFIX, lang):
        self.DLL_PATH = DLL_PATH
        self.TESSDATA_PREFIX = TESSDATA_PREFIX
        self.lang = lang
        self.ready = False
        if self.do_init():
            self.ready = True

    def do_init(self):
        self.tesseract = ctypes.cdll.LoadLibrary(self.DLL_PATH)
        # 设置函数返回的数据类型
        self.tesseract.TessBaseAPICreate.restype = ctypes.c_uint64
        # 创建一个Tesseract API实例
        self.api = self.tesseract.TessBaseAPICreate()
        # 初始化Tesseract API实例，并指定tessdata路径和语言模式。
        rc = self.tesseract.TessBaseAPIInit3(ctypes.c_uint64(self.api), self.TESSDATA_PREFIX, self.lang)
        if rc:
            self.tesseract.TessBaseAPIDelete(ctypes.c_uint64(self.api))
            print('Could not initialize tesseract.\n')
            return False
        self.tesseract.TessVersion.restype = ctypes.c_char_p
        tesseract_version = self.tesseract.TessVersion()
        print('tesseract版本 > ' + str(tesseract_version))
        return True

    def get_text(self, path):
        if not self.ready:
            return False
        self.tesseract.TessBaseAPIProcessPages(ctypes.c_uint64(self.api), path, None, 0, None)
        # 设置函数返回的数据类型
        self.tesseract.TessBaseAPIGetUTF8Text.restype = ctypes.c_uint64
        # 获取识别结果的UTF-8编码文本。
        text_out = self.tesseract.TessBaseAPIGetUTF8Text(ctypes.c_uint64(self.api))
        return bytes.decode(ctypes.string_at(text_out)).strip()

    def __del__(self):
        # 释放Tesseract API实例。
        self.tesseract.TessBaseAPIDelete(ctypes.c_uint64(self.api))


if __name__ == '__main__':
    # 加载dll和语言包
    DLL_PATH = '../dll/tesseract50.dll'
    TESSDATA_PREFIX = b'../tessdata'
    lang = b'chi_sim'
    ocr = OCR(DLL_PATH, TESSDATA_PREFIX, lang)
    # 传入图片进行识别
    result = ocr.get_text(b'../1.jpg')
    print(result)

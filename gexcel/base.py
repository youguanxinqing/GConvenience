"""
TIME: 2019/11/24 10:51
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import os

from .config import *


class BaseExcelMetaClass(type):
    def __new__(mcs, name, bases, attrs):
        if name == "ToExcel":
            attrs.update({"_wb_obj_name": "Workbook"})
        elif name == "FromExcel":
            attrs.update({"_wb_obj_name": "load_workbook"})

        return super().__new__(mcs, name, bases, attrs)


class BaseExcel(metaclass=BaseExcelMetaClass):
    def __init__(self, d=".", f="test.xlsx"):
        """
        :param d: 目录
        :param f: 文件名字.格式
        """
        self._target_dir = d
        self._target_file = f
        self._target_path = os.path.join(d, f)

        self._wb = None

    @classmethod
    def _alpha_to_digit(cls, symbol):
        """
        转十进制编码
        """
        if isinstance(symbol, str) and len(symbol) == 1:
            return cls._convert_alpha(symbol)
        elif isinstance(symbol, int):
            return symbol
        else:
            raise ValueError("err arg")

    @staticmethod
    def _convert_alpha(alpha):
        """
        'a' -> 1, 'A' -> 1
        'b' -> 2, 'B' -> 2 ...
        """
        return ord(alpha.upper()) - ASCII_CODE_START

    def close(self):
        """
        资源释放
        """
        pass

    def __del__(self):
        self.close()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __enter__(self):
        return self

"""
TIME: 2019/11/24 10:52
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import os
import gc

import openpyxl

from .base import BaseExcel
from .config import *


class FromExcel(BaseExcel):
    def __init__(self, d=".", f: str = "", sheet=None):
        self._check_args(f)

        super().__init__(d, f)
        self._wb = getattr(openpyxl, self._wb_obj_name)(self._target_path, read_only=True)
        self._sheet = self._wb.active

    def read_page(self,
                  row=DEFAULT_ROW_START,
                  num=DEFAULT_NUM,
                  col_start: int = DEFAULT_COL_START,
                  end_flag: int = END_FLAG_NUN):
        """
        按 sheet 进行整页读取，最大支持 1000 行
        """
        count = 0
        for i in range(row, 10000):
            # 设置内容结束栅栏
            # 如果出现连续 end_flag 行内容都为空, 那么认为该 excel 文件
            # 的内容已经结束, 退出
            if count == end_flag:
                break

            line_content = self.read_line(i, num, col_start)
            if not any(line_content):
                count += 1
                continue
            else:
                count = 0

            yield line_content

    def read_rect(self,
                  lt: int,
                  lb: int,
                  rt: Union[str, int],
                  rb: Union[str, int]):
        """
        按照矩形范围读取内容

        lt: left top
        lb: left bottom
        rt: right top
        rb: right bottom
        """
        rt = self._alpha_to_digit(rt)
        rb = self._alpha_to_digit(rb)
        num = rb - rt + 1
        return [self.read_line(i, num, rt) for i in range(lt, lb + 1)]

    def read_line(self, row, num, col_start=1):
        """
        row: 某行
        nums: 读取个数
        col_start: 从哪一行起
        """
        num = self._alpha_to_digit(num)
        return [
            self.read_point(row, col)
            for col in range(col_start, col_start + num)
        ]

    def read_point(self, row, col):
        """
        定位读取 (row, col)
        """
        col = self._alpha_to_digit(col)
        location = self._sheet.cell(row=row, column=col)
        return location.value

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

    @staticmethod
    def _check_args(filename):
        """
        参数校验：1. 文件名是否有效， 2. 文件类型是否支持
        """
        if not filename:
            raise ValueError("filename must be valid")

        if os.path.splitext(filename)[-1] not in SUPPORT_FILE_EXTEND:
            raise NotImplementedError(f"only support {SUPPORT_FILE_EXTEND}, currently")

    def close(self):
        self.__del__()
        gc.collect()

    def __del__(self):
        self._wb.close()

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.__del__()

    def __enter__(self):
        return self

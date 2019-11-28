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


class ToExcel(BaseExcel):
    def __init__(self, d=".", f="gtest.xlsx"):
        super().__init__(d, f)

        self._wb = getattr(openpyxl, self._wb_obj_name)()
        self._sheet = self._wb.active
        self._row = DEFAULT_ROW_START

    @staticmethod
    def _pre_check(path):
        if not os.path.exists(path):
            raise FileNotFoundError(f"could not found the file specified '{path}'")

    @property
    def locate_point(self):
        raise NotImplementedError

    @locate_point.setter
    def locate_point(self, location_and_value):
        """
        指定单元格写入数据
        """
        row, col, value = location_and_value
        col = self._alpha_to_digit(col)

        self._sheet.cell(row=row, column=col, value=value)

    @property
    def locate_line(self):
        raise NotImplementedError

    @locate_line.setter
    def locate_line(self, location_and_values):
        """
        指定行写入数据
        """
        col = DEFAULT_COL_START
        if len(location_and_values) == 1:
            row, vals = self._row, location_and_values[0]
        elif len(location_and_values) == 2:
            row, vals = location_and_values
        elif len(location_and_values) == 3:
            row, col, vals = location_and_values
        else:
            raise ValueError("err args")

        num = len(vals)
        [self._sheet.cell(row=row, column=i, value=vals[i-col]) for i in range(col, col+num)]
        self._row += 1

    @property
    def title(self):
        raise NotImplementedError

    @title.setter
    def title(self, desps):
        """
        设置表头
        """
        if self._row == DEFAULT_ROW_START:
            self.locate_line = (desps,)

    def save(self):
        self._wb.save(self._target_path)
        return self

    def close(self):
        self._tear_down()

    def _tear_down(self):
        self.row = DEFAULT_ROW_START
        self._wb.close()
        gc.collect()

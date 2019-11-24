"""
TIME: 2019/11/24 10:51
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import os
import openpyxl


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

    # def __del__(self):
    #     self._wb.close()
    #
    # def __exit__(self, exc_type, exc_val, exc_tb):
    #     self.__del__()
    #
    # def __enter__(self):
    #     return self


class ToExcel(BaseExcel):
    def __init__(self, d=".", f="gtest.xlsx"):
        super().__init__(d, f)
        self._wb = getattr(openpyxl, self._wb_obj_name)





if __name__ == "__main__":
    te = ToExcel()
    fe = FromExcel()

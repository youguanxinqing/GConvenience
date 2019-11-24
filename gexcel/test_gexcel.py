"""
TIME: 2019/11/24 12:01
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import os
import unittest

from .fromexcel import FromExcel


class TestFromExcel(unittest.TestCase):
    def test_read_point(self):
        d = os.path.dirname(os.path.abspath(__file__))
        d = os.path.dirname(d)
        fe = FromExcel(d, "2015.xlsx")
        res = fe.read_point(3, "A")
        self.assertEquals(res, "种类")

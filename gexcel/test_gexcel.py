"""
TIME: 2019/11/24 12:01
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import os
import hashlib
import unittest

import openpyxl

from .fromexcel import FromExcel


class TestFromExcel(unittest.TestCase):
    def setUp(self):
        m = hashlib.md5("demo".encode("utf-8"))
        self.filename = f"{m.hexdigest()}.xlsx"
        wb = openpyxl.Workbook()
        sheet = wb.active
        # 制作模拟数据
        [[sheet.cell(row, col, f"({row}, {col})") for col in range(1, 21)] for row in range(1, 21)]
        wb.save(self.filename)
        wb.close()

        self.fe = FromExcel(f=self.filename)

    def tearDown(self):
        self.fe.close()
        if os.path.exists(self.filename):
            try:
                os.remove(self.filename)
            except PermissionError:
                pass

    def test_read_point(self):
        """
        测试单元格读取
        """
        res = self.fe.read_point(3, "A")
        self.assertEqual(res, "(3, 1)")

        res = self.fe.read_point(100, "Z")
        self.assertEqual(res, None)

    def test_read_line(self):
        row, num = 10, 5
        res = self.fe.read_line(row, num)
        expected_res = [f"{(10, i)}" for i in range(1, num+1)]
        self.assertEqual(expected_res, res)

        col_start = 5
        res = self.fe.read_line(row, num, 5)
        expected_res = [f"{(10, i)}" for i in range(col_start, num+col_start)]
        self.assertEqual(expected_res, res)

    def test_read_rect(self):
        lt, lb, rt, rb = 3, 10, "B", "J"
        res = self.fe.read_rect(lt, lb, rt, rb)

        j_start = FromExcel._alpha_to_digit(rt)
        j_end = FromExcel._alpha_to_digit(rb)
        expected_res = [
            [f"{(i, j)}" for j in range(j_start, j_end+1)] for i in range(lt, lb+1)
        ]
        self.assertEqual(expected_res, res)

    def test_read_page(self):
        res = self.fe.read_page(num=20)
        expect_res = [[f"({row}, {col})" for col in range(1, 21)] for row in range(1, 21)]
        self.assertEqual(expect_res, list(res))


if __name__ == "__main__":
    TestFromExcel.run()

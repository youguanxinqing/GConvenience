"""
TIME: 2019/11/24 21:52
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import pymysql
from .base import BaseConnector


class MysqlConnector(BaseConnector):
    def __init__(self, uri):
        super().__init__(uri)
        self._conn = pymysql.connect(self._server,
                                     self._user,
                                     self._passwd,
                                     self._db)
        self._cursor = self._conn.cursor()

    def execute(self, sql, args=None):
        if args:
            return self._cursor.execute(sql, args)
        return self._cursor.execute(sql)

    def commit(self):
        self._conn.commit()

    def close(self):
        self._cursor.close()
        self._conn.close()

    def extract_uri(self, uri):
        """
        uri = user|passwd|localhost|db
        """
        infos = uri.split("|")
        if len(infos) != 4:
            raise ValueError("currently, uri is supported by mysql")
        self._user, self._passwd,\
            self._server, self._db = infos

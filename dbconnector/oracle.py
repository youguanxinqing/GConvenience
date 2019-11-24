"""
TIME: 2019/11/24 22:20
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import cx_Oracle

from .base import BaseConnector


class OracleConnector(BaseConnector):
    def __init__(self, uri):
        super().__init__(uri)
        self._conn = cx_Oracle.connect(self._user,
                                       self._passwd,
                                       self._server)
        self._cursor = self._conn.cursor()
        self._cursor.execute()

    def execute(self, sql, args):
        if args:
            self._cursor.execute(sql, args)
        return self._cursor.execute()

    def commit(self):
        self._conn.commit()

    def close(self):
        self._cursor.close()
        self._conn.close()

    def extract_uri(self, uri):
        """
        uri = user|passwd|host
        eg: system|oracle|192.168.1.101/xe
        """
        infos = uri.split("|")
        if len(infos) != 3:
            raise ValueError("currently, uri is supported by oracle")

        self._user, self._passwd, self._server = infos

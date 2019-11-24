"""
TIME: 2019/11/24 20:50
Author: Guan
GitHub: https://github.com/youguanxinqing
EMAIL: youguanxinqing@qq.com
"""
import abc


class BaseConnector(metaclass=abc.ABCMeta):
    def __init__(self, uri):
        self._uri = uri
        self._user = None
        self._passwd = None
        self._server = None
        self._db = None

        self.extract_uri(self._uri)

        self._conn = None
        self._cursor = None

    @abc.abstractmethod
    def extract_uri(self, uri: str) -> str:
        pass

    @abc.abstractmethod
    def execute(self, sql, args):
        pass

    @abc.abstractmethod
    def close(self):
        pass

    def __exit__(self, exc_type, exc_val, exc_tb):
        self.close()

    def __enter__(self):
        return self

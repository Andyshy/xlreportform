# -*- coding:utf-8 -*-
from abc import ABCMeta, abstractmethod
import pymysql
import pymssql
__author__ = "Andy Yang"


class Database(metaclass=ABCMeta):
    @abstractmethod
    def connect(self):
        pass

    @abstractmethod
    def execute(self):
        pass


class MysqlClient(Database):
    def connect(self):
        print('okokoconne')

    def execute(self, condition=None):
        print('execute')


class SqlServerClient(Database):
    def connect(self):
        print('okokoconne')

    def execute(self, condition=None):
        print('execute')


if __name__ == '__main__':
    m = MysqlClient()
    print(m)
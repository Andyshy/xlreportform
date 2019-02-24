# -*- coding:utf-8 -*-
import time
from abc import ABCMeta, abstractmethod
from xlreportform.worksheet import WorkSheet

__author__ = "Andy Yang"


class Bases(metaclass=ABCMeta):
    def __init__(self):
        pass

    @abstractmethod
    def set_style(self):
        """set workshet's style, indent,border,font,and so on"""

    @abstractmethod
    def query(self):
        """query from mysql, sqlserver"""

    @abstractmethod
    def clean(self):
        """clean data"""

    @abstractmethod
    def export(self):
        """export data"""


class ReportForm(Bases, WorkSheet):
    def __init__(self, visible=False, filename=None, sheetname=None):
        WorkSheet.__init__(self, visible, filename, sheetname)

    def __new__(cls, *args, **kwargs):
        cls.query(cls)
        cls.clean(cls)
        cls.set_style(cls)
        cls.export(cls)
        return object.__new__(cls)


class DayRport(ReportForm):
    def query(self):
        print('query')
    def set_style(self):
        print('set_style')
    def export(self):
        print('export')


if __name__ == '__main__':
    d = DayRport(visible=True, filename='okok.xlsx', sheetname='dageda')
    time.sleep(5)
    print(d)
# -*- coding:utf-8 -*-
__author__ = "Andy Yang"
from win32com.client import Dispatch


class Application:
    """create excel application"""
    def __init__(self, visible=False):
        self.app = Dispatch('Excel.Application')
        self.app.Visible = visible
        self.app.DisplayAlerts = False

    def __repr__(self):
        """repr(Application)"""
        return "Excel <class '{0}'>".format(self.__class__.__name__)

    def __del__(self):
        """quit excel application when process finished"""
        self.app.Quit()


if __name__ == '__main__':
    w = Application()
    w()



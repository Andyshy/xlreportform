# -*- coding:utf-8 -*-
from win32com.client import Dispatch
__author__ = "Andy Yang"


class Application:
    """create excel application"""
    def __init__(self, visible=False):
        self._app = Dispatch('Excel.Application')
        self._app.Visible = visible
        self._app.DisplayAlerts = False

    def __repr__(self):
        """repr(Application)"""
        return "Excel <class '{0}'>".format(self.__class__.__name__)

    def __del__(self):
        """quit excel application when process finished"""
        self._app.Quit()


if __name__ == '__main__':
    w = Application()
    w()



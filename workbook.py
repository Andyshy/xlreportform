# -*- coding:utf-8 -*-
import os
from xlreportform.application import Application

__author__ = "Andy Yang"


class WorkBook(Application):
    """create excel's workbook"""
    def __init__(self, visible=False, filename=None):
        Application.__init__(self, visible)
        self._filename = filename
        self._workbook = self._create_book()

    def _create_book(self):
        """create a xlsx file if not exists, open it when exists"""
        if self._filename is None:
            raise ValueError('filename must a xlsx file string, not NoneType.')
        if not self._filename.endswith('.xlsx'):
            raise ValueError('filename is not a xlsx file or path with xlsx file.')
        # is not a file path
        if '\\' not in self._filename:
            current_dir = os.path.dirname(os.path.abspath(__file__))
            # current file's path
            filename = os.path.join(current_dir, self._filename)
            # if not exists
            if not os.path.exists(filename):
                # create a workbook
                workbook = self._app.Workbooks.Add()
                workbook.SaveAs(filename)
                return workbook
            # if exists,open it
            workbook = self._app.Workbooks.open(filename)
            return workbook
        # a file path, if not exists
        if not os.path.exists(self._filename):
            # create
            workbook = self._app.Workbooks.Add()
            workbook.SaveAs(self._filename)
            return workbook
        # open a workbook if exists
        workbook = self._app.Workbooks.open(self._filename)
        return workbook



if __name__ == '__main__':
    w = WorkBook(visible=True,filename=r'D:\WinterPlan\xlreportform\okok.xlsx')
    print(w)
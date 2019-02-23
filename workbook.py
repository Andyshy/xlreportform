# -*- coding:utf-8 -*-
__author__ = "Andy Yang"
import os
from xlreportform.application import Application


class WorkBook(Application):
    """create excel's workbook"""
    def __init__(self, visible=False, filename=None):
        Application.__init__(self, visible)
        self.filename = filename
        self.workbook = self.create_book()

    def create_book(self):
        """create a xlsx file if not exists, open it when exists"""
        if self.filename is None:
            raise ValueError('filename must a xlsx file string, not NoneType.')
        if not self.filename.endswith('.xlsx'):
            raise ValueError('filename is not a xlsx file or path with xlsx file.')
        if not os.path.exists(self.filename):
            workbook = self.app.Workbooks.Add()
            workbook.SaveAs(self.filename)
            return workbook
        workbook = self.app.Workbooks.open(self.filename)
        return workbook


if __name__ == '__main__':
    w = WorkBook(visible=True,filename=r'D:\WinterPlan\xlreportform\okok.xlsx')
    print(w)
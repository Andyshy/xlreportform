# -*- coding:utf-8 -*-
__author__ = "Andy Yang"
from xlreportform.workbook import WorkBook


class WorkSheet(WorkBook):
    """create worksheets"""
    def __init__(self, visible=False, filename=None, sheetname=None):
        WorkBook.__init__(self, visible, filename)
        self.sheetname = sheetname
        self.create_sheets()

    def create_sheets(self):
        """create worksheets"""
        if isinstance(self.sheetname, type(None)):
            raise ValueError("sheetname must a str or list")
        if isinstance(self.sheetname, str):
            self.app.Worksheets.Add().Name = self.sheetname
        if isinstance(self.sheetname, list):
            for name in self.sheetname:
                self.app.Worksheets.Add().Name = name


if __name__ == '__main__':
    w = WorkSheet(visible=True,filename=r'D:\WinterPlan\xlreportform\okok.xlsx', sheetname=['dada','00'])
    print(w)
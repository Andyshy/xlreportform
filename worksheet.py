# -*- coding:utf-8 -*-
from xlreportform.workbook import WorkBook

__author__ = "Andy Yang"


class WorkSheet(WorkBook):
    """create worksheets"""
    def __init__(self, visible=False, filename=None, sheetname=None):
        WorkBook.__init__(self, visible, filename)
        self.sheetname = sheetname
        self._create_sheets()

    def _create_sheets(self):
        """create worksheets, only sheet"""
        if isinstance(self.sheetname, type(None)):
            raise ValueError("sheetname must a str or list")
        if isinstance(self.sheetname, str):
            self._app.Worksheets.Add().Name = self.sheetname
        if isinstance(self.sheetname, list):
            for name in self.sheetname:
                self._app.Worksheets.Add().Name = name


if __name__ == '__main__':
    w = WorkSheet(visible=True,filename=r'okok.xlsx', sheetname=['dada','00'])
    print(w)
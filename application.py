# -*- coding:utf-8 -*-
import os
import platform
from win32com.client import Dispatch
__author__ = "Andy Yang"


class ExcelApplicationWrapper(object):
	def __init__(self):
		self._app = self._create_application()

	def _create_application(self):
		"""create excel application"""
		if not platform.system() in 'Windows':
			raise OSError('xlreportform not support other platform except Windows.')
		app = Dispatch('Excel.Application')
		return app

	def __del__(self):
		self._app.Quit


class Application(ExcelApplicationWrapper):
	"""create excel application"""
	def __init__(self, visible=False, filename=None, sheetname=None):
		super().__init__()
		self._visible = visible
		self._filename = filename
		self._sheetname = sheetname
		self._displayalerts()

	def __repr__(self):
		"""repr(Application)"""
		return "<class '{0}'>".format(self.__class__.__name__)

	def _visible(self):
		"""set visible
		0代表隐藏对象，但可以通过菜单再显示
		-1代表显示对象
		2代表隐藏对象，但不可以通过菜单显示，只能通过VBA修改为显示状态

		"""
		if not isinstance(self.visible, bool):
			raise ValueError('visible accept bool only. ')
		if not self.visible:
			self._app.Visible = 0
		else:
			self._app.Visible = -1

	def _displayalerts(self):
		"""ignore displayalerts"""
		self._app.DisplayAlerts = False

	def _add_book(self):
		return Workbook(_app=self._app, filename=self._filename).add()

	def _add_sheet(self, _book):
		return Worksheet(_book=_book, sheetname=self._sheetname).add()

	def add(self):
		self._WorkBook = self._add_book()
		return self._add_sheet(_book=self._WorkBook)

	def save(self):
		self._WorkBook.SaveAs(self._filename)
		self._WorkBook.save
		self._WorkBook.close


class Workbook:
	"""create excel's workbook"""
	def __init__(self, _app=None, filename=None):
		self._filename = filename
		self._app = _app

	def add(self):
		return self._create_book()

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
				return workbook
			# if exists,open it
			workbook = self._app.Workbooks.open(filename)
			return workbook
		# a file path, if not exists
		if not os.path.exists(self._filename):
			# create
			workbook = self._app.Workbooks.Add()
			return workbook
		# open a workbook if exists
		workbook = self._app.Workbooks.open(self._filename)
		return workbook


class Worksheet:
	"""create worksheets"""
	def __init__(self, _book=None, sheetname=None):
		self._sheetname = sheetname
		self._book = _book

	def add(self):
		return self._create_sheets()

	def _create_sheets(self):
		"""create worksheets, only sheet"""
		if isinstance(self._sheetname, type(None)):
			raise ValueError("sheetname must a str or list")
		if isinstance(self._sheetname, str):
			try:
				if not self._book.Worksheets[self._sheetname]:
					self._book.Worksheets.Add().Name = self._sheetname
			except Exception as e:
				if '发生意外。' in list(e.args):
					self._book.Worksheets.Add().Name = self._sheetname
		if isinstance(self._sheetname, list):
			for name in self._sheetname:
				try:
					if not self._book.Worksheets[name]:
						self._book.Worksheets.Add().Name = name
				except Exception as e:
					if '发生意外。' in list(e.args):
						self._book.Worksheets.Add().Name = name


if __name__ == '__main__':
	w = Application(
		visible=True, filename=r'D:\practice\xlreportform-master\nihao1.xlsx', sheetname='daa')
	w.add()
	w.save()




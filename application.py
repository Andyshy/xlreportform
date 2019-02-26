# -*- coding:utf-8 -*-
import os
import platform
from win32com.client import Dispatch
__author__ = "Andy Yang"


class ExcelApplicationWrapper(object):
	def __init__(self):
		self._app = self._create()

	def _create(self):
		"""create excel application"""
		if not platform.system() in 'Windows':
			raise OSError('xlreportform not support other platform except Windows.')
		app = Dispatch('Excel.Application')
		return app


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

	def __del__(self):
		"""quit excel application when process finished"""
		self._app.Quit()

	def _visible(self):
		"""set visible"""
		if not isinstance(self.visible, bool):
			raise ValueError('visible accept bool only. ')
		self._app.Visible = self.visible

	def _displayalerts(self):
		"""ignore displayalerts"""
		self._app.DisplayAlerts = False

	def _add_book(self):
		return WorkBook(_app=self._app, filename=self._filename).add()

	def _add_sheet(self):
		return WorkSheet(_app=self._app, sheetname=self._sheetname).add()

	def add(self):
		self._add_book()
		self._add_sheet()


class WorkBook:
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


class WorkSheet:
	"""create worksheets"""
	def __init__(self, _app=None, sheetname=None):
		self._sheetname = sheetname
		self._app = _app

	def add(self):
		self._create_sheets()

	def _create_sheets(self):
		"""create worksheets, only sheet"""
		if isinstance(self._sheetname, type(None)):
			raise ValueError("sheetname must a str or list")
		if isinstance(self._sheetname, str):
			self._app.Worksheets.Add().Name = self._sheetname
		if isinstance(self._sheetname, list):
			for name in self._sheetname:
				print(name)
				self._app.Worksheets.Add().Name = name


if __name__ == '__main__':
	w = Application(visible=False, filename='ni.xlsx', sheetname=['1ok', '3eqwe', '4dad'])
	w.add()
	print(w)



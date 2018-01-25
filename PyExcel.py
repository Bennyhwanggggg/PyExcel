'''
Authors: Benny Hwang 
Date: 25/01/2017
Version: 1.0

Use PyWin32 and win32com libary to interact with Excel from Python.


'''
from __future__ import print_function
import os, sys
import win32com.client

"""
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
								Methods
+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Initialisation: def __init__(self, workbook_path, show = False)
	Input: 
			- workbook_path = path to the excel file you want to open. e.g. os.path.join(path to excel file, excel file name).
			- visbility = False to let excel do work in the background, true otherwise.

		Usage: 
		PATH =os.path.dirname(os.path.realpath(__file__))
		example = os.path.join(PATH, "Example.xlsm")
		ExcelOperater(example, True)

		Expected result: This will open an excel file called Example.xlsm")

=======================================================================

Run macro: def RunMacro(self, macro)
	Input:
		- VBA macro name

	Usage:
	ExcelOperater(example).RunMacro("testmacro")

	Expected result: This will the example excel file and run the macro called testmacro

=======================================================================

Export PDFs: ExportAsPDF(self, target_sheets, output_name, output_location)
	Input:
		- A list of sheet names to be exported together
		- Output PDF name
		- Location of output file.

	Usage:
	ExcelOperater(example).ExportAsPDF("sheet1", test, example path)

	Expected Result: This will produce a pdf called test that contains information in sheet 1 of excel file example in example path

=======================================================================

Copy and Paste Values: CopyPasteAsValue(self, dest_sheet, src_book, src_sheet, CopyRange, PasteRange)
	Input:
		- src_sheet = sheet to copy from (either sheet number or sheet name)
		- dest_book = Excel workbook to copy to path directory e.g destination = os.path.join(PATH, "example.csv")
		- dest_sheet = Excel sheet to copy to (either sheet number or sheetname)
		- CopyRange and PasteRange = Range to copy to as strings. e.g "A1:B7"

	Usage:
	ExcelOperator(example1).CopyPastAsValue("sheet1", example2, "sheet2", "A1:A2", "B1:B2")

	Expected Result: This will copy values in sheet2 of example2 from A2 to A2 to example1's sheet1 at cells B1 to B2
	Special Behaviours:
		- If Copy Range is smaller than Paste Range, the value inside copy range will be repeated until it fills up the paste range.
		- If Paste Range is smaller than Copy Range, the later row/column will not be copied. (Not all data will be copied)

=======================================================================

Deleting Data: DeleteValues(self, wk_sheet, delRange)
			   DeleteCell(self, wk_sheet, delRange)
			   DeleteRowVal(self, wk_sheet, delRange)
			   DeleteRowCell(self, wk_sheet, delRange)
			   DeleteColumnsVal(self, wk_sheet, delRange)
			   DeleteColumnsCell(self, wk_sheet, delRange)
	Input: (for all following delete functions)
		- delRnage = Cells to delete as stirng e.g "A1:B7"
		- wk_sheet = worksheet to delete from

	Usage:
	ExcelOperator(example1).DeleteValues("sheet1","A1:A2")

	Expected Result: This will delete the values from A1 to A2 in sheet1 of example1. Does not delete formulas.
	Similar behaviour for other delete functions.

=======================================================================

Inserting Data: insertVal(self, wk_sheet, insertRange, data)
	Input:
		- wk_sheet = worksheet to insert
		- insertRange = location to insert
		- data = value to insert

	Usage:
	ExcelOperator(example1).insertVal("sheet1","A1:A1", 'abcd')

	Expected Result: This will insert 'abcd' to the cell A1 in sheet1 of example 1
	Special Behvaiours: 
		- If a list is given as a data, and assuming the range to insert is compatible with the size of that list, this values of the list will be inserted accross that row.
		- If the range  of insert is given accross serveral rows, the data will be inserted accross those rows. If data is a list, the list will be repeated along those rows.

======================================================================

Getting a cell value: GetVal(self, wk_sheet, row, col)
	Input:
		- wk_sheet = worksheet to get value from
		- row, col = cell number. e.g A1 = 1, 1

	Usage:
	ExcelOperator(example1).GetVal("sheet1", 1, 1)

	Expected Result: This will insert return the value in cell A1 in sheet1 of example 1

======================================================================

Other Functions:
	QuitExcel(self)
	RefreshCalculation(self)
	CloseWorkBook(self, save = False):
	AddWorkSheet(self, sheetname):
	QuitExcel(self):
	RefreshCalculation(self):
	MakeVisible(self):
	Hide(self):


"""

class ExcelOperaterError(Exception):
	def __init__(self, message):
		self.message = message

class ExcelOperater:
	'''
	Input: 
		- workbook_path = os.path.join(path to excel file, excel file name)
		- visbility = False to let excel do work in the background, true otherwise.

	Usage: 
	PATH =os.path.dirname(os.path.realpath(__file__))
	example = os.path.join(PATH, "Example.xlsm")
	ExcelOperater(example, True)

	Expected result: This will open an excel file called Example.xlsm")
	'''

	def __init__(self, workbook_path, show = False):
		self.excelObj = win32com.client.Dispatch("Excel.Application")
		self.show = show
		self.excelObj.Visible = show
		self.workbook_path = workbook_path
		self.workBookObj = self.excelObj.Workbooks.Open(workbook_path)


	'''
	Input:
		- VBA macro name

	Usage:
	ExcelOperater(example).RunMacro("testmacro")

	Expected result: This will the example excel file and run the macro called testmacro
	'''
	def RunMacro(self, macro):
		try:
			print("Running macro: {}".format(macro))
			self.excelObj.Run(macro)
			return True
		except:
			raise ExcelOperaterError('The macro you are trying to run does not exist or the workbook is not macro enabled.')

	'''
	Input:
		- A list of sheet names to be exported together
		- Output PDF name
		- Location of output file.

	Usage:
	ExcelOperater(example).ExportAsPDF("sheet1", test, example path)

	Expected Result: This will produce a pdf called test that contains information in sheet 1 of excel file example in example path
	'''
	def ExportAsPDF(self, target_sheets, output_name, output_location):
		try:
			print("Exporting the following sheets:")
			if(isinstance(target_sheets, list)):
				if all(isinstance(target_sheets, int)):
					for shtNum in target_sheets:
						print("Sheet Number:", shtNum)
				else:
					print("\n".join(target_sheets))
			else:
				if isinstance(target_sheets, int):
					print("Sheet Number:", target_sheets)
				else:
					print(target_sheets)

			self.excelObj.WorkSheets(target_sheets).Select()
			output = os.path.join(output_location, output_name)
			self.excelObj.ActiveSheet.ExportAsFixedFormat(0, output)
			return True
		except:
			raise ExcelOperaterError("Error occured while exporting sheets as PDFs.")

	'''
	Input:
		- src_sheet = sheet to copy from (either sheet number or sheet name)
		- dest_book = Excel workbook to copy to path directory e.g destination = os.path.join(PATH, "example.csv")
		- dest_sheet = Excel sheet to copy to (either sheet number or sheetname)
		- CopyRange and PasteRange = Range to copy to as strings. e.g "A1:B7"

	Usage:
	ExcelOperator(example1).CopyPastAsValue("sheet1", example2, "sheet2", "A1:A2", "B1:B2")

	Expected Result: This will copy values in sheet2 of example2 from A2 to A2 to example1's sheet1 at cells B1 to B2
	Special Behaviours:
		- If Copy Range is smaller than Paste Range, the value inside copy range will be repeated until it fills up the paste range.
		- If Paste Range is smaller than Copy Range, the later row/column will not be copied. (Not all data will be copied)
	'''
	def CopyPasteAsValue(self, dest_sheet, src_book, src_sheet, CopyRange, PasteRange):
		try:
			print("Copying from: \n\t{}\n\tSheet: {}\n\t\tto\n\t{}\n\tSheet: {}".format(src_book, src_sheet, self.workbook_path, dest_sheet))
			print("Copy Range is: {}\nPasteRange is: {}\nPlease make sure the two range size match manually, otherwise not all data will be copied".format(CopyRange, PasteRange))
			
			dest = self._openSheet(dest_sheet)
			target = None
			# If not copying data within the same workbook, open the src data work book and the relevant sheet.
			if src_book != self.workbook_path:
				target = ExcelOperater(src_book, self.show)
				src = target._openSheet(src_sheet)
			else:
				# if copying data in the same workbook but different sheet
				if src_sheet != dest_sheet:
					src = self._openSheet(src_sheet)
				# if copying within the same sheet
				else:
					src = dest

			dest.Range(PasteRange).Value = src.Range(CopyRange).Value

			# Close the newly opened workbook if any
			if target:
				target.CloseWorkBook()

		except:
			raise ExcelOperaterError("An error occured while copying data. Target or destination sheets may not exist.")

	'''
	Input: (for all following delete functions)
		- delRnage = Cells to delete as stirng e.g "A1:B7"
		- wk_sheet = worksheet to delete from

	Usage:
	ExcelOperator(example1).DeleteValues("sheet1","A1:A2")

	Expected Result: This will delete the values from A1 to A2 in sheet1 of example1. Does not delete formulas.
	Similar behaviour for other delete functions.
	'''
	def DeleteValues(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).Value = ''
		except:
			raise ExcelOperaterError("An error occured while deleting data. Target or destination sheets may not exist.")

	def DeleteCell(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).Delete()
		except:
			raise ExcelOperaterError("An error occured while deleting single cells. Target or destination sheets may not exist.")

	def DeleteRowVal(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).EntireRow.Value = ''
		except:
			raise ExcelOperaterError("An error occured while deleting rows values. Target or destination sheets may not exist.")

	def DeleteRowCell(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).EntireRow.Delete()
		except:
			raise ExcelOperaterError("An error occured while deleting rows cells. Target or destination sheets may not exist.")

	def DeleteColumnsVal(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).EntireColumn.Value = ''
		except:
			raise ExcelOperaterError("An error occured while deleting column values. Target or destination sheets may not exist.")
	
	def DeleteColumnsCell(self, wk_sheet, delRange):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(delRange).EntireColumn.Delete()
		except:
			raise ExcelOperaterError("An error occured while deleting column cells. Target or destination sheets may not exist.")

	'''
	Input:
		- wk_sheet = worksheet to insert
		- insertRange = location to insert
		- data = value to insert

	Usage:
	ExcelOperator(example1).insertVal("sheet1","A1:A1", 'abcd')

	Expected Result: This will insert 'abcd' to the cell A1 in sheet1 of example 1
	Special Behvaiours: 
		- If a list is given as a data, and assuming the range to insert is compatible with the size of that list, this values of the list will be inserted accross that row.
		- If the range  of insert is given accross serveral rows, the data will be inserted accross those rows. If data is a list, the list will be repeated along those rows.
	'''
	def insertVal(self, wk_sheet, insertRange, data):
		try:
			src = self._openSheet(wk_sheet)
			src.Range(insertRange).Value = data
		except:
			raise ExcelOperaterError("An error occured while inserting values. Target or destination sheets may not exist.")
	'''
	Input:
		- wk_sheet = worksheet to get value from
		- row, col = cell number. e.g A1 = 1, 1

	Usage:
	ExcelOperator(example1).GetVal("sheet1", 1, 1)

	Expected Result: This will insert return the value in cell A1 in sheet1 of example 1
	'''

	def GetVal(self, wk_sheet, row, col):
		try:
			src = self._openSheet(wk_sheet)
			return src.Cells(row,col).Value
		except:
			raise ExcelOperaterError("An error occured while getting values. Target sheet may not exist or invalid row or column number.")

	'''
	Other useful functions
	'''
	def CloseWorkBook(self, save = False):
		try:
			self.workBookObj.Close(save)
			return True
		except:
			ExcelOperaterError('Errorered occured while closing excel workbook.')

	def AddWorkSheet(self, sheetname):
		try:
			newWS = self.workBookObj.Worksheets.Add()
			newWS.NAME = sheetname
		except:
			ExcelOperaterError('Errorered occured while adding excel worksheet.')

	def QuitExcel(self):
		self.excelObj.Quit()
		return

	def RefreshCalculation(self):
		self.workBookObj.Application.Calculate()
		return

	def MakeVisible(self):
		self.excelObj.Visible = True

	def Hide(self):
		self.excelObj.Visible = False



	'''
	Helper functions
	'''
	def _openSheet(self, sheet_to_open):
		wkBook = self.workBookObj
		if isinstance(sheet_to_open, int):
			sht = wkBook.WorkSheets[sheet_to_open]
		else:
			sht = wkBook.WorkSheets(sheet_to_open)
		return sht

	def _getWorkBookObj(self):
		return self.workBookObj

	



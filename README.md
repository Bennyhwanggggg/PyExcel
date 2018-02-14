# PyExcel

This script allow users to use python to interact with Microsoft Excel. 

+++++++++++++++++++++++
	Methods
+++++++++++++++++++++++

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
		       CopyPasteEntireCol(self, dest_sheet, src_book, src_sheet, CopyRange, PasteRange) (will keep formatting)
	    	       CopyPasteEntireRow(self, dest_sheet, src_book, src_sheet, CopyRange, PasteRange) (will keep formatting)


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

Setting font and cell colour: setFontColor(self, wk_sheet, row, col, colour)
			      highlightCell(self, wk_sheet, row, col, colour)

	Input:
		- wk_sheet = worksheet of the cell you want to modify
		- row, col = row and column number  e.g A1 = 1, 1
		- colour to change to. See avaialble colours at the function and adjust accordingly.

	Usage:
	ExcelOperator(example1).setFontColor("sheet1", 1, 1, 3)

	Expected Result: This will change cell A1 in sheet1's colour to red




Other Functions:
	QuitExcel(self)
	RefreshCalculation(self)
	CloseWorkBook(self, save = False):
	AddWorkSheet(self, sheetname):
	QuitExcel(self):
	RefreshCalculation(self):
	MakeVisible(self):
	Hide(self):
  	Save(self)
	SaveAs(self, fileName_With_Location)
	NewWorkbook(self, fileName_With_Location = "new.csv")
	turnAlerts(self, alertStatus)
  
 # Future Developments
 - Ability to add and format charts


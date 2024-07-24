'automatically create Microsoft Excel file (.xlsx)

'create excel object
	Set objExcel = CreateObject("Excel.Application") 

'view excel program and file (set to false to hide process)
	objExcel.Visible = True 

'add new workbook
	Set objWorkbook = objExcel.Workbooks.Add 

'set a cell value at row 3 column 5
	objExcel.Cells(3,5).Value = "test"

'change a cell value
	objExcel.Cells(3,5).Value = "something different"
	
'delete a cell value
	objExcel.Cells(3,5).Value = ""

'get a cell value and set it equal to a variable
	r3c5 = objExcel.Cells(3,5).Value

'save the new excel file (make sure to change the location)
	objWorkbook.SaveAs "C:\Users\jeffl\Documents\vbsTest.xlsx" 

'close workbook
	objWorkbook.Close 

'exit excel program
	objExcel.Quit

'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing
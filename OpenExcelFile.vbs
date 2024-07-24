'automatically open and edit Microsoft Excel file (.xlsx)

'create excel object
	Set objExcel = CreateObject("Excel.Application") 

'view excel program and file (set to false to hide process)
	objExcel.Visible = True 

'open excel file (make sure to change the location)
	Set objWorkbook = objExcel.Workbooks.Open("C:\Users\jeffl\Documents\vbsTest.xlsx")

'set a cell value at row 3 column 5
	objExcel.Cells(3,5).Value = "new value"

'save the existing excel file (use SaveAs to save it as something else)
	objWorkbook.Save

'close workbook
	objWorkbook.Close 

'exit excel program
	objExcel.Quit

'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing
Set objExcel = CreateObject("Excel.Application")

' select Excel file to open
Set objWorkbook = objExcel.Workbooks.Open("C:\Users\jeffl\Documents\vbaMacroTest.xlsm")

objExcel.Application.Visible = True

' specify which macro in workbook to run
objExcel.Application.Run "vbaMacroTest.xlsm!AddCheckBoxes" 

objExcel.ActiveWorkbook.Close

objExcel.Application.Quit

WScript.Echo "Finished."
WScript.Quit
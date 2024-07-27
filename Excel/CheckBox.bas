Attribute VB_Name = "AddCheckBox" ' remove this line before running macro in Excel
Sub AddCheckBoxes()

    ' abbreviate definitions 
    Dim cb As CheckBox
    Dim myRange As Range, cel As Range
    Dim wks As Worksheet

    ' set worksheet and range
    Set wks = Sheets("Sheet1")
    Set myRange = wks.Range("A1:A100")

    ' loop for adding checkboxes for each cell in range
    For Each cel In myRange

        Set cb = wks.CheckBoxes.Add(cel.Left, cel.Top, 30, 6)

        With cb
            .Caption = ""
            .OnAction = ""
        End With

    Next

End Sub
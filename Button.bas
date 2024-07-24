Attribute VB_Name = "Button"
Sub AddButton()

' add a command button to current document
Set doc = ThisDocument
' Dim doc As Word.Document
Dim shp As Word.InlineShape
' Set doc = Documents.Add

Set shp = doc.Content.InlineShapes.AddOLEControl(ClassType:="Forms.CommandButton.1")
shp.OLEFormat.Object.Caption = "Click"

' add a procedure for the click event of the inlineshape
'**Note: The click event resides in the This Document module
Dim sCode As String
sCode = "Private Sub " & shp.OLEFormat.Object.Name & "_Click()" & vbCrLf & _
        "   MsgBox ""CommandButton Clicked!""" & vbCrLf & _
        "End Sub"
doc.VBProject.VBComponents("ThisDocument").CodeModule.AddFromString sCode

End Sub
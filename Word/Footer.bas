Attribute VB_Name = "AddFooter"

Sub CreateFooter()

Dim wd As Word.Application
Dim doc As Word.Document

Dim ftr As Word.HeaderFooter

Set wd = New Word.Application

wd.Visible = True

' new document with header created for testing purposes
Set doc = wd.Documents.Add

Set ftr = doc.Sections(1).Footers(wdHeaderFooterPrimary)

With ftr.Range
    .Text = "Test Footer"
    .Font.Size = 14
    .Bold = True
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
End With

End Sub
Attribute VB_Name = "AddHeader"

Option Explicit

Sub CreateHeader()

Dim wd As Word.Application
Dim doc As Word.Document

Dim hdr As Word.HeaderFooter

Set wd = New Word.Application

wd.Visible = True

' new document with header created for testing purposes
Set doc = wd.Documents.Add

Set hdr = doc.Sections(1).Headers(wdHeaderFooterPrimary)

With hdr.Range
    .Text = "Test Header"
    .Font.Size = 14
    .Bold = True
    .ParagraphFormat.Alignment = wdAlignParagraphCenter
End With

End Sub



Attribute VB_Name = "Módulo1"
Sub ReplaceTextInWord()

Dim wdApp As Word.Application
Dim wdDoc As Word.Document
Dim wdRange As Word.Range
Dim xlApp As Excel.Application
Dim xlWb As Excel.Workbook
Dim xlSh As Excel.Worksheet
Dim LastRow As Long
Dim CurrentName As String

Set xlApp = New Excel.Application
Set xlWb = xlApp.Workbooks.Open("C:\Users\daniel.farias\Documents\Recon\PruebaPrefe\base-htls-ratescore.xlsx") 'path base nomes hotéis
Set xlSh = xlWb.Sheets("Planilha1")

LastRow = xlSh.Cells(xlSh.Rows.Count, "A").End(xlUp).Row

For i = 2 To LastRow

CurrentName = xlSh.Cells(i, 1).Value

CurrentRatescore = xlSh.Cells(i, 2).Value 'teste //pega ratescore

Set wdApp = New Word.Application
Set wdDoc = wdApp.Documents.Open("C:\Users\daniel.farias\Documents\Recon\PruebaPrefe\modelo-prefe.docx") 'template
Set wdRange = wdDoc.Content

Set wdRange = wdDoc.Shapes("Rectangle 2").TextFrame.TextRange 'referência forma, pegar com debug
With wdRange.Find
    .Text = "hotel"
    .Replacement.Text = CurrentName
    .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindStop
End With

Set wdRange = wdDoc.Shapes("Rectangle 3").TextFrame.TextRange
With wdRange.Find
    .Text = "0"
    .Replacement.Text = CurrentRatescore
    .Execute Replace:=wdReplaceAll, Forward:=True, Wrap:=wdFindStop
End With


'Debug.Print wdDoc.Name
'Debug.Print wdRange.Find.MatchCase
'Debug.Print CurrentName
'Debug.Print wdRange.Text
'Debug.Print wdDoc.Shapes(2).Name


wdDoc.SaveAs "C:\Users\daniel.farias\Documents\Recon\PruebaPrefe\Samples\" & CurrentName & ".pdf", wdFormatPDF
wdDoc.SaveAs "C:\Users\daniel.farias\Documents\Recon\PruebaPrefe\Samples\" & CurrentName & ".docx"

wdDoc.Close
wdApp.Quit

Next i

xlWb.Close
xlApp.Quit

End Sub


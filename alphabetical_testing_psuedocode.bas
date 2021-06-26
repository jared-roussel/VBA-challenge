Attribute VB_Name = "Module1"
Sub SingleSheetTesting():

'Add new columns to the worksheet (Ticker, Yearly Change, Percent Change, Total Stock Volume)

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Value"

'Dim ws As Worksheet

Set ws = Worksheets("Sheet1")


'Find the last row in the worksheet
'Dim LastRow As String
LastRow = ActiveSheet.Cells(Row.Count, 1).End(xlUp).Row
'MsgBox (LastRow)


End Sub

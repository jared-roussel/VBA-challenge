Attribute VB_Name = "Module1"
Sub SingleSheetTesting():

'Add new columns to the worksheet (Ticker, Yearly Change, Percent Change, Total Stock Volume)

Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Value"

'Dim ws As Worksheet
' --------------------------
'Set ws = Worksheets("Sheet1")
'---------------------------

'Find the last row in the worksheet
'Dim LastRow As String
'LastRow = ActiveSheet.Cells(Row.Count, 1).End(xlUp).Row
'MsgBox (LastRow)
'-----------------------------

'Set up loop to go through the tickers and get value

Dim ticker As String
Dim ticker_next As String


ticker = Cells(2, 1).Value
'MsgBox (ticker)

For i = 2 To 264
ticker = Cells(i, 1).Value
ticker_next = Cells(i + 1, 1).Value

If ticker = ticker_next Then
  ticker = Cells(i + 1, 1).Value
  ticker_next = Cells((i + 1) + 1, 1).Value
  
ElseIf ticker <> ticker_next Then
Cells(2, 9).Value = ticker
  End If
Next i

End Sub

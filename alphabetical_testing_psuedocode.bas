Attribute VB_Name = "Module1"
Sub Test():

'Part A: Add new columns to the worksheet (Ticker, Yearly Change, Percent Change, Total Stock Volume)
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"


'Dim wb As Workbook
'Dim ws As Worksheet

'Set ws = Sheet("Sheet1")






'Set up loop to go through the tickers and get value

'Dim ws As Worksheets

'For Each ws In Worksheets
'ws.Activate

'Declare variables
Dim ticker As String
Dim ticker_next As String
Dim Summary_Row As Integer
Dim open_value As Double
Dim close_value As Double
Dim total_volume As Double
Dim LastRow As Double

'Find the last row in the worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (LastRow)

'Initialize Summary_Row index
Summary_Row = 2

'Initialize ticker, open_value, total_volume
ticker = Cells(2, 1).Value
open_value = Cells(2, 3).Value
total_volume = Cells(2, 7).Value

'MsgBox (ticker)

'Set up loop for column A for ticker
For i = 2 To LastRow
    ticker = Cells(i, 1).Value
    ticker_next = Cells(i + 1, 1).Value

'Set conditionals for grouping ticker information while ticker and next ticker are equal
If ticker = ticker_next Then
    ticker = Cells(i + 1, 1).Value
    ticker_next = Cells((i + 1) + 1, 1).Value
    total_volume = total_volume + Cells(i, 7).Value
  
'Set conditionals for populating the summary table once the ticker and next ticker do not equal
ElseIf ticker <> ticker_next Then
    Range("I" & Summary_Row).Value = ticker

'Reset values for next conditional pass
    close_value = Cells(i - 1, 6).Value
    total_volue = total_volume

'Update summary table values
Range("J" & Summary_Row).Value = (open_value - close_value)
Range("K" & Summary_Row).Value = ((open_value - close_value) / open_value)
Range("L" & Summary_Row).Value = total_volume


'Increase Summary_Row
Summary_Row = Summary_Row + 1
  End If

'Reset values for next loop
  ticker = ticker_next
  open_value = Cells(i, 3).Value
  



Next i

'Reset the volume for the next i
total_volume = 0

End Sub

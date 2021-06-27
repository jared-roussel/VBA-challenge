Attribute VB_Name = "Module1"
Sub TickerLoop():

'Set workbook and worksheet variables and loops
Dim ws As Worksheet
For Each ws In Worksheets

Activate Worksheet
ws.Activate


'Add new columns to the worksheet (Ticker, Yearly Change, Percent Change, Total Stock Volume)
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Declare variables
Dim ticker As String
Dim ticker_next As String
Dim Summary_Row As Integer
Dim open_value As Double
Dim close_value As Double
Dim total_volume As Double
Dim LastRow As Double
Dim yearly_change As Double
Dim percent_change As Double

'Find the last row in the worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (LastRow)

'Initialize Summary_Row index
Summary_Row = 2

'Initialize ticker, open_value, total_volume
'ticker = Cells(2, 1).Value
open_value = Cells(2, 3).Value
total_volume = 0

'MsgBox (ticker)

'Set up loop for column A for ticker
For i = 2 To LastRow
    ticker = Cells(i, 1).Value
    ticker_next = Cells(i + 1, 1).Value


    'Set conditionals for grouping ticker information while ticker and next ticker are equal
    If ticker = ticker_next Then
        'ticker = Cells(i + 1, 1).Value
        'ticker_next = Cells((i + 1) + 1, 1).Value
        total_volume = total_volume + Cells(i, 7).Value
      
    'Set conditionals for populating the summary table once the ticker and next ticker do not equal
    Else
        Range("I" & Summary_Row).Value = ticker
    
        'Reset values for next conditional pass
        close_value = Cells(i, 6).Value
        total_volume = total_volume + Cells(i, 7).Value
    
        'Update summary table values
        yearly_change = (close_value - open_value)
        
        'Nest If for evaluating the open value
        If open_value > 0 And close_value > 0 Then
            percent_change = (((close_value - open_value) / open_value)) * 100
        Else
            percent_change = 0
        End If
        
        Range("J" & Summary_Row).Value = yearly_change
        Range("K" & Summary_Row).Value = percent_change
        Range("L" & Summary_Row).Value = total_volume
    
    
        'Increase Summary_Row
        Summary_Row = Summary_Row + 1
        
        'Reset the volume and open_value for the next i
        total_volume = 0
        open_value = Cells(i + 1, 3).Value
        
    End If

    'Reset values for next loop
    ticker = ticker_next
    

Next i


Next ws


End Sub

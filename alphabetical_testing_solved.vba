Attribute VB_Name = "Module1"
Sub TickerLoop():

'Set workbook and worksheet variables and loops
Dim ws As Worksheet
For Each ws In Worksheets

'Activate Worksheet
ws.Activate


'Add new columns to the worksheet (Ticker, Yearly Change, Percent Change, Total Stock Volume)
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

'Declare variables
Dim Ticker As String
Dim ticker_next As String
Dim Summary_Row As Integer
Dim open_value As Double
Dim close_value As Double
Dim total_volume As Double
Dim LastRow As Double
Dim yearly_change As Double
Dim percent_change As Double
'Dim LastSummaryRow As Double
Dim max_percent_change As Double
Dim min_percent_change As Double
Dim max_total_volume As Double


'Find the last row in the worksheet
LastRow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox (LastRow)

'Initialize Summary_Row index
Summary_Row = 2

'Initialize ticker, open_value, total_volume
open_value = Cells(2, 3).Value
total_volume = 0

'MsgBox (ticker)

'Set up loop for column A for ticker
For I = 2 To LastRow
    Ticker = Cells(I, 1).Value
    ticker_next = Cells(I + 1, 1).Value


    'Set conditionals for grouping ticker information while ticker and next ticker are equal
    If Ticker = ticker_next Then
        total_volume = total_volume + Cells(I, 7).Value
      
    'Set conditionals for populating the summary table once the ticker and next ticker do not equal
    Else
        Range("I" & Summary_Row).Value = Ticker
    
        'Reset values for next conditional pass
        close_value = Cells(I, 6).Value
        total_volume = total_volume + Cells(I, 7).Value
    
        'Update summary table values
        yearly_change = (close_value - open_value)
        
        'Nest If for evaluating the open value
        If open_value > 0 And close_value > 0 Then
            percent_change = (((close_value - open_value) / open_value))
        Else
            percent_change = open_value
        End If
        
        Range("J" & Summary_Row).Value = yearly_change
        Range("K" & Summary_Row).Value = percent_change
        Range("L" & Summary_Row).Value = total_volume
    
    'Conditionally format the Percent Change
    If Range("K" & Summary_Row).Value <= 0 Then
        Range("K" & Summary_Row).Interior.Color = vbRed
    Else
        Range("K" & Summary_Row).Interior.Color = vbGreen
    End If
    
        'Increase Summary_Row
        Summary_Row = Summary_Row + 1
        
        'Reset the volume and open_value for the next i
        total_volume = 0
        open_value = Cells(I + 1, 3).Value
        
    End If

    'Reset values for next loop
    Ticker = ticker_next
        

Next I

'Challenge Code for greatest % increase, % decrease and greatest volume

Range("O1").Value = "Ticker"
Range("P1").Value = "Value"
Range("N2").Value = "Greatest % Increase"
Range("N3").Value = "Greatest % Decrease"
Range("N4").Value = "Greatest Total Volume"

max_percent_change = WorksheetFunction.Max(Range("K:K"))
Range("P2").Value = max_percent_change
min_percent_change = WorksheetFunction.Min(Range("K:K"))
Range("P3").Value = min_percent_change
max_total_volume = WorksheetFunction.Max(Range("L:L"))
Range("P4").Value = max_total_volume

'---- Attempt to find ticker value for each max, min value
'LastSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row
'Dim j As Integer
'j = Range("I2:I & LastSummaryRow")

'For j = 2 To LastSummaryRow
    'If max_percent_change = Cells(j,





Next ws


End Sub


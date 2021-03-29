Attribute VB_Name = "Module1"
Sub vba_stocks()

'Step 1: Loop through multiple worksheets and collect data on single sheet
'Step 2: Define variables and create Summary table for ticker symbol, yearly change,
'percent Change & Total volume
'Step 3: create calculations for each column in summary table /end loops
'Step 4: Add conditional formatting to yearly change

'Step 1:
'walk through each worksheet
Dim ws As Worksheet

For Each ws In ThisWorkbook.Worksheets
ws.Activate

'define last row
Dim Lastrow As Long
Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
'MsgBox ("This is the last row")

'Step 2:
'Set initial variables
Dim i As Long
Dim open_price As Double
Range("C2") = open_price
Dim close_price As Double

'variables for summary table
Dim Ticker As String
Dim Yearly_change As Double
Dim Percent_change As Double
Dim Stock_vol As Double

'Step 3:
'set the total stock volume to zero
Stock_vol = 0

'create a summary table for the information
Dim Summary_table_row As Integer
Summary_table_row = 2

'place variables in summary table
Range("I1").Value = "Ticker"
Range("J1").Value = "Yearly change"
Range("K1").Value = "Percent change"
Range("L1").Value = "Stock volume"

'loop through the tickers
For i = 2 To Lastrow
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    Ticker = Cells(i, 1).Value
    Range("I" & Summary_table_row).Value = Ticker

'Find Yearly change
close_price = Cells(i, 6).Value
Yearly_change = close_price - open_price
Range("J" & Summary_table_row).Value = Yearly_change

'Find Percent change
    If open_price = 0 Then
    Percent_change = 1
    
    Else
    
    Percent_change = (Yearly_change / open_price)
    End If
    
'output data onto summary table
Range("K" & Summary_table_row).Value = Percent_change
    
'find total stock volume
Stock_vol = Stock_vol + Cells(i, 7).Value
Range("L" & Summary_table_row).Value = Stock_vol

'reset the total stock volume
Stock_vol = 0
open_price = Cells(i + 1, 3).Value

'add one to the summary table row
Summary_table_row = Summary_table_row + 1

'if the cell following immediately after is the same ticker
Else

'add to the stock volume
Stock_vol = Stock_vol = Cells(i, 7).Value

End If
Next i
Next ws

'Step 4:
'Conditional  Formatting

'define variables
Dim SumYC As Range
Dim condition_positive As FormatCondition
Dim condition_negative As FormatCondition
Set SumYC = Range("J2", "J" & (Summary_table_row - 1))

'make the rule for the conditional format
Set condition_positive = SumYC.FormatConditions.Add(xlCellValue, xlGreater, "0")
Set condition_negative = SumYC.FormatConditions.Add(xlCellValue, xlLess, "1")

'add color to formatting
With condition_positive
    .Interior.Color = vbGreen
    .Font.Color = vbBlack
End With

With condition_negative
    .Interior.Color = vbRed
    .Font.Color = vbBlack
End With

Range("K1", "K2836").NumberFormat = "0.00%"
Cells.EntireColumn.AutoFit


End Sub

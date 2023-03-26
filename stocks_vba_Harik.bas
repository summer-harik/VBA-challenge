Attribute VB_Name = "Module1"
Sub stocks()

'loop thru all worksheets
For Each ws In Worksheets

'set header names for summary table
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Volume"

'create/store summary_table_row variable
Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

'determine and create/store lastrow variable
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

'create/store variables
Dim ticker As String
Dim yearly_change As Double
Dim open_price As Double
Dim closing_price As Double
Dim percent_change As Double
Dim volume As LongLong
Dim total_volume As LongLong

'set an initial variable for holding the total volume
total_volume = 0

'set initial open price so that first open price of year is always used *found on stackoverflow*
open_price = ws.Cells(2, 3).Value

'loop thru all rows
For i = 2 To lastrow

'check if we are within the same ticker
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'set the ticker name
ticker = ws.Cells(i, 1).Value

'set the volume
volume = ws.Cells(i, 7).Value

'set the closing price
closing_price = ws.Cells(i, 6).Value

'calculate the yearly change in price
yearly_change = closing_price - open_price

'calculate the percentage change in price
percent_change = (yearly_change / open_price)

'calculate the total volume
total_volume = total_volume + volume

'print the ticker name in the summary table
ws.Range("I" & Summary_Table_Row).Value = ticker

'print the yearly change value in the summary table
ws.Range("J" & Summary_Table_Row).Value = yearly_change
'format the yearly change cell color
If yearly_change > 0 Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
    Else: ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
    
'print the percent change value in the summary table
ws.Range("K" & Summary_Table_Row).Value = percent_change
'format the percent change cell color
 If percent_change > 0 Then
    ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
    Else: ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
    End If
'format the percent change column as a percentage *found on stackoverflow*
ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

'print the total volume value in the summary table
ws.Range("L" & Summary_Table_Row).Value = total_volume

'add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

'reset the total volume
total_volume = 0

'get the next open price value
open_price = ws.Cells(i + 1, 3)

'set row names
ws.Range("N2").Value = "Greatest % increase"
ws.Range("N3").Value = "Greatest % decrease"
ws.Range("N4").Value = "Greatest total volume"


'greatest percent increase *function found on stackoverflow*
max_percent = Application.WorksheetFunction.Max(ws.Range("K:K"))

'greatest percent decrease
min_percent = Application.WorksheetFunction.Min(ws.Range("K:K"))

'greatest total volume
max_volume = Application.WorksheetFunction.Max(ws.Range("L:L"))

'print greatest values
ws.Range("O2").Value = max_percent
ws.Range("O3").Value = min_percent
ws.Range("O4").Value = max_volume

Else

'add to the total volume
total_volume = total_volume + volume

End If

Next i
Next ws

End Sub


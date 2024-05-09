Attribute VB_Name = "Module1"
Sub summary_table():

Dim sumtable_row_ticker, sumtable_row_qtr, sumtable_row_vol As Integer
Dim i, k As Integer
Dim ticker As String
Dim open_price As Double
Dim close_price As Double
Dim vol As LongLong
Dim lastrow, lastrow_sum As Long

sumtable_row_ticker = 2
sumtable_row_qtr = 2
sumtable_row_vol = 2
lastrow = Cells(Rows.Count, 1).End(xlUp).Row
lastrow_sum = Cells(Rows.Count, 9).End(xlUp).Row
vol = 0

For Each ws In ActiveWorkbook.Worksheets

'Add titles to summary table
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Quarterly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

'Add titles to summary table (greatest values)
    ws.Cells(1, 15).Value = "Ticker"
    ws.Cells(1, 16).Value = "Value"
    ws.Cells(2, 14).Value = "Greatest % Increase"
    ws.Cells(3, 14).Value = "Greatest % Decrease"
    ws.Cells(4, 14).Value = "Greatest Total Volume"

'Pull ticker symbol
    'Loop through <ticker> column
        For i = 2 To lastrow
        'Pull unique ticker symbols
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'Output to summary table
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & sumtable_row_ticker).Value = ticker
        sumtable_row_ticker = sumtable_row_ticker + 1
        End If
        Next i

'Quarterly and percent change
    For i = 2 To lastrow
    'Calculate quarterly change
        'Pull value for opening qtr price
        If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
            open_price = ws.Cells(i, 3).Value
        End If
    'pull value for closing qtr price
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            close_price = ws.Cells(i, 6).Value
        'calc qtrly change (close - open price)and print to summary table
            ws.Range("J" & sumtable_row_qtr).Value = close_price - open_price
            'Conditional formatting
                If ws.Range("j" & sumtable_row_qtr).Value > 0 Then
                    ws.Range("j" & sumtable_row_qtr).Interior.ColorIndex = 4
                ElseIf ws.Range("j" & sumtable_row_qtr).Value < 0 Then
                    ws.Range("j" & sumtable_row_qtr).Interior.ColorIndex = 3
                End If
        
    'Calc percent change = (final price-opening price)/opening price and print to summary table
        ws.Range("k" & sumtable_row_qtr).Value = (close_price - open_price) / open_price
        'format cell as percent
            ws.Range("k" & sumtable_row_qtr).NumberFormat = "0.00%"
        'conditional formatting
            If ws.Range("k" & sumtable_row_qtr).Value > 0 Then
                ws.Range("k" & sumtable_row_qtr).Interior.ColorIndex = 4
                ElseIf ws.Range("k" & sumtable_row_qtr).Value < 0 Then
                ws.Range("k" & sumtable_row_qtr).Interior.ColorIndex = 3
                End If
    
    sumtable_row_qtr = sumtable_row_qtr + 1
        End If

    Next i

'Total stock volume = summation of all volumes for unique ticker symbol
    For i = 2 To lastrow
      If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        vol = vol + ws.Cells(i, 7).Value
        ws.Range("l" & sumtable_row_vol) = vol
        sumtable_row_vol = sumtable_row_vol + 1
        vol = 0
            Else
                vol = vol + ws.Cells(i, 7).Value
        End If
    Next i

'Greatest Values table
Dim max_pct, min_pct, max_vol As LongLong
For k = 2 To lastrow_sum

  'Pull greatest values data, max % change
  max_pct = WorksheetFunction.Max(ws.Range("K" & 2 & ":" & "K" & lastrow_sum))
  If ws.Cells(k, 11).Value = max_pct Then
    ws.Cells(2, 16).Value = max_pct
    ws.Range("p" & 2).NumberFormat = "0.00%"
    ws.Cells(2, 15).Value = ws.Cells(k, 9).Value
  Exit For
  End If
  Next k
  
  'Greatest values, min % change
For k = 2 To lastrow_sum
  min_pct = WorksheetFunction.Min(ws.Range("K" & 2 & ":" & "K" & lastrow_sum))
  If ws.Cells(k, 11).Value = min_pct Then
  ws.Cells(3, 16).Value = min_pct
  ws.Range("p" & 3).NumberFormat = "0.00%"
  ws.Cells(3, 15).Value = ws.Cells(k, 9).Value
  Exit For
 End If
  Next k
  
'Pull greatest values data, max volume
For k = 2 To lastrow_sum
    max_vol = WorksheetFunction.Max(ws.Range("L" & 2 & ":" & "L" & lastrow_sum))
    If ws.Cells(k, 12).Value = max_vol Then
  ws.Cells(4, 16).Value = max_vol
  ws.Cells(4, 15).Value = ws.Cells(k, 9).Value
  
  Exit For
End If
  Next k
   
ws.Columns("i").AutoFit
ws.Columns("j").AutoFit
ws.Columns("k").AutoFit
ws.Columns("l").AutoFit
ws.Columns("n").AutoFit
ws.Columns("P").AutoFit
   
sumtable_row_ticker = 2
sumtable_row_qtr = 2
sumtable_row_vol = 2


Next ws
End Sub





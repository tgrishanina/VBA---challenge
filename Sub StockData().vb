Sub StockData()
    
    'Define all variables
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim LastRow As Long
    Dim TickerSymbol As String
    Dim Summary_Row As Integer
    Dim Opening_Price As Double
    Dim Closing_Price As Double
    Dim Price_Change As Double
    Dim Percent_Change As Double
    
    Dim Greatest_Increase As Double
    Dim Greatest_Decrease As Double
    Dim Greatest_Volume As Double
    Dim Ticker_Greatest_Increase As String
    Dim Ticker_Greatest_Decrease As String
    Dim Ticker_Greatest_Volume As String
    
    Dim Worksheet_Count As Integer
    Dim j As Integer
    Worksheet_Count = ActiveWorkbook.Worksheets.Count
    
    
    Dim Total_Vol As Double
    Total_Vol = 0

    
    'reference the current workbook
    Set wb = ThisWorkbook
    
    'loop through all worksheets
    For j = 1 To Worksheet_Count
    
        For Each ws In wb.Worksheets
    
            LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
            Summary_Row = 2

        For i = 2 To LastRow
            ' Check if the current row is the first occurrence of the ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                TickerSymbol = ws.Cells(i, 1).Value
                ' Opening price at the beginning of the quarter
                Opening_Price = ws.Cells(i, 3).Value
                
            End If

            ' Check if the current row is the last occurrence of the ticker symbol
            If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                ' Closing price at the end of the quarter
                Closing_Price = ws.Cells(i, 6).Value

                ' Calculate price change
                Price_Change = Closing_Price - Opening_Price
                
                'Calculate percent change
                Percent_Change = Price_Change / Opening_Price
                
                'Add to the total volume
                Total_Vol = Total_Vol + Cells(i, 7).Value
                
                'Return the total volume
                ws.Cells(Summary_Row, 12).Value = Total_Vol
                
                'reset total volume
                Total_Vol = 0

                ' Output results to summary row
                ws.Cells(Summary_Row, 9).Value = TickerSymbol
                ws.Cells(Summary_Row, 10).Value = Price_Change
                ws.Cells(Summary_Row, 11).Value = Percent_Change
    
                'add color grading for pos and neg values
                If Price_Change > 0 Then
                  ws.Cells(Summary_Row, 10).Interior.ColorIndex = 4
                ElseIf Price_Change < 0 Then
                    ws.Cells(Summary_Row, 10).Interior.ColorIndex = 3
                End If
                
                 If Percent_Change > 0 Then
                  ws.Cells(Summary_Row, 11).Interior.ColorIndex = 4
                ElseIf Percent_Change < 0 Then
                    ws.Cells(Summary_Row, 11).Interior.ColorIndex = 3
                End If
                

                ' Move to the next summary row
                Summary_Row = Summary_Row + 1
                
            Else
                'add to the total volume before moving on
                Total_Vol = Total_Vol + Cells(i, 7).Value
                
              
            End If
        Next i
    
        'find the greatest increase, decrease, and volume
        Greatest_Increase = Application.WorksheetFunction.Max(ws.Columns("K"))
        Greatest_Decrease = Application.WorksheetFunction.Min(ws.Columns("K"))
        Greatest_Volume = Application.WorksheetFunction.Max(ws.Columns("L"))

        'find the rows for the corresponding ticker symbols
        Row_Greatest_Increase = Application.Match(Greatest_Increase, ws.Columns("K"), 0)
        Row_Greatest_Decrease = Application.Match(Greatest_Decrease, ws.Columns("K"), 0)
        Row_Greatest_Volume = Application.Match(Greatest_Volume, ws.Columns("L"), 0)

        'Determine the ticker symbols for the greatest increase, decrease, and volume
        Ticker_Greatest_Increase = ws.Cells(Row_Greatest_Increase, 9).Value
        Ticker_Greatest_Decrease = ws.Cells(Row_Greatest_Decrease, 9).Value
        Ticker_Greatest_Volume = ws.Cells(Row_Greatest_Volume, 9).Value

        ' Output the ticker symbols to column P
        ws.Range("P2").Value = Ticker_Greatest_Increase
        ws.Range("P3").Value = Ticker_Greatest_Decrease
        ws.Range("P4").Value = Ticker_Greatest_Volume
        
        'Output the values themselves to column Q
        ws.Range("Q2").Value = Greatest_Increase
        ws.Range("Q3").Value = Greatest_Decrease
        ws.Range("Q4").Value = Greatest_Volume
    
        'Add row labels
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Volume"
        
        'Add headers for all tables
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Quarterly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
    
    Next ws
  
  Next j
  
End Sub



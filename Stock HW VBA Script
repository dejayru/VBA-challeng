Sub Stock_Analysis_Multi_Year():

'For each worksheet create a loop to get the following:
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Sheets
        ws.Activate
        
    'Create summary table
            ws.Cells(1, 9).Value = "Ticker"
            ws.Cells(1, 10).Value = "Yearly Change"
            ws.Cells(1, 11).Value = "Percent Change"
            ws.Cells(1, 12).Value = "Total Stock Volume"
            
        'Set Variables
        
            Dim ticker_symbol As String
            Dim total_vol As Double
            total_vol = 0
            Dim rowcount As Long
            rowcount = 2
            Dim year_open As Double
            year_open = 0
            Dim year_close As Double
            year_close = 0
            Dim year_change As Double
            year_change = 0
            Dim percent_change As Double
            percent_change = 0
            
        'Get Ticker Symbol
        
            Dim lastrow As Long
            lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
            For i = 2 To lastrow
            
                If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                    year_open = ws.Cells(i, 3).Value
                    
                End If
        'Calc total stock vloume
        
            total_vol = total_vol + ws.Cells(i, 7)
                If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
                
                    ws.Cells(rowcount, 9).Value = ws.Cells(i, 1).Value
                    ws.Cells(rowcount, 12).Value = total_vol
                    
                    year_close = ws.Cells(i, 6).Value
                    ws.Range("L:L").NumberFormat = "#,##0"
                            
        'Calc the change in stock price (Closing at EOY - Open at BOY)
                    year_change = year_close - year_open
                    ws.Cells(rowcount, 10).Value = year_change
           
        'Conditional format the change (Positive = Green & Negative = Red)
                    If year_change >= 0 Then
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 4
                    Else
                        ws.Cells(rowcount, 10).Interior.ColorIndex = 3
                    End If
                
        'Calc the percent change in stock price((Closing at EOY - Open at BOY)/Open at Boy)*100
                    If year_open = 0 And year_close = 0 Then
                        percent_change = 0
                        ws.Cells(rowcount, 11).Value = percent_change
                        ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    ElseIf year_open = 0 Then
                
        'Calc for new stock (BOY open is 0, need this as the change would be large and skew the data)
                       Dim percent_change_NA As String
                       percent_change_NA = "New Ticker"
                       ws.Cells(rowcount, 11).Value = percent_change
                    Else
                        percent_change = year_change / year_open
                        ws.Cells(rowcount, 11).Value = percent_change
                        ws.Cells(rowcount, 11).NumberFormat = "0.00%"
                    End If
         'Format Columns to Expand
                    ws.Columns("J:L").EntireColumn.AutoFit
                                
        'Move to next Row
                rowcount = rowcount + 1
            
        'reset vol, BOY price, EOY close, year, and year percent
                total_vol = 0
                year_open = 0
                year_close = 0
                year_change = 0
                percent_change = 0
            
                End If
              
             Next i
             
    'Bonus - Creat tabel to show Greatest % increase, Greatest % decrease, & Greatest Total Vol
                                
            ws.Cells(2, 15).Value = "Greatest % Increase"
            ws.Cells(3, 15).Value = "Greatest % Decrease"
            ws.Cells(4, 15).Value = "Greatest Total Volume"
            ws.Cells(1, 16).Value = "Ticker"
            ws.Cells(1, 17).Value = "Value"
            
            lastrow = ws.Cells(Rows.Count, 9).End(xlUp).Row
            
        
            For i = 2 To lastrow
      'Look for best performing stock
                If ws.Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & lastrow)) Then
                    ws.Cells(2, 16) = ws.Cells(i, 9).Value
                    ws.Cells(2, 17) = ws.Cells(i, 11).Value
                    ws.Cells(2, 17).NumberFormat = "0.00%"
                            
      'Look for worst performing stock
                
                
                ElseIf ws.Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & lastrow)) Then
                        ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
                        ws.Cells(3, 17).NumberFormat = "0.00%"
                            
      'Look for highest Vol stock
                ElseIf ws.Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & lastrow)) Then
                        ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
                        ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
                        ws.Cells(4, 17).NumberFormat = "#,##0"
                     
                End If
                
            
             Next i
            
       'Adjut Column width
            ws.Columns("O:Q").EntireColumn.AutoFit
        
             
    Next ws
    
    
End Sub

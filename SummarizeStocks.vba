Sub SummarizeStocks()
For Each ws In Worksheets
        
        Dim RunningTotal As Double
        Dim StockCount As Double
        Dim LastRow As Double
        Dim StockStart As Double
        Dim StockEnd As Double
        Dim GreatestIncrease As Double
        Dim IncreaseTicker As String
        Dim GreatestDecrease As Double
        Dim DecreaseTicker As String
        Dim GreatestStock As Double
        Dim StockTicker As String
        
        
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        RunningTotal = 0
        StockCount = 1
        StockStart = ws.Cells(2, 3).Value
        
        GreatestIncrease = 0
        GreatestDecrease = 0
        GreatestStock = 0
        
        'Add headings to columns
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percentage Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        'Check all rows
        For i = 2 To LastRow
                
            'Look for where the stocks change between rows
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                    
                'Increase running total by new vol
                RunningTotal = RunningTotal + ws.Cells(i, 7)
                StockCount = StockCount + 1
                StockEnd = ws.Cells(i, 6)
                    
                'Record values for each stockincluding prercentage change
                ws.Cells(StockCount, 9).Value = ws.Cells(i, 1).Value
                ws.Cells(StockCount, 10).Value = StockEnd - StockStart
                    
                'Check for zeros
                If StockStart = 0 Then
                    ws.Cells(StockCount, 11).Value = 0
                Else
                    ws.Cells(StockCount, 11).Value = (StockEnd - StockStart) / StockStart
                    ws.Cells(StockCount, 11).NumberFormat = "0.00%"
                    ws.Cells(StockCount, 12).Value = RunningTotal
                End If
                    
                'Format positive and negative percentages, red and green
                If ws.Cells(StockCount, 11).Value < 0 Then
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 3
                ElseIf ws.Cells(StockCount, 11).Value > 0 Then
                    ws.Cells(StockCount, 10).Interior.ColorIndex = 4
                End If
                    
                'Reset the value of the stock at the start
                StockStart = ws.Cells(i + 1, 3).Value
        
                'Reset running total
                RunningTotal = 0
                    
            Else
                'Add amount for running total when stock ticker doesn't change
                RunningTotal = RunningTotal + ws.Cells(i, 7).Value
                    
                    
            End If
                
        Next i
            
        'Search through summary
        For j = 2 To StockCount
            
            'Find greatest increase and greatest decrease
            If ws.Cells(j, 11).Value > GreatestIncrease Then
                GreatestIncrease = ws.Cells(j, 11).Value
                IncreaseTicker = ws.Cells(j, 9).Value
            ElseIf ws.Cells(j, 11).Value < GreatestDecrease Then
                GreatestDecrease = ws.Cells(j, 11).Value
                DecreaseTicker = ws.Cells(j, 9).Value
            End If
                    
            'Find largest stock value
            If ws.Cells(j, 12).Value > GreatestStock Then
                GreatestStock = ws.Cells(j, 12).Value
                StockTicker = ws.Cells(j, 9)
            End If
                    
        Next j
            
        'Write headers for summary
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
            
        'Record summary
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(2, 15).Value = IncreaseTicker
        ws.Cells(2, 16).Value = GreatestIncrease
        ws.Cells(2, 16).NumberFormat = "0.00%"
            
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(3, 15).Value = DecreaseTicker
        ws.Cells(3, 16).Value = GreatestDecrease
        ws.Cells(3, 16).NumberFormat = "0.00%"
                    
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(4, 15).Value = StockTicker
        ws.Cells(4, 16).Value = GreatestStock
        
        ws.Columns("I:P").AutoFit
            
      Next ws
      
End Sub

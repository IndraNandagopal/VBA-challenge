Attribute VB_Name = "Module1"
'Create a VBA script to analyze generated stock market data

Sub stockData():
    
'looping through the worksheets
    
  For Each ws In Worksheets
  
    'Variable to hold the ticker symbol
    tickerName = ""
    
    'Variable to hold the total Volume
    totalStockVolume = 0

    'Variable to hold the summary table starter row
    summaryTableRow = 2
       
    Dim openStockPrice As Double
    
    'Set initial open Stock price.
    openStockPrice = ws.Cells(2, 3).Value
    
    Dim CloseStockPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    
    
    'To Label summary table headers
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    'Used Range for setting the labels for second summary table
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
     
    'use function to find the last row in the sheet
    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'loop from 2 in Column A to last row
    
    For Row = 2 To lastRow
        'if the ticker changes do...
                            
            If ws.Cells(Row + 1, 1).Value <> ws.Cells(Row, 1).Value Then
               'first set the ticker Name
                tickerName = ws.Cells(Row, 1).Value
                
                'add the last stock Volume from the row
                totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
                                
                'set the close stock price
                CloseStockPrice = ws.Cells(Row, 6).Value
                                
                'add the ticker symbol to I column in the summary tabe row
                ws.Cells(summaryTableRow, 9).Value = tickerName
                
                'add the total charges to the L column in the summary table row
                ws.Cells(summaryTableRow, 12).Value = totalStockVolume
                
                yearlyChange = CloseStockPrice - openStockPrice
                
                If openStockPrice = 0 Then
                    percentChange = 0
                Else
                    percentChange = yearlyChange / openStockPrice
                End If
                
                'add yearly change and percent change to summary Table
                
                ws.Cells(summaryTableRow, 10).Value = yearlyChange
                ws.Cells(summaryTableRow, 11).Value = percentChange
                ws.Cells(summaryTableRow, 10).NumberFormat = "0.00"
                ws.Cells(summaryTableRow, 11).NumberFormat = "0.00%"
                
                
                'Reset open stock price
                openStockPrice = ws.Cells(Row + 1, 3).Value
                
                
                'add the currency format to H column
                'ws.Cells(summaryTableRow, 8).Style = "Currency"
                'ws.Cells(summaryTableRow, 8).NumberFormat = "0.00%"
                'ws.Cells(summaryTableRow, 8).NumberFormat = "$#,##"
                
                'go to next summary table row
                summaryTableRow = summaryTableRow + 1
                
                'reset the total stock volume to 0
                totalStockVolume = 0
                
            Else
            
            'if the ticker name stays the same..add on to the total stock volume from column G
            totalStockVolume = totalStockVolume + ws.Cells(Row, 7).Value
            
        End If
    Next Row
    
    'Conditional formatting that will highlight positive change in green and negative change in red
    'use function to find the last row in the sheet
    
    lastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    For i = 2 To lastRow
    
        If ws.Cells(i, 10).Value > 0 Then
            ws.Cells(i, 10).Interior.ColorIndex = 4
        Else
            ws.Cells(i, 10).Interior.ColorIndex = 3
        End If
        
    Next i
    
    'Add functionality to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume"
    
    'find the ticker name with the maxPercent
    
    ws.Cells(2, 17).Value = 0
    ws.Cells(3, 17).Value = 0
    ws.Cells(4, 17).Value = 0
    
    For i = 2 To lastRow
   
        If ws.Cells(i, 11).Value > ws.Cells(2, 17).Value Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(2, 17).Value = ws.Cells(i, 11).Value
        End If
    Next i
    
        
    'find the ticker name with the minPercent
    For i = 2 To lastRow
   
        If ws.Cells(i, 11).Value < ws.Cells(3, 17).Value Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(3, 17).Value = ws.Cells(i, 11).Value
        End If
    Next i
    
    'find the greatest total volume with ticker
    For i = 2 To lastRow
   
        If ws.Cells(i, 12).Value > ws.Cells(4, 17).Value Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            ws.Cells(4, 17).Value = ws.Cells(i, 12).Value
        End If
    Next i
    
    'Final formatting..
    
    ws.Cells(2, 17).NumberFormat = "0.00%"
    ws.Cells(3, 17).NumberFormat = "0.00%"
    ws.Cells(4, 17).NumberFormat = "0.00E+00"
    ws.Columns("I:Q").AutoFit
    
    'Modified teact to Bold for columns from A to Q
    ws.Columns("A:Q").Font.Bold = True
   
   'Process Next worksheet
   Next ws
   
End Sub

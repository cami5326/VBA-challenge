Attribute VB_Name = "Module1"
Sub tickerData():

   'to loop through all the worksheets
   For Each ws In Worksheets
   
        Dim WorksheetName As String
        WorksheetName = ws.Name 'stores the names of the worksheets
          
    'Name the columns in all the worksheets
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    ws.Range("O2").Value = "Greatest % increase"
    ws.Range("O3").Value = "Greatest % decrease"
    ws.Range("O4").Value = "Greatest total volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    
    
        ' check on the ticker names
        Dim tickerName As String
    
        ' variable to hold the totals for the tickers
        Dim tickerVolume As Double
             
        ' variable to hold the rows in the total colums (columns A and G)
        Dim tickerRows As Double
        
        ' declare variable to hold the row
        Dim row As Double
     
        ' declare variable to hold the first open ticker value
        Dim openTicker As Double
        
        ' declare variable to hold the last close ticker value
        Dim closeTicker As Double
        
        ' declare variable to hold the yearly change value = last close - first open
        Dim YearlyChange As Double
        
        ' declare variable to hold the percentage change value = (last closed - first open) / first open
        Dim percentageChange As Double
        
        ' declare variable to hold the ticker name of the Max % increase value (Greatest % increase)
        Dim maxTicker As String
        
        ' declare variable to hold the the Max % increase value
        Dim maxPercentage As Double
        
        ' declare variable to hold the ticker name of the Min % increase value (Greatest % decrease)
        Dim minTicker As String
        
         ' declare variable to hold the Min % increase value (Greatest % decrease)
        Dim minPercentage As Double
        
        ' declare variable to hold the ticker name of the Greatest Total Volume value
        Dim volumeTicker As String
        
        ' declare variable to hold the the Greatest Total Volume value
        Dim maxVolume As Double
        
        tickerVolume = 0   ' start the initial total at 0
        tickerRows = 2 ' first row to populate in columns will be row 2
        maxPercentage = 0  ' start the initial total at 0
        
        ' find the last row in all the worksheets
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).row
    
    ' loop through the rows and check the changes in the tickers
    For row = 2 To lastRow
    
       ' check the changes in the tickers, last row of the ticker
        If ws.Cells(row + 1, 1).Value <> ws.Cells(row, 1).Value Then
        
            ' set the ticker name
             tickerName = ws.Cells(row, 1).Value ' grabs the value from column A BEFORE the change
             
            ' add to the ticker total
             tickerVolume = tickerVolume + ws.Cells(row, 7).Value ' grabs the value from G BEFORE the change
             
             ' display the ticker name on the current row of the total columns
             ws.Cells(tickerRows, 9).Value = tickerName
                     
             ' display the ticker total on the current row of the total columns
             ws.Cells(tickerRows, 12).Value = tickerVolume
             
             ' set the last close ticker value
             closeTicker = ws.Cells(row, 6).Value
             
             ' set the Yearly Change value
             YearlyChange = closeTicker - openTicker
              
              ' set the Yearly Change value
             percentageChange = (closeTicker - openTicker) / openTicker
              
         
                ' calculate the Greatest % increase value
                If percentageChange > maxPercentage Then
                        
                maxTicker = tickerName
                
                maxPercentage = percentageChange
            
                        
                End If
                    
                    
                ' calculate the Greatest % decrease value
                If percentageChange < minPercentage Then
                    
                minTicker = tickerName
                    
                minPercentage = percentageChange
                   
                End If
                    
                
                ' calculate the Greatest total volume value
                If tickerVolume > maxVolume Then
                    
                volumeTicker = tickerName
                    
                maxVolume = tickerVolume
                    
                End If
            
                            
         ' prints the results for YearlyChange and percentageChange values and sets formatting
         ws.Cells(tickerRows, 10).Value = YearlyChange
         
         ws.Cells(tickerRows, 10).NumberFormat = "0.00"
         
         ws.Cells(tickerRows, 11).Value = percentageChange
          
         ws.Cells(tickerRows, 11).NumberFormat = "0.00%"
          
        
                'sets color for Yearly Change Column J
                If YearlyChange > 0 Then
                       
                   ws.Cells(tickerRows, 10).Interior.ColorIndex = 4
                       
                Else
                    
                   ws.Cells(tickerRows, 10).Interior.ColorIndex = 3
                    
                End If
          
                'sets color for Yearly Change Column K
                If percentageChange > 0 Then
                       
                   ws.Cells(tickerRows, 11).Interior.ColorIndex = 4
                       
                Else
                    
                   ws.Cells(tickerRows, 11).Interior.ColorIndex = 3
                    
                End If
            
             ' add 1 to the ticker row to go to the next row
             tickerRows = tickerRows + 1
            
             ' reset the ticker total for the next ticker
             tickerVolume = 0
        
      
        ' check the changes in the tickers, first row of the ticker
        ElseIf ws.Cells(row - 1, 1).Value <> ws.Cells(row, 1).Value Then
         
            openTicker = ws.Cells(row, 3).Value
            
            tickerVolume = tickerVolume + ws.Cells(row, 7).Value
            
        
        ' if there is no change in the tickers, keep adding to the total (neither the first or last row)
        Else
        
            tickerVolume = tickerVolume + ws.Cells(row, 7).Value
             
        End If
    
    Next row
        
        
        ' prints the results of variables values and sets formatting
          ws.Cells(2, 16).Value = maxTicker
          ws.Cells(2, 17).Value = maxPercentage
          ws.Cells(3, 16).Value = minTicker
          ws.Cells(3, 17).Value = minPercentage
          ws.Cells(4, 16).Value = volumeTicker
          ws.Cells(4, 17).Value = maxVolume
          ws.Cells(2, 17).NumberFormat = "0.00%"
          ws.Cells(3, 17).NumberFormat = "0.00%"
          ws.Cells.EntireColumn.AutoFit
                   
    
   
   Next ws

End Sub


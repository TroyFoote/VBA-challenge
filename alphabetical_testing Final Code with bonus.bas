Sub stock_results()
    
    'create loop to add headers to each sheet
    For Each ws In Worksheets
    
    'create variables
    Dim WorksheetName As String
    WorksheetName = ws.Name
    Dim Total_Volume As Double
    TotalVolume = 0
    Dim Ticker_Name As String
    Ticker_Name = " "
    Dim Start_Price As Double
    Start_Price = 0
    Dim Percent_Change As Double
    
    'Create Column Headers for Required data and add to each worksheet
    ws.Cells(1, 10).Value = "Ticker"
    ws.Cells(1, 11).Value = "Yearly Change"
    ws.Cells(1, 12).Value = "Percent Change"
    ws.Cells(1, 13).Value = "Total Stock Volume"
    
    'Create column headers for bonus features
    ws.Cells(1, 17).Value = "Ticker"
    ws.Cells(1, 18).Value = "Value"
    
    'Create Row Headers for Bonus Features
    ws.Cells(2, 16).Value = "Greatest % Increase"
    ws.Cells(3, 16).Value = "Greatest % Decrease"
    ws.Cells(4, 16).Value = "Greatest Total Volume"
    

    'set a location for each ticker name in a summary table
    Dim Summary_Table_Row As Long
    Summary_Table_Row = 2
          
    'set last row for ticker names
    Last_Row = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
      
    'set initial start price value of each sheet
    Start_Price = ws.Cells(2, 3).Value
    
    'loop through all ticker names
    For i = 2 To Last_Row
    
    'set close price
    Close_Price = ws.Cells(i, 6).Value
    
    'set open price of each new ticker
    Dim New_Year_Price As Boolean
        
        If New_Year_Price = False Then
        'Set Opening_Price
        Dim Opening_Price As Double
        Opening_Price = ws.Cells(i, 3).Value
        
        New_Year_Price = True
        
    End If
      
        'check if ticker name is repeating, if not
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        'Set ticker name
        Ticker_Name = ws.Cells(i, 1).Value
        
        'set Yearly Change value
        Yearly_Change = Close_Price - Opening_Price
                
        'set percent change value
        Percent_Change = (Yearly_Change / Opening_Price)
                
        'set total volume per ticker
        Total_Volume = Total_Volume + Cells(i, 7).Value
        
        'add ticker name to summary table
        ws.Range("J" & Summary_Table_Row).Value = Ticker_Name
        
        'add yearly change to Summary table
        ws.Range("K" & Summary_Table_Row).Value = Yearly_Change
        
        'add percent change to Summary table
        ws.Range("L" & Summary_Table_Row).Value = Percent_Change
              
        'add stock volume per ticker
        ws.Range("M" & Summary_Table_Row).Value = Total_Volume
        
        'move to next empty row in summary table
        Summary_Table_Row = Summary_Table_Row + 1
        
        'reset total volume
        Total_Volume = 0
        
        'switch back to next price
        New_Year_Price = False
        
        Else
        
        Total_Volume = Total_Volume + ws.Cells(i, 7).Value
          
    End If
         
    Next i
    
        'convert yearly change column to show two decimal places and $
        ws.Range("K" & Summary_Table_Row).NumberFormat = "$0.00"
        
        'convert percent change column to show two decimal places and %
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
        
        'set figures for Max Percentage, Min Percentage and Max Volume
        Increase = WorksheetFunction.Max(ws.Range("L:L"))
        Decrease = WorksheetFunction.Min(ws.Range("L:L"))
        GTVolume = WorksheetFunction.Max(ws.Range("M:M"))
        
        ws.Range("R2") = Increase
        ws.Range("R3") = Decrease
        ws.Range("R4") = GTVolume
        
        'Set formatting to show two decimal places and %
        ws.Range("R2:R3").NumberFormat = "0.00%"
        
        'find matching ticker symbols for Max Percentage, Min Percentage and Max Volume
        If ws.Cells(i, 12).Value = Increase Then
        ws.Range("Q2").Value = ws.Cells(i, 10).Value
        
    End If
    
        If ws.Cells(i, 12).Value = Decrease Then
        ws.Range("Q3").Value = ws.Cells(i, 10).Value
        
    End If
        
        If ws.Cells(i, 13).Value = GTVolume Then
        ws.Range("Q4").Value = ws.Cells(i, 10).Value
        
    End If
    
    
    
    'set last row for summary table
    Final_Row = ws.Cells(ws.Rows.Count, 11).End(xlUp).Row
    
      For j = 2 To Final_Row
      
        'set conditional formatting to yearly change column
        If ws.Cells(j, 11).Value > 0 Then
        ws.Cells(j, 11).Interior.ColorIndex = 4
                
        Else
        
        ws.Cells(j, 11).Interior.ColorIndex = 3
           
    End If
    
        'set conditional formatting to percent change column
        If ws.Cells(j, 12).Value > 0 Then
        ws.Cells(j, 12).Interior.ColorIndex = 4
                
        Else
        
        ws.Cells(j, 12).Interior.ColorIndex = 3
           
    End If
    
    Next j
    
    'find matching ticker symbols for Max Percentage, Min Percentage and Max Volume
        If ws.Cells(i, 12).Value = Range("R2").Value Then
        ws.Range("Q2").Value = ws.Cells(i, 10).Value
        
    End If
    
        If ws.Cells(i, 12).Value = Range("R3").Value Then
        ws.Range("Q3").Value = ws.Cells(i, 10).Value
        
    End If
        
        If ws.Cells(i, 13).Value = Range("R4").Value Then
        ws.Range("Q4").Value = ws.Cells(i, 10).Value
        
    End If
    
    
    'move to next worksheet
    Next ws
  
End Sub


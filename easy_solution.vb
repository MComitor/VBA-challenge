Sub ticker_symbol_easy()

    ' Set initial variables to hold all the things
    Dim ticker_symbol As String
    Dim yearly_change As Double
    Dim percent_change As Double
    Dim stock_vol As LongLong
    stock_vol = 0
    percent_change = 0
    yearly_change = 0
      
    ' Create locations for the output data of all the things
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    ' Keep track of the last row of the worksheet and the summary table
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    lastSummaryRow = Cells(Rows.Count, 10).End(xlUp).Row

  
    ' Keep track of the location for each ticker symbol  in the summary table
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
  
    ' Loop through all stock market ticker symbols
    For i = 2 To lastRow
      
      ' Check if we are still within the same ticker symbol, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set the ticker symbol name
        ticker_symbol = Cells(i, 1).Value
        
        ' Add  the total stock volume
        stock_vol = stock_vol + Cells(i, 7).Value
              
        ' Print the ticker symbol in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker_symbol
        
        ' Print the yearly change to summary table
        Range("J" & Summary_Table_Row).Value = yearly_change
        
        ' Print the % change in the Summary Table
        Range("K" & Summary_Table_Row).Value = "%" & percent_change
        
        ' Print the stock volume amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = stock_vol
                      
        'If yearly change value is 0 or greater, turn the cell green
        If Cells(Summary_Table_Row, 10).Value >= 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 4
            
        'Otherwise turn the cell red
        ElseIf Cells(Summary_Table_Row, 10).Value < 0 Then
            Cells(Summary_Table_Row, 10).Interior.ColorIndex = 3
            
        End If
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset all running tallies to zero
        stock_vol = 0

      ' If the cell immediately following a row is the same ticker symbol...
      Else
        ' Add to running tallies
        stock_vol = stock_vol + Cells(i, 7).Value
        yearly_change = Cells(i, 6) - Cells(i, 3).Value
        percent_change = Round((yearly_change / Cells(i, 3) * 100), 2)
      End If

    Next i
    
    Dim maxIncreaseTicker As String
    Cells(1, 16) = maxIncreaseTicker
    
    Dim maxDecreaseTicker As String
    Cells(2, 16) = maxDecreaseTicker
    
    Dim maxIncrease As Double
    Cells(1, 14) = "Greatest % Increase"
    Cells(1, 16) = maxIncrease
    
    Dim maxDecrease As Double
    Cells(2, 14) = "Greatest % Decrease"
    Cells(2, 16) = maxDecrease
    
    Dim maxVolume As LongLong
    Cells(3, 14) = "Greatest Total Volume"
    Cells(3, 16).Value = maxVolume
  
    maxIncrease = Application.WorksheetFunction.Max(Range("K:K"))
    maxDecrease = Application.WorksheetFunction.Max(Range("K:K"))
    maxVolume = Application.WorksheetFunction.Max(Range("L:L"))

End Sub


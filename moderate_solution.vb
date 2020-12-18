' Create a script that will loop through all the stocks for one year and output the following information.
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.

' This script works on multiple sheets
Sub ticker_symbol_moderate()

    ' Add a sheet named "Combined Data"
    Sheets.Add.Name = "Combined_Data"
    'move created sheet to be first sheet
    Sheets("Combined_Data").Move Before:=Sheets(1)
    ' Specify the location of the combined sheet
    Set combined_sheet = Worksheets("Combined_Data")

    ' Loop through all sheets
    For Each ws In Worksheets

        ' Find the last row of the combined sheet after each paste
        ' Add 1 to get first empty row
        lastRow = combined_sheet.Cells(Rows.Count, "A").End(xlUp).Row + 1

        ' Find the last row of each worksheet
        ' Subtract one to return the number of rows without header
        lastRowTicker = ws.Cells(Rows.Count, "A").End(xlUp).Row - 1

        ' Copy the contents of each state sheet into the combined sheet
        combined_sheet.Range("A" & lastRow & ":G" & ((lastRowTicker - 1) + lastRow)).Value = ws.Range("A2:G" & (lastRowState + 1)).Value

  ' Set variable to hold all the things
  Dim ticker_symbol As String
  Dim yearly_change As Double
  Dim percent_change As Double
  Dim stock_vol As LongLong
  stock_vol = 0

  ' Get last rows of sheet and summary table row
  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  lastSummaryRow = Cells(Rows.Count, 9).End(xlUp).Row

     ' Create locations for the summary table column headers
      Range("I1") = "Ticker"
      Range("J1") = "Yearly Change"
      Range("K1") = "Percent Change"
      Range("L1") = "Total Stock Volume"
   
  ' Keep track of the location for each ticker symbol in the summary table
  Dim Summary_Table_Row As Long
  Summary_Table_Row = 2
  
    ' Loop through all stock market ticker symbols
    For i = 2 To lastRow
      
      ' Check if we are still within the same ticker symbol, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Add Ticker Symbol Name to summary table
        ticker_symbol = Cells(i, 1).Value
        Range("I" & Summary_Table_Row).Value = ticker_symbol

        ' Add Total Stock Volume to summary table
        stock_vol = stock_vol + Cells(i, 7).Value
        Range("L" & Summary_Table_Row).Value = stock_vol

        ' Add Yearly Change to Summary table
        Range("J" & Summary_Table_Row).Value = Round(yearly_change, 2)
        
        ' Print the % change in the Summary Table
        Range("K" & Summary_Table_Row).Value = "%" & percent_change
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset all running tallies to zero
        stock_vol = 0

            Else

              ' If the cell immediately following a row is the same ticker symbol...
              ' Add to the stock volume total
              ' open_rate = open_rate + Cells(i, 3).Value
              ' close_rate = close_rate + Cells(i, 6).Value
              stock_vol = stock_vol + Cells(i, 7).Value
              yearly_change = Cells(i, 6) - Cells(i, 3)
              percent_change = Round((yearly_change / Cells(i, 3) * 100), 2)
            

            If yearly_change < 0 Then
                      Cells(i, 10).Interior.ColorIndex = 3
                  
                  Else 
                      Cells(i, 10).Interior.ColorIndex = 4
                  
                  End If
            End If
          
        Next i

      Next ws

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
  
Next ws  
    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A:G").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:K").AutoFit
      
End Sub

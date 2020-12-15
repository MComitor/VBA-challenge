' Create a script that will loop through all the stocks for one year and output the following information.
' The ticker symbol.
' Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
' The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
' The total stock volume of the stock.

' This script works on individual sheets
Sub ticker_symbol()

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

    Next ws

  ' Set an initial variable for holding the ticker symbol name
  Dim ticker_symbol As String
  
  ' Set variables to hold the calculations of yearly change and percent change
  Dim yearly_change As Double
  Dim percent_change As Double
  
  ' Set an initial variable for holding the total per credit card brand
  Dim stock_vol As Long
  stock_vol = 0

  lastRow = Cells(Rows.Count, 1).End(xlUp).Row
  
  ' Set variable to 0 to hold the total of column c (opening rate)
  ' Dim open_rate As Double
  ' open_rate = 0

  ' Set variable to 0 to hold the total of column f (closing rate)
  ' Dim close_rate As Double
  ' close_rate = 0
  
  ' Keep track of the location for each ticker symbol  in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
    ' Loop through all stock market ticker symbols
    For i = 2 To lastRow
      
      ' Check if we are still within the same ticker symbol, if it is not...
      If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
        ' Set the ticker symbol name
        ticker_symbol = Cells(i, 1).Value
        
        ' Add  the total stock volume by unique symbol
        stock_vol = stock_vol + Cells(i, 7).Value
              
        ' Print the ticker symbol in the Summary Table
        Range("I" & Summary_Table_Row).Value = ticker_symbol
        
        ' Print the stock volume amount to the Summary Table
        Range("L" & Summary_Table_Row).Value = stock_vol
        
        ' Print the yearly change in the Summary table
        Range("J" & Summary_Table_Row).Value = yearly_change
        
        ' Print the % change in the Summary Table
        Range("K" & Summary_Table_Row).Value = percent_change
        
        ' Add one to the summary table row
        Summary_Table_Row = Summary_Table_Row + 1
        
        ' Reset all running tallies to zero
        stock_vol = 0
        open_rate = 0
        close_rate = 0

      ' If the cell immediately following a row is the same ticker symbol...
      Else
    
        ' Add to the stock volume total
        stock_vol = stock_vol + Cells(i, 7).Value
        open_rate = open_rate + Cells(i, 3).Value
        close_rate = close_rate + Cells(i, 6).Value
        yearly_change = close_rate - open_rate
        ' percent_change = some text
      
      End If
    
  Next i
  
    ' Copy the headers from sheet 1
    combined_sheet.Range("A1:G1").Value = Sheets(2).Range("A:G").Value
    
    ' Autofit to display data
    combined_sheet.Columns("A:K").AutoFit
      
End Sub

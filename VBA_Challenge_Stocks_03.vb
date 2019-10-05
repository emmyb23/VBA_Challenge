Sub Stocks_03()

  ' create new headers for summary data
  Range("I1").Value = "Ticker"
  Range("J1").Value = "Total Stock Volume"
  Range("K1").Value = "Yearly Change"
  Range("L1").Value = "Percent Change"
  

  ' Set a variable for holding the Ticker, Yearly change, and percent change
  Dim Ticker As String
  Dim Yearly_Change As Double
  Dim Percent_Change As Double
  Dim Open_stock As Double
  Dim Close_stock As Double
  Yearly_Change = 0
  Percent_Change = 0
  

  ' Set a variable for holding the total stock volume
  Dim total_stock_volume As Double
  total_stock_volume = 0
  
  'set worksheets and last rows
  Dim sht As Worksheet
  Dim LastRow As Long
  
  

  ' Keep track of the location for each ticker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  LastRow = Cells(Rows.Count, 1).End(xlUp).Row

  ' Loop through all tickers
  For i = 2 To LastRow
   
    ' Check if we are still within the same ticker
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      ' Set the ticker
      Ticker = Cells(i, 1).Value
      
      ' Add to the total_stock_volume, make sure you point to the right column!
      total_stock_volume = total_stock_volume + Cells(i, 7).Value

      'Add yearly change
      Yearly_Change = Yearly_Change + Cells(i, 3).Value - Cells(i, 6).Value
      
      'Add Percent change
      Percent_Change = Percent_Change + Cells(i, 3).Value / Cells(i, 6).Value
      
      'need a variable to hold the formatted percentage
      formatted_Percent_Change = Format(Percent_Change, "Percent")
      
      ' Print the ticker in the Summary Table
      Range("I" & Summary_Table_Row).Value = Ticker

      ' Print the total stock volume to the Summary Table
      Range("J" & Summary_Table_Row).Value = total_stock_volume
      
      'print yearly change
      Range("K" & Summary_Table_Row).Value = Yearly_Change
      
      'print percent change
      Range("L" & Summary_Table_Row).Value = formatted_Percent_Change

      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
    

    ' If the cell immediately following a row is the same ticker...
    Else

      ' Add to the ticker Total
      total_stock_volume = total_stock_volume + Cells(i, 7).Value



    End If
    
    

  Next i


End Sub


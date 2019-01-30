Sub Stock_Market()

 Dim ws_count As Integer
 Dim i As Integer
 
 ws_count = ActiveWorkbook.Worksheets.Count

 For i = 1 To ws_count
 
 Sheets(i).Select
 
  ' Set an initial variable for holding the stock ticker
  Dim Ticker As String
  
  ' Set an initial variable for holding the total volume traded
  Dim Volume_Total As Double
  Volume_Total = 0
  
  ' Keep track of the location for each ticker symbol in the table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2

  ' Create headers for the columns Ticker, Vol, Yearly Change, % Change
    Range("J1").Select
    ActiveCell.Value = "Ticker Symbol"
    Range("J1").Columns.AutoFit
    
    Range("K1").Select
    ActiveCell.Value = "Trading Volume"
    Range("K1").Columns.AutoFit
    
    Range("L1").Select
    ActiveCell.Value = "Yearly Change"
    Range("L1").Columns.AutoFit
    
    Range("M1").Select
    ActiveCell.Value = "% Change"
    Range("M1").Columns.AutoFit
  
  
  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  ' Loop through all ticker symbols
  For J = 2 To lastrow

    ' Check are we are still within the same ticker symbol, if we are not...
    If Cells(J + 1, 1).Value <> Cells(J, 1).Value Then

    ' Set ticker
     Ticker = Cells(J, 1).Value
    
    ' Add to the Volume Total
     Volume_Total = Volume_Total + Cells(J, 7).Value
    
    ' Print the Ticker symbol in the summary table
     Range("J" & Summary_Table_Row).Value = Ticker

    ' Print the Volume Total to the Summary Table
     Range("K" & Summary_Table_Row).Value = Volume_Total

    ' Add one to the summary table row
     Summary_Table_Row = Summary_Table_Row + 1
      
    ' Reset the Volume Total
     Volume_Total = 0

    Else

      ' Add to the Volume Total
      Volume_Total = Volume_Total + Cells(J, 7).Value

    End If

  Next J

Next i


End Sub
Sub stock_tracker()

Set Sheet = ThisWorkbook.Sheets
  ' Define variables
  Dim Stock_Name As String
  Dim year_end_price As Double
  Dim year_open_price As Double
  Dim year_per_chng As Double
  Dim max_vol_tot As Double
  Dim max_per_dec As Double
  Dim max_per_inc As Double
  Dim mvt_tick As String
  Dim mpd_tick As String
  Dim mpi_tick As String
  
  ' Set an initial variable for holding the total volume per stock
  Dim Stock_Tot_Volume As Double
  Stock_Tot_Volume = 0
  
  'Loop through all sheets
  For Each Sheet In Sheets
    Sheet.Activate

      ' Label the Summary Table
      Cells(1, 10).Value = "Ticker"
      Cells(1, 11).Value = "Ann_Chng"
      Cells(1, 12).Value = "Per_Chng"
      Cells(1, 13).Value = "Tot_Volume"
      
      ' Label the "Max" Table
      Cells(1, 17).Value = "Ticker"
      Cells(1, 18).Value = "Value"
      Cells(2, 16).Value = "Greatest % Increase"
      Cells(3, 16).Value = "Greatest % Decrease"
      Cells(4, 16).Value = "Greatest Total Volume"
    
      ' Keep track of the location for each stock in the summary table
      Dim Summary_Table_Row As Integer
      Summary_Table_Row = 2
    
      ' Loop through all stocks
      lastrow = Cells(Rows.Count, 1).End(xlUp).Row
      For i = 2 To lastrow
    
        ' Check if we are still within the same stock, if it is not...
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
          ' Set the stock name
          Stock_Name = Cells(i, 1).Value
          
          ' Set the year close value
          year_end_price = Cells(i, 6).Value
    
          ' Add to the Stock Volume Total
          Stock_Tot_Volume = Stock_Tot_Volume + Cells(i, 7).Value
    
          ' Print the stock name in the Summary Table
          Cells(Summary_Table_Row, 10).Value = Stock_Name
          
          ' Print the yearly change
          Cells(Summary_Table_Row, 11).Value = year_end_price - year_open_price
          
          ' Conditional formatting
          If Cells(Summary_Table_Row, 11).Value > 0 Then
            Cells(Summary_Table_Row, 11).Interior.ColorIndex = 4
            ElseIf Cells(Summary_Table_Row, 11).Value < 0 Then
            Cells(Summary_Table_Row, 11).Interior.ColorIndex = 3
            Else: Cells(Summary_Table_Row, 11).Interior.ColorIndex = 2
          End If
            
          
          ' Print the percent yearly change & make sure you don't divide by zero
          If year_open_price > 0 Then
            year_per_chng = (year_end_price - year_open_price) / year_open_price
            Cells(Summary_Table_Row, 12).Value = year_per_chng
            Else: Range("L" & Summary_Table_Row).Value = 0
          End If
    
          ' Print the total volume to the Summary Table
          Cells(Summary_Table_Row, 13).Value = Stock_Tot_Volume
    
          ' Add one to the summary table row
          Summary_Table_Row = Summary_Table_Row + 1
          
          ' Reset the Stock Total
          Stock_Tot_Volume = 0
          
        ' Checking if it's the start of a new stock, and thus a new year
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        
          ' Setting the year open price
          year_open_price = Cells(i, 3).Value
          
        ' If the cell immediately following a row is the same stock...
        Else
    
          ' Add to the Stock Volume Total
          Stock_Tot_Volume = Stock_Tot_Volume + Cells(i, 7).Value
    
        End If
    
      Next i
      
      ' Run a separate for loop on our summary table to obtain greatest values
        lastrow_tbl = Cells(Rows.Count, 10).End(xlUp).Row
        
        ' Set the maxes at 0
        max_per_inc = 0
        max_per_dec = 0
        max_vol_tot = 0
        
        
        For j = 2 To lastrow_tbl
        If Cells(j, 12).Value > max_per_inc Then
            max_per_inc = Cells(j, 12).Value
            mpi_tick = Cells(j, 10).Value
        ElseIf Cells(j, 12).Value < max_per_dec Then
            max_per_dec = Cells(j, 12).Value
            mpd_tick = Cells(j, 10).Value
        Else
        End If
    
        'Getting the Ticker symbol for Max Volume
        If Cells(j, 13).Value > max_vol_tot Then
            max_vol_tot = Cells(j, 13).Value
            mvt_tick = Cells(j, 10).Value
        Else
        End If
        Next j
        
        ' Fill the Max Tables
        Cells(2, 17).Value = mpi_tick
        Cells(2, 18).Value = max_per_inc
        Cells(3, 17).Value = mpd_tick
        Cells(3, 18).Value = max_per_dec
        Cells(4, 17).Value = mvt_tick
        Cells(4, 18).Value = max_vol_tot
        
        ' Reset the maxes
        max_per_inc = 0
        max_per_dec = 0
        max_vol_tot = 0
        
      ' Formatting
        Range("J:M").EntireColumn.AutoFit
        Range("P:R").EntireColumn.AutoFit
        Range("K2:K" & lastrow_tbl).NumberFormat = "$#,##0.0000000"
        Range("L2:L" & lastrow_tbl).NumberFormat = "#0.00%"
        Range("M2:M" & lastrow_tbl).NumberFormat = "###,###,###,##0"
        Range("R2:R3").NumberFormat = "0.00%"
        Range("R4").NumberFormat = "###,###,###,##0"

   Next Sheet
   
End Sub


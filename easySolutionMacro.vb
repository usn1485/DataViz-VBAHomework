Sub Multiple_year_stock_easy_solution()


'Loop through all sheets in a workbook

For Each ws In Worksheets

    Dim CurrentCell, NextCell, TickerName As String
    
    Dim TotalStockVolume As Double
    Dim SummaryRowTable As Integer
    
    SummaryRowTable = 2
    
    'Get last row number
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

    
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Total Stock Volume"
    
    For i = 2 To LastRow
    
    CurrentCell = ws.Cells(i, 1).Value
    NextCell = ws.Cells(i + 1, 1).Value
    
        If CurrentCell <> NextCell Then
        
          TickerName = ws.Cells(i, 1).Value
          TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
              
          ws.Range("I" & SummaryRowTable).Value = CurrentCell
          ws.Range("J" & SummaryRowTable).Value = TotalStockVolume
          SummaryRowTable = SummaryRowTable + 1
         
          ' Reset the TotalStockVolume Total
          TotalStockVolume = 0
        
        Else
        
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
            
        
        
        End If
    
   Next i

Next ws


End Sub




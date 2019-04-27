Sub Mulit_Year_Stock_Hard_solution()


'Loop through all sheets in a workbook

For Each ws In Worksheets
    
    Dim CurrentCell, NextCell, TickerName As String
    Dim TotalStockVolume As Double
    Dim SummaryRowTable As Integer
    Dim YearlyChange As Double
    Dim Opening_Stock_price As Double
    Dim closing_Stock_price As Double
    Dim PercentChange As Double
    Dim Opening_stock_flag As Boolean
    
    SummaryRowTable = 2
      
    'Get last row number
    lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    'Add the column headers
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    
    'Loop though all rows in a active sheet
    
    For i = 2 To lastrow
    
    CurrentCell = ws.Cells(i, 1).Value      'Get the value of current cell
    NextCell = ws.Cells(i + 1, 1).Value     'Get the value of Next Cell
   
    
        If CurrentCell <> NextCell Then     ' Do the stuff if current cell are not matching
        
            TickerName = CurrentCell
            closing_Stock_price = ws.Cells(i, 6).Value
    
    
    
            YearlyChange = (closing_Stock_price - Opening_Stock_price)  ' calculate yearly change of stock price
                
            If Opening_Stock_price <> 0 Then                                 'Check if Opening Stock price is not zero
                PercentChange = (YearlyChange / Opening_Stock_price)          'Calculate Percent Change of stocks.
            End If
                
                          
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
               
            ws.Range("I" & SummaryRowTable).Value = TickerName
            ws.Range("J" & SummaryRowTable).Value = YearlyChange
            ws.Range("K" & SummaryRowTable).Value = PercentChange
            ws.Range("K" & SummaryRowTable).NumberFormat = "0.00%"
            ws.Range("L" & SummaryRowTable).Value = TotalStockVolume
               
          
            If (YearlyChange < 0) Then                                      ' applying conditional formatting by color: red if Negative and green if positive.
                ws.Range("J" & SummaryRowTable).Interior.ColorIndex = 3
            Else
                ws.Range("J" & SummaryRowTable).Interior.ColorIndex = 4
            End If
          
           SummaryRowTable = SummaryRowTable + 1                            'Increament the summary row table to add value in next row
         
           'Reset the TotalStockVolume Total
            TotalStockVolume = 0
            Opening_stock_flag = False
          
        
        Else
        
            If Opening_stock_flag = False Then                           'Added flag to get opening Stock price of given Ticker.
            Opening_Stock_price = ws.Cells(i, 3).Value
            Opening_stock_flag = True
                                         
            End If
            
            
            TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
          
        
        End If
    
    Next i
   

 'For Hard Sol
    
    lastSummaryTableRow = ws.Cells(Rows.Count, 9).End(xlUp).Row    'get the last row no from summary table
    columnK = ws.Range("K2:K" & lastSummaryTableRow)
    columnL = ws.Range("L2:L" & lastSummaryTableRow)
     
      ws.Range("p1").Value = "Ticker"
      ws.Range("Q1").Value = "Value"
      ws.Range("O2").Value = "Greatest % Increase"
      ws.Range("O3").Value = "Greatest % Decrease"
      ws.Range("O4").Value = "Greatest Total Volume"
     
     
     MaxValue = WorksheetFunction.Max(columnK)                      'Get Max Value from the Column K
     maxValueRow = WorksheetFunction.Match(MaxValue, columnK, 0) + 1  'Use Match function to get the position of Max value
     TickerName = ws.Cells(maxValueRow, 9).Value                        'Get Ticker name
      ws.Range("p2").Value = TickerName
      ws.Range("Q2").Value = MaxValue
          
    
      MaxTotalValue = WorksheetFunction.Max(columnL)                'Get Max Value from the Column L
      maxTotalValueRow = WorksheetFunction.Match(MaxTotalValue, columnL, 0) + 1
      TickerName = ws.Cells(maxTotalValueRow, 9).Value
       ws.Range("p4").Value = TickerName
       ws.Range("Q4").Value = MaxTotalValue
    
      MinValue = WorksheetFunction.Min(columnK)                    'Get Min Value from the Column K
      MinValueRow = WorksheetFunction.Match(MinValue, columnK, 0) + 1
      TickerName = ws.Cells(MinValueRow, 9).Value
       ws.Range("p3").Value = TickerName
       ws.Range("Q3").Value = MinValue
      
       ws.Range("Q2:Q3").NumberFormat = "0.00%"
       ws.Range("I:R").Columns.AutoFit
 
Next ws               'Go on NEXT worksheet and repeat.

End Sub


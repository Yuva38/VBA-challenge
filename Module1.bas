Attribute VB_Name = "Module1"
Sub walstreet()


Dim ws As Worksheet
'For Each ws In Worksheets

Dim Ticker As String

Dim yearlyChange As Double
'yearlyChange = 0
Dim percentChange As String
'percentChange = 0
Dim TotalStockVolume As Double
TotalStockVolume = 0
Dim MaxValue As Double
Dim MinValue As Double

    'writes the column header
      
        
Dim answer_table_row As Integer
answer_table_row = 2
        
    
    'MsgBox lastRow
    'worksheetName = ws.Name
    'MsgBox worksheetName
    
     
    
For Each ws In Worksheets

        
        Range("M" & answer_table_row).Value = Cells(2, 3).Value
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "YearlyChange"
        Cells(1, 11).Value = "PercentChange"
        Cells(1, 12).Value = "TotalStockVolume"
        Cells(1, 13).Value = "openingValue"
        Cells(1, 14).Value = "ClosingValue"
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastRow
    
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
    
        'prints  Ticker only once when it finds those values on the cells are not equal
        
        Ticker = ws.Cells(i, 1).Value
        
        'add the stock volume if Tickers are not equal. (only execute the if statement) for total, we need to add argument in else section
        
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        'prints  yearly opening value for each ticker
        Range("M" & answer_table_row + 1).Value = ws.Cells(i + 1, 3).Value
        
        'prints yearend closing value for each ticker
        Range("N" & answer_table_row).Value = ws.Cells(i, 6).Value
        
        'calculates changes between opening and closing for each year of each ticker
        yearlyChange = Range("N" & answer_table_row).Value - Range("M" & answer_table_row).Value
        
        'calculate the percentage change
        'MsgBox
        percentChange = FormatPercent(Range("J" & answer_table_row).Value / Range("M" & answer_table_row).Value)
        
        
        
        'Print Ticker, yealrychange, percentage and totalstock volume on the respective column to the summary table
        Range("I" & answer_table_row).Value = Ticker
        Range("J" & answer_table_row).Value = yearlyChange
        Range("K" & answer_table_row).Value = percentChange
        Range("L" & answer_table_row).Value = TotalStockVolume
        If ws.Cells(answer_table_row, 10).Value > 0 Then
                ws.Cells(answer_table_row, 10).Interior.ColorIndex = 4
            Else
                ws.Cells(answer_table_row, 10).Interior.ColorIndex = 3
            End If
            
        'need to add row other wise it will keep rewriting on the same row
        
        answer_table_row = answer_table_row + 1
        
        TotalStockVolume = 0
        
         
    Else
             
        TotalStockVolume = TotalStockVolume + ws.Cells(i, 7).Value
        
    End If

            
            
               
    
 Next i
  
  Next ws
          

 
End Sub

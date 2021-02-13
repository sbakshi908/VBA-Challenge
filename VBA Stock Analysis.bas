Attribute VB_Name = "Module1"

Sub albhapetical_testing()

For Each ws In Worksheets
          ws.Activate
          
          Range("i1") = "Ticker"
          Range("j1") = "Yearly Change"
          Range("K1") = "Percent Change"
          Range("L1") = "Total Stock Volume"
        
    'set initial variable to hold ticker symbol
    
    Dim ticker As String
    
    'set initial variable to get total volume
    
    Dim stock_volume As Double
    stock_volume = 0
    Dim lastrow As Long
    
    'set variable for yearly change from open that year to close and variable to count the rows
    
    Dim year_change As Double
    year_change = 0
    Dim rowcount As Long
    rowcount = 0
    
    'set variable for % chnage of stockecs opening price at begining of year to close
    
    Dim percent_change As Double
    percent_change = 0
    'keep track of location for where each output (ticker/volume etc) will be printed
    
    Dim sum_table_row As Integer
    sum_table_row = 2
    
    'start looping through the different tickers
    lastrow = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastrow
    
        'set ticker name
        ticker = Cells(i, 1).Value
        
            If Cells(i + 1, 1).Value <> Cells(i, 1) Then
            
                stock_volume = stock_volume + Cells(i, 7)
                Cells(sum_table_row, 12).Value = stock_volume
                Cells(sum_table_row, 9).Value = ticker
                
                year_change = Cells(i, 6).Value - Cells(i - rowcount, 3).Value
                
               If Cells(i, 6).Value <> 0 And Cells(i - rowcount, 3).Value <> 0 Then
        
                    percent_change = year_change / Cells(i - rowcount, 3)
                    Cells(sum_table_row, 10).Value = year_change
                    Cells(sum_table_row, 11).Value = percent_change
                    Cells(sum_table_row, 11).NumberFormat = "0.00%"
                    
                End If
                    
                'condition al formatting to make negative year change red, postive gree
                        
                If year_change < 0 Then
                    Cells(sum_table_row, 10).Interior.Color = vbRed
                Else
                    Cells(sum_table_row, 10).Interior.Color = vbGreen
                End If
                    
                sum_table_row = sum_table_row + 1
                stock_volume = 0
                rowcount = 0
                
            Else
            rowcount = rowcount + 1
            stock_volume = stock_volume + Cells(i, 7)
            
            End If
        
    
        Next i
    Next ws

End Sub

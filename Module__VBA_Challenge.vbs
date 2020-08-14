Attribute VB_Name = "Module1"
Sub stockcounter()

    Dim ws As Worksheet
    
    For Each ws In Worksheets
    
        Dim ticker As String
        Dim ticker_row As Integer
        ticker_row = 2
        Dim openprice As Double
        Dim closeprice As Double
        Dim yearchange As Double
        Dim percentchange As String
        Dim volumetotal As Double
        
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Year Change"
        ws.Cells(1, 10).Value = "Percent Change"
        ws.Cells(1, 11).Value = "Total Stock Volume"
        
        openprice = ws.Cells(2, 3).Value
        
        NumRows = ws.Range("A2", ws.Range("A2").End(xlDown)).Rows.Count
    
        volumetotal = 0
    
        For i = 2 To NumRows
        
            'start adding volumes
            
            volumetotal = (volumetotal + ws.Cells(i, 7).Value)
    
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
            
            ' ticker symbol
            
                ticker = ws.Cells(i, 1).Value
                
                ws.Range("H" & ticker_row).Value = ticker
                
            ' year change
            
                closeprice = ws.Cells(i, 6).Value
                
                yearchange = closeprice - openprice
                
                ws.Range("I" & ticker_row).Value = yearchange
                
            ' conditional formatting
                
                If yearchange > 0 Then
                
                    ws.Range("I" & ticker_row).Interior.ColorIndex = 4
                    
                Else
                
                    ws.Range("I" & ticker_row).Interior.ColorIndex = 3
                
                End If
                
            ' percent change
            
                If openprice <> 0 Then
                
                    percentchange = FormatPercent(yearchange / openprice)
            
                    ws.Range("J" & ticker_row).Value = percentchange
                
                Else
                
                    ws.Range("J" & ticker_row).Value = "N/A"
                
                End If
                
            ' print and reset total
                
                ws.Range("K" & ticker_row).Value = volumetotal
                
                volumetotal = 0
                
            ' move to next row
            
                ticker_row = ticker_row + 1
                
            ' change to next open price
            
                openprice = ws.Cells(i + 1, 3).Value
            
                
            End If
    
        Next i
    
    
    Next ws

End Sub

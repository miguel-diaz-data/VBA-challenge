Attribute VB_Name = "Module1"
Sub MediumSolution():
    
    'loop through all stocks in wb and add up all the stock volume of each ticker
    
    Dim TickerIndex As Integer
    Dim TotalVolume As Double
    Dim OpeningPrice As Double
    Dim StocksInSheet As Integer
            
    For Each ws In Worksheets
    'loop through each ws
        ws.Activate
        
        TotalVolume = 0
        TickerIndex = 2
        'stores price of first stock in sheet
        OpeningPrice = Range("C2").Value
        
        For i = 2 To 73000
        
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                
                If (Cells(i, 1).Value = Null) Then
                    Exit For
                '
                End If
                                
                'assuming the sheet is sorted by ticker,then if current ticker is the last one, put totalvolume on to summary table
                Cells(TickerIndex, 9).Value = Cells(i, 1).Value
                Cells(TickerIndex, 10).Value = Cells(i, 6).Value - OpeningPrice
                
                If (OpeningPrice = 0) Then
                    Cells(TickerIndex, 11).Value = FormatPercent(0)
                    Else
                    Cells(TickerIndex, 11).Value = FormatPercent(Cells(TickerIndex, 10).Value / OpeningPrice)
                'if statement to avoid division by zero errors when opening price remains at zero
                
                End If
                
                Cells(TickerIndex, 12).Value = TotalVolume
                
                'highlights positive and negative changes of stock price via green / red highlighting
                If Cells(TickerIndex, 10).Value > 0 Then
                    Cells(TickerIndex, 10).Interior.ColorIndex = 4
                ElseIf Cells(TickerIndex, 10).Value < 0 Then
                    Cells(TickerIndex, 10).Interior.ColorIndex = 3
                
                End If
                
                'prepared variables for next iteration with a new ticker
                TotalVolume = 0
                TickerIndex = TickerIndex + 1
                OpeningPrice = Cells(i + 1, 3).Value
                
                                
            End If
        
        Next i
        
        
    
    Next ws
                
End Sub

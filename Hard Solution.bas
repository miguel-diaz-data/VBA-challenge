Attribute VB_Name = "Module1"
Sub HardSolution():
    
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
        
        
        For i = 2 To 1000000
        
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                
                'exits for loop when the next cell has no data, aka reached the bottom of sheet
                If (Cells(i, 1).Value = Null) Then
                    Exit For
                
                End If
                                
                'assuming the sheet is sorted by ticker,then if current ticker is the last one, put totalvolume on to summary table
                Cells(TickerIndex, 9).Value = Cells(i, 1).Value
                Cells(TickerIndex, 10).Value = Cells(i, 6).Value - OpeningPrice
                
                'if statement to avoid division by zero errors when opening price remains at zero
                If (OpeningPrice = 0) Then
                    Cells(TickerIndex, 11).Value = FormatPercent(0)
                    Else
                    Cells(TickerIndex, 11).Value = FormatPercent(Cells(TickerIndex, 10).Value / OpeningPrice)
                            
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
                StocksInSheet = StocksInSheet + 1
                                
            End If
        
        Next i
        'names rows and columns for the table values generated
        Range("I1") = "Ticker"
        Range("J1") = "Yearly Change"
        Range("K1") = "Percent Change"
        Range("L1") = "Total Volume"
        
        Range("O2") = "Greatest % Inc"
        Range("O3") = "Greatest % Dec"
        Range("O4") = "Greatest Total Vol"
        Range("P1") = "Ticker"
        Range("Q1") = "Value"
        
        For j = 2 To (StocksInSheet + 1)
            'go through 1st table to generate values for the second table
            If Range("Q3").Value > Cells(j, 11) Then
                Range("Q3").Value = Cells(j, 11)
                Range("P3").Value = Cells(j, 9)
            ElseIf Range("Q2").Value < Cells(j, 11) Then
                Range("Q2").Value = Cells(j, 11)
                Range("P2").Value = Cells(j, 9)
            End If
            If Range("Q4").Value < Cells(j, 12) Then
                Range("Q4").Value = Cells(j, 12)
                Range("P4").Value = Cells(j, 9)
            End If
        Next j
        
        Range("Q2:Q3").NumberFormat = "0.00%"
    Next ws
                
End Sub

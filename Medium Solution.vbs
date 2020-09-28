Attribute VB_Name = "Module1"
Sub EasySolution():
    
    'loop through all stocks in wb and add up all the stock volume of each ticker
    
    
    'Dim LastRow As Integer
    'assuming number of rows is known for now
    
    Dim TickerIndex As Integer
    Dim TotalVolume As Double
    Dim OpeningPrice As Integer
    
    
    'For Each ws In Worksheets
    'loop through each ws will be done after it can be done for a single sheet
        TotalVolume = 0
        TickerIndex = 2
        OpeningPrice = Range("C2").Value
        'stores price of first stock in sheet
        
        For i = 2 To 70926
        
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                'assuming the sheet is sorted by ticker,then if current ticker is the last one, put totalvolume on to summary table
                Cells(TickerIndex, 9).Value = Cells(i, 1).Value
                Cells(TickerIndex, 10).Value = Cells(i, 6).Value - OpeningPrice
                Cells(TickerIndex, 11).Value = Cells(TickerIndex, 10).Value / OpeningPrice
                Cells(TickerIndex, 12).Value = TotalVolume
                
                'prepared variables for next iteration with a new ticker
                TotalVolume = 0
                TickerIndex = TickerIndex + 1
                OpeningPrice = Cells(i + 1, 3).Value
                
                                
                
            End If
        
        Next i
                
End Sub


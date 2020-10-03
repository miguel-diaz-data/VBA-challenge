Attribute VB_Name = "Module11"
Sub EasySolution():
    
    'loop through all stocks in wb and add up all the stock volume of each ticker
    
    
    'Dim LastRow As Integer
    'assuming number of rows is known for now
    
    Dim TickerIndex As Integer
    Dim TotalVolume As Double
    
   
        
    'For Each ws In Worksheets
    'loop through each ws will be done after it can be done for a single sheet
        TotalVolume = 0
        TickerIndex = 2
        
        For i = 2 To 70926
        
            TotalVolume = TotalVolume + Cells(i, 7).Value
            
            If Cells(i, 1).Value <> Cells(i + 1, 1) Then
                'assuming the sheet is sorted by ticker,then if current ticker is the last one, put totalvolume on to summary table
                Cells(TickerIndex, 9).Value = Cells(i, 1).Value
                Cells(TickerIndex, 10).Value = TotalVolume
                
                TotalVolume = 0
                TickerIndex = TickerIndex + 1
                
            End If
        
        Next i
    
                
                
                
                
            
            
            
            
            
            
            
        
        
    
End Sub


Sub Ticker()
    
    'Set All Variables
    
    'Ticker and Summary Table
    Dim Ticker As String
    Dim SummaryTableRow As Integer
    
    'ClosePrice, OpenPrice, Yearly Change
    Dim ClosePrice As Double
    Dim OpenPrice As Double
    Dim YearlyChange As Double
    
    'PercentChange, Total Volumne, Max Percent Increase, Decrease, Volume
    Dim PercentChange As Double
    Dim TotalVolume As Double
    Dim MaxPercentIncrease As Double
    Dim MaxPercentDecrease As Double
    Dim MaxVolume As Double
    
    'Max Increase, Decrease, Volume Tickers
    Dim MaxPercentIncreaseTicker As String
    Dim MaxPercentDecreaseTicker As String
    Dim MaxVolumeTicker As String

    'Looping for all worksheets`
    For Each ws In Worksheets
        
        'Create new column names
        ws.Range("I" & 1).Value = "Ticker"
        ws.Range("J" & 1).Value = "Yearly Change"
        ws.Range("K" & 1).Value = "Percent Change"
        ws.Range("L" & 1).Value = "Total Stock Volume"
        ws.Range("O" & 2).Value = "Greatest % Increase"
        ws.Range("O" & 3).Value = "Greatest % Decrease"
        ws.Range("O" & 4).Value = "Greatest Total Volume"
        ws.Range("P" & 1).Value = "Ticker"
        ws.Range("Q" & 1).Value = "Value"
    
        'Set SummaryTableRow value
        SummaryTableRow = 2
        
        'Set YearlyChange, Open and Close Price variable values
        YearlyChange = 0
        OpenPrice = 0
        ClosePrice = 0
        
        'Set Percent Change and Total Stock Volume variable values
        PercentChange = 0
        TotalVolume = 0
        
        'Determine Last Row
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        'Set Max Percent changes and Stock Volume values
        MaxPercentIncrease = 0
        MaxPercentDecrease = 0
        MaxVolume = 0
    
        'Start main loop
        For i = 2 To LastRow
            TotalVolume = TotalVolume + ws.Cells(i, 7).Value
            
            'First row of a ticker
            If ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1).Value Or OpenPrice = 0 Then
                OpenPrice = ws.Cells(i, 3).Value
            End If
        
            'Searches for when value of next cell is different than that of current cell
            'Last row of a ticker
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
                Ticker = ws.Cells(i, 1).Value
                ws.Range("I" & SummaryTableRow).Value = Ticker
                
                ClosePrice = ws.Cells(i, 6).Value
                
                YearlyChange = ClosePrice - OpenPrice
                ws.Range("J" & SummaryTableRow).Value = YearlyChange
                
                If OpenPrice > 0 Then
                    PercentChange = (YearlyChange * 100#) / OpenPrice
                Else
                    PercentChange = 0
                End If
                ws.Range("K" & SummaryTableRow).Value = PercentChange
                    
                'Color Coding Yearly Change
                If YearlyChange > 0 Then
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 4
                Else
                    ws.Range("J" & SummaryTableRow).Interior.ColorIndex = 3
                End If
           
                'Calculating Max Percent Increase and Decrease
                If PercentChange > MaxPercentIncrease Then
                    MaxPercentIncrease = PercentChange
                    MaxPercentIncreaseTicker = Ticker
                End If
                
                If PercentChange < MaxPercentDecrease Then
                    MaxPercentDecrease = PercentChange
                    MaxPercentDecreaseTicker = Ticker
                End If
           
                ws.Range("L" & SummaryTableRow).Value = TotalVolume
       
                'Calculating Max Volume
                If TotalVolume > MaxVolume Then
                    MaxVolume = TotalVolume
                    MaxVolumeTicker = Ticker
                End If
            
                TotalVolume = 0
                
                SummaryTableRow = SummaryTableRow + 1
            End If
            
        Next i
        
        'Isolate max increase, decrease, and total volume and tickers in new columns
        ws.Range("P" & 2).Value = MaxPercentIncreaseTicker
        ws.Range("P" & 3).Value = MaxPercentDecreaseTicker
        ws.Range("P" & 4).Value = MaxVolumeTicker
        ws.Range("Q" & 2).Value = MaxPercentIncrease
        ws.Range("Q" & 3).Value = MaxPercentDecrease
        ws.Range("Q" & 4).Value = MaxVolume
        
    Next ws
    

End Sub




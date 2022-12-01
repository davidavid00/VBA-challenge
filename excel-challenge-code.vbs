Attribute VB_Name = "Module1"
Sub ticker_summary()
    'declaring empty variables that will be used later in the code
    Dim tickerSymbol As String
    Dim firstOpen As Double
    Dim changeTotal As Double
    Dim pChange As Double
    Dim volumeTotal As Double
    Dim summaryCounter As Integer
    Dim lastrow As Double
    
    'bonus section variables
    Dim greatIncrease As Double
    Dim greatDecrease As Double
    Dim greatTotal As Double
    


    'script to run on each worksheet
    For Each ws In Worksheets
        
        'inital configuration of the variables
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        firstOpen = ws.Cells(2, 3).Value
        summaryCounter = 2
        
        'Create headers in the sheet
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        'Format  percentage column
        ws.Range("K2:K" & lastrow).NumberFormat = "0.00%"
        
        'for rows 2 onward, perform this
        For r = 2 To lastrow
            'if a new ticker is detected in the cell below
            If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
                'fetch current ticker
                tickerSymbol = ws.Cells(r, 1).Value
                'calculate the change
                changeTotal = ws.Cells(r, 6) - firstOpen
                'calculat the percentage of change
                pChange = changeTotal / firstOpen
            
                'format the cells green if positive or red if equal or negative
                If changeTotal > 0 Then
                    ws.Cells(summaryCounter, 10).Interior.Color = RGB(153, 255, 153)
                ElseIf changeTotal <= 0 Then
                    ws.Cells(summaryCounter, 10).Interior.Color = RGB(255, 102, 102)
                End If
                
                'place all values in the appropriate columns
                ws.Cells(summaryCounter, 9).Value = tickerSymbol
                ws.Cells(summaryCounter, 10).Value = changeTotal
                ws.Cells(summaryCounter, 11).Value = pChange
                ws.Cells(summaryCounter, 12).Value = volumeTotal
            
                'reset variables for new ticker
                volumeTotal = 0
                firstOpen = ws.Cells(r + 1, 3).Value
                summaryCounter = summaryCounter + 1
            Else
            'if it is not a new ticker, add to the volume
            volumeTotal = volumeTotal + ws.Cells(r, 7)
            End If
        'next row
        Next r
    
        lastTotal = ws.Cells(Rows.Count, 9).End(xlUp).Row
        greatIncrease = 0
        greatDecrease = 0
        greatVolume = 0

        For tR = 2 To lastTotal
            If ws.Cells(tR + 1, 11).Value > greatIncrease Then
                greatIncrease = ws.Cells(tR + 1, 11).Value
                greatIncreaseN = ws.Cells(tR + 1, 9).Value
            ElseIf ws.Cells(tR + 1, 11).Value < greatDecrease Then
                greatDecrease = ws.Cells(tR + 1, 11).Value
                greatDecreaseN = ws.Cells(tR + 1, 9).Value
            End If
            If ws.Cells(tR + 1, 12).Value > greatVolume Then
                greatVolume = ws.Cells(tR + 1, 12).Value
                greatVolumeN = ws.Cells(tR + 1, 9).Value
            End If
        Next tR
        
        ws.Cells(2, 15).Value = "Greatest % Increase"
        ws.Cells(3, 15).Value = "Greatest % Decrease"
        ws.Cells(4, 15).Value = "Greatest Total Volume"
        ws.Cells(1, 16).Value = "Ticker"
        ws.Cells(1, 17).Value = "Value"
        ws.Cells(2, 16).Value = greatIncreaseN
        ws.Cells(2, 17).Value = greatIncrease
        ws.Cells(3, 16).Value = greatDecreaseN
        ws.Cells(3, 17).Value = greatDecrease
        ws.Cells(4, 16).Value = greatVolumeN
        ws.Cells(4, 17).Value = greatVolume
        ws.Range("Q2:Q3").NumberFormat = "0.00%"
        
    'proceed to next worksheet
    Next ws

End Sub


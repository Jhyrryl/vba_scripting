Sub process()
    Dim ticker As String
    Dim nextTicker As String
    Dim totalVolume As LongLong
    Dim startValue As Double
    Dim endValue As Double
    Dim yearlyChange As Double
    Dim percChange As Double
    Dim greatestPercIncreaseTicker As String
    Dim greatestPercDecreaseTicker As String
    Dim greatestTotalVolumeTicker As String
    Dim greatestPercIncreaseValue As Double
    Dim greatestPercDecreaseValue As Double
    Dim greatestTotalVolumeValue As Double
    
    For Each ws In Worksheets
        MsgBox ("Processing Sheet: " + ws.Name)
        
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percentage Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        nextRecord = 1
        
        ticker = ""
        totalVolume = 0
        startValue = 0
        endValue = 0
        
        greatestPercIncreaseValue = 0
        greatestPercDecreaseValue = 0
        greatestTotalVolumeValue = 0
        
        For i = 2 To lastrow
            nextTicker = ws.Cells(i, 1)
            If ticker = "" Then
                ticker = nextTicker
                startValue = ws.Cells(i, 3).Value
                totalVolume = ws.Cells(i, 7)
            ElseIf nextTicker = ticker Then
                totalVolume = totalVolume + ws.Cells(i, 7)
            Else
                ' calc ticker results
                prevRow = i - 1
                endValue = ws.Cells(prevRow, 6).Value
                yearlyChange = startValue - endValue
                If yearlyChange = 0 Then
                    percChange = 0
                ElseIf startValue = 0 Then
                    If yearlyChange > 0 Then
                        percChange = 100
                    ElseIf yearlyChange < 0 Then
                        percChange = -100
                    Else
                        percChange = 0
                    End If
                Else
                    percChange = yearlyChange / startValue
                End If
                
                ' report ticker
                nextRecord = nextRecord + 1
                ws.Cells(nextRecord, 9).Value = ticker
                ws.Cells(nextRecord, 10).Value = yearlyChange
                ws.Cells(nextRecord, 11).Value = percChange
                ws.Cells(nextRecord, 12).Value = totalVolume
                
                ' highlight
                If yearlyChange > 0 Then
                    ws.Cells(nextRecord, 10).Interior.ColorIndex = 4
                    ws.Cells(nextRecord, 11).Interior.ColorIndex = 4
                ElseIf yearlyChange < 0 Then
                    ws.Cells(nextRecord, 10).Interior.ColorIndex = 3
                    ws.Cells(nextRecord, 11).Interior.ColorIndex = 3
                End If
                
                ' update greatest values
                If percChange > 0 Then
                    If percChange > greatestPercIncreaseValue Then
                        greatestPercIncreaseValue = percChange
                        greatestPercIncreaseTicker = ticker
                    End If
                End If
                    
                If percChange < 0 Then
                    If percChange < greatestPercDecreaseValue Then
                        greatestPercDecreaseValue = percChange
                        greatestPercDecreaseTicker = ticker
                    End If
                End If
                    
                If totalVolume > greatestTotalVolumeValue Then
                    greatestTotalVolumeValue = totalVolume
                    greatestTotalVolumeTicker = ticker
                End If
                
                ' init next ticker
                ticker = nextTicker
                startValue = ws.Cells(i, 3).Value
                totalVolume = 0
            
            End If
            
        Next i
        
        ' done - report greatest results
        ws.Range("P2") = greatestPercIncreaseTicker
        ws.Range("P3") = greatestPercDecreaseTicker
        ws.Range("P4") = greatestTotalVolumeTicker
        ws.Range("Q2") = greatestPercIncreaseValue
        ws.Range("Q3") = greatestPercDecreaseValue
        ws.Range("Q4") = greatestTotalVolumeValue
        
        ws.Range("P2:Q2").Interior.ColorIndex = 4
        ws.Range("P3:Q3").Interior.ColorIndex = 3
        
    Next ws
    
    MsgBox ("Finished!")
            
End Sub

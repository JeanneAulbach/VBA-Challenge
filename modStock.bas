Attribute VB_Name = "modStock"
Sub SummarizeData()
    Dim Ticker As String
    Dim BeginningPrice As Currency
    Dim EndingPrice As Currency
    Dim PriceChange As Currency
    Dim PercentChange As Single
    Dim Volume As Single
    Dim NumRows As Long
    Dim NumColumns As Long
    Dim PctIncreaseTicker As String
    Dim PctDecreaseTicker As String
    Dim PctIncreaseValue As Double
    Dim PctDecreaseValue As Double
    Dim GreatestVolumeTicker As String
    Dim GreatestVolumeValue As Single
    Dim i As Long
    Dim j As Integer
    
    For Each ws In Worksheets
    
        NumRows = Range("A2", Range("A2").End(xlDown)).Rows.Count
        Ticker = ws.Cells(2, 1).Value
        Volume = ws.Cells(2, 7).Value
        BeginningPrice = ws.Cells(2, 3).Value
        j = 2
        EndingPrice = 0
        PriceChange = 0
        PercentChange = 0
        PctIncreaseTicker = ""
        PctDecreaseTicker = ""
        GreatestVolumeTicker = ""
        PctIncreaseValue = 0
        PctDecreaseValue = 0
        GreatestVolumeValue = 0
        
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Price Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        
        For i = 2 To NumRows
            If ws.Cells(i, 1).Value = Ticker Then
                Volume = Volume + ws.Cells(i, 7).Value
                GreatestVolumeValue = GreatestVolumeValue + Volume
            Else
                EndingPrice = ws.Cells(i - 1, 6).Value
                ws.Cells(j, 9) = Ticker
                With ws.Cells(j, 10)
                    .Value = EndingPrice - BeginningPrice
                    If EndingPrice - BeginningPrice < 0 Then
                        .Interior.ColorIndex = 3
                    Else
                        .Interior.ColorIndex = 4
                    End If
                End With
                If BeginningPrice > 0 Then
                    ws.Cells(j, 11) = (EndingPrice - BeginningPrice) / BeginningPrice
                Else
                    ws.Cells(j, 11) = 0
                End If
                ws.Cells(j, 12).Value = Volume
                
                If PctIncreaseTicker = "" Then
                    PctIncreaseTicker = Ticker
                    PctDecreaseTicker = Ticker
                    GreatestVolumeTicker = Ticker
                    PctIncreaseValue = ws.Cells(j, 11).Value
                    PctIncreaseValue = ws.Cells(j, 11).Value
                    GreatestVolumeValue = Volume
                Else
                    If PctIncreaseValue < ws.Cells(j, 11).Value Then
                        PctIncreaseTicker = Ticker
                        PctIncreaseValue = ws.Cells(j, 11).Value
                    ElseIf PctDecreaseValue > ws.Cells(j, 11).Value Then
                        PctDecreaseTicker = Ticker
                        PctDecreaseValue = ws.Cells(j, 11).Value
                    End If
                End If
                
                If GreatestVolumeValue > Volume Then
                    GreatestVolumeTicker = Ticker
                    GreatestVolumeValue = Volume
                End If
        
                        
                j = j + 1
                Ticker = ws.Cells(i, 1).Value
                BeginningPrice = ws.Cells(i, 3).Value
                Volume = ws.Cells(i, 7).Value
                    
            End If
                
        Next i
        
        ws.Cells(2, 15).Value = PctIncreaseTicker
        ws.Cells(2, 16).Value = PctIncreaseValue
        ws.Cells(3, 15).Value = PctDecreaseTicker
        ws.Cells(3, 16).Value = PctDecreaseValue
        ws.Cells(4, 15).Value = GreatestVolumeTicker
        ws.Cells(4, 16).Value = GreatestVolumeValue
        
     Next ws
    

End Sub



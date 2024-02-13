Sub AnalyzeWorkbookByYear()
    Dim ws As Worksheet
    Dim symbol As String
    Dim beginValue As Double
    Dim endValue As Double
    Dim currentRow As Long
    Dim prevRow As String
    Dim stockTotal As Double
    Dim count As Long
    Dim analysisRow As Long
    Dim rng As Range
    Dim totalVolume As Double
    
    ' Make sure the loop goes through each sheet in the workbook
    For Each ws In ThisWorkbook.Sheets
        
        ' Set up the first table
        analysisRow = 2
        Dim wsRange As Range
        Set wsRange = ws.Range("I1")
        
        ' Headers for the first table
        With wsRange
            .Offset(0, 0).Value = "Symbol"
            .Offset(0, 1).Value = "Yearly Change"
            .Offset(0, 2).Value = "Percent Change"
            .Offset(0, 3).Value = "Total Volume"
        End With
     
        prevRow = ws.Cells(2, "A").Value
        stockTotal = 0
        count = 0
        totalVolume = 0
        
        ' Arrays to store data for the second analysis table
        Dim maxIncrease As Double
        Dim maxDecrease As Double
        Dim maxVolume As Double
        Dim symbolIncrease As String
        Dim symbolDecrease As String
        Dim symbolVolume As String
        
        ' Nest the loop for each worksheet
        For currentRow = 2 To ws.Cells(ws.Rows.count, "A").End(xlUp).Row
            symbol = ws.Cells(currentRow, "A").Value
            beginValue = ws.Cells(currentRow, "C").Value
            endValue = ws.Cells(currentRow, "F").Value
   
            Dim yearlyChange As Double
            Dim percentChange As Double
            
            yearlyChange = endValue - beginValue
            If beginValue <> 0 Then
                percentChange = (endValue - beginValue) / beginValue * 100
            Else
                percentChange = 0
            End If
            
            ' Get the total values for each range
            If symbol = prevRow Then
                stockTotal = stockTotal + yearlyChange
                count = count + 1
                totalVolume = totalVolume + ws.Cells(currentRow, "G").Value
            ElseIf symbol <> prevRow Then
                ' Store the total values before moving to the next symbol in the first analysis table
                With wsRange
                    .Offset(analysisRow, 0).Value = prevRow
                    .Offset(analysisRow, 1).Value = stockTotal
                    .Offset(analysisRow, 2).Value = percentChange
                    .Offset(analysisRow, 3).Value = totalVolume
                End With
                
'                ' Formatting for red and green
               Set rng = wsRange.Offset(analysisRow, 1)

                If Left(rng.Value, 1) = "-" Then
                 rng.Interior.Color = RGB(255, 0, 0)
                Else
                  rng.Interior.Color = RGB(0, 255, 0)
                End If

                ' Maximums for the second table
                If percentChange > maxIncrease Then
                    maxIncrease = percentChange
                    symbolIncrease = prevRow
                End If
                If percentChange < maxDecrease Then
                    maxDecrease = percentChange
                    symbolDecrease = prevRow
                End If
                If totalVolume > maxVolume Then
                    maxVolume = totalVolume
                    symbolVolume = prevRow
                End If
                
                ' Reset variables for the next symbol
                analysisRow = analysisRow + 1
                stockTotal = yearlyChange
                count = 1
                totalVolume = ws.Cells(currentRow, "G").Value
                prevRow = symbol
            End If
        Next currentRow
        
        ' Store the final values for the last symbol in the first analysis table
        With wsRange
            .Offset(analysisRow, 0).Value = prevRow
            .Offset(analysisRow, 1).Value = stockTotal
            .Offset(analysisRow, 2).Value = percentChange
            .Offset(analysisRow, 3).Value = totalVolume
        End With
        
        ' Set up second analysis table
        Dim wsRange2 As Range
        Set wsRange2 = ws.Range("N1")
        
        ' fill second analysis table
        With wsRange2
            .Offset(0, 0).Value = "Metric"
            .Offset(0, 1).Value = "Symbol"
            .Offset(0, 2).Value = "Value"
            .Offset(1, 0).Value = "Greatest Increase"
            .Offset(1, 1).Value = symbolIncrease
            .Offset(1, 2).Value = maxIncrease
            .Offset(2, 0).Value = "Greatest Decrease"
            .Offset(2, 1).Value = symbolDecrease
            .Offset(2, 2).Value = maxDecrease
            .Offset(3, 0).Value = "Greatest Total Volume"
            .Offset(3, 1).Value = symbolVolume
            .Offset(3, 2).Value = maxVolume
        End With
        
    Next ws
    
    ' Let them know it's done
    MsgBox "Analysis tables have been populated."
End Sub


